package com.filemapping.aix;

import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.function.IntFunction;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/** Streaming workbook input and values-only report output. */
public final class WorkbookIO {
    private static final SpreadsheetVersion XLSX = SpreadsheetVersion.EXCEL2007;
    private static final long MILLIS_PER_DAY = 86_400_000L;

    private WorkbookIO() {
    }

    @FunctionalInterface
    public interface RowConsumer {
        /** Row numbers start at 1. Column positions in cells start at 0. */
        void accept(int rowNumber, List<String> cells) throws Exception;
    }

    public static void read(Path input, String sheetName, RowConsumer consumer) throws Exception {
        try (OPCPackage workbook = open(input)) {
            WorkbookReader reader = new WorkbookReader(workbook);
            reader.readSelected(sheetName, consumer);
        }
    }

    public static void writeResult(Path input, String sheetName, Path output,
                                   IntFunction<String[]> outcome) throws Exception {
        if (input.toAbsolutePath().normalize().equals(output.toAbsolutePath().normalize())
                || (Files.exists(output) && Files.isSameFile(input, output))) {
            throw new IllegalArgumentException("Result workbook must not overwrite the input workbook");
        }

        try (OPCPackage workbook = open(input)) {
            WorkbookReader reader = new WorkbookReader(workbook);
            ReportLayout layout = new ReportLayout();
            String selectedSheet = reader.readSelected(sheetName, layout::observe);
            if (layout.columnCount > XLSX.getMaxColumns() - 2) {
                throw new IllegalArgumentException("Worksheet has no room for two result columns: " + selectedSheet);
            }

            try (SXSSFWorkbook report = new SXSSFWorkbook(100)) {
                report.setCompressTempFiles(true);
                XSSFReader.SheetIterator sheets = reader.sheets();
                while (sheets.hasNext()) {
                    try (InputStream data = sheets.next()) {
                        String name = sheets.getSheetName();
                        Sheet reportSheet = report.createSheet(name);
                        if (name.equals(selectedSheet)) {
                            writeSelectedSheet(reader, data, reportSheet, layout, outcome);
                        } else {
                            reader.parse(data, (rowNumber, cells) ->
                                    writeCells(reportSheet.createRow(rowNumber - 1), cells));
                        }
                    }
                }
                try (OutputStream stream = Files.newOutputStream(output)) {
                    report.write(stream);
                }
            }
        }
    }

    private static OPCPackage open(Path input) throws Exception {
        if (!input.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
            throw new IllegalArgumentException("Only .xlsx workbooks are supported: " + input);
        }
        return OPCPackage.open(input.toFile(), PackageAccess.READ);
    }

    private static void writeSelectedSheet(WorkbookReader reader, InputStream data, Sheet sheet,
                                           ReportLayout layout, IntFunction<String[]> outcome) throws Exception {
        Row header = sheet.createRow(0);
        setText(header, layout.columnCount, layout.uniqueHeader("Mapping Status"));
        setText(header, layout.columnCount + 1, layout.uniqueHeader("Failure Reason"));
        sheet.createFreezePane(0, 1);
        reader.parse(data, (rowNumber, cells) -> {
            if (rowNumber == 1) {
                writeCells(header, cells);
                return;
            }
            Row row = sheet.createRow(rowNumber - 1);
            writeCells(row, cells);
            String[] status = outcome.apply(rowNumber);
            if (status != null) {
                setText(row, layout.columnCount, status[0]);
                setText(row, layout.columnCount + 1, status[1]);
            }
        });
    }

    private static void writeCells(Row row, List<String> cells) {
        for (int column = 0; column < cells.size(); column++) {
            setText(row, column, cells.get(column));
        }
    }

    private static void setText(Row row, int column, String text) {
        if (text != null && !text.isEmpty()) {
            int maxLength = XLSX.getMaxTextLength();
            if (text.length() > maxLength) {
                int end = maxLength - 1;
                if (Character.isHighSurrogate(text.charAt(end - 1))) {
                    end--;
                }
                text = text.substring(0, end) + "\u2026";
            }
            row.createCell(column).setCellValue(text);
        }
    }

    private static final class ReportLayout {
        private final Set<String> headers = new HashSet<>();
        private int columnCount;

        void observe(int rowNumber, List<String> cells) {
            columnCount = Math.max(columnCount, cells.size());
            if (rowNumber == 1) {
                for (String cell : cells) {
                    headers.add(cell.trim().toLowerCase(Locale.ROOT));
                }
            }
        }

        String uniqueHeader(String base) {
            String candidate = base;
            for (int suffix = 2; !headers.add(candidate.toLowerCase(Locale.ROOT)); suffix++) {
                candidate = base + " " + suffix;
            }
            return candidate;
        }
    }

    private static final class WorkbookReader {
        private final XSSFReader reader;
        private final Styles styles;
        private final SharedStrings strings;
        private final DataFormatter formatter;

        WorkbookReader(OPCPackage workbook) throws Exception {
            reader = new XSSFReader(workbook);
            reader.setUseReadOnlySharedStringsTable(true);
            styles = reader.getStylesTable();
            strings = reader.getSharedStringsTable();
            formatter = new WorkbookFormatter(uses1904Dates(reader));
        }

        XSSFReader.SheetIterator sheets() throws Exception {
            return (XSSFReader.SheetIterator) reader.getSheetsData();
        }

        String readSelected(String sheetName, RowConsumer consumer) throws Exception {
            XSSFReader.SheetIterator sheets = sheets();
            while (sheets.hasNext()) {
                try (InputStream data = sheets.next()) {
                    String name = sheets.getSheetName();
                    if (sheetName == null || sheetName.equals(name)) {
                        parse(data, consumer);
                        return name;
                    }
                }
            }
            throw new IllegalArgumentException(sheetName == null ? "Workbook has no worksheet"
                    : "Worksheet not found: " + sheetName);
        }

        void parse(InputStream data, RowConsumer consumer) throws Exception {
            XMLReader parser = XMLHelper.newXMLReader();
            parser.setContentHandler(new XSSFSheetXMLHandler(styles, strings,
                    new SheetRows(consumer), formatter, false));
            try {
                parser.parse(new InputSource(data));
            } catch (RowConsumerException error) {
                throw error.getOriginal();
            }
        }
    }

    private static boolean uses1904Dates(XSSFReader workbook) throws Exception {
        boolean[] date1904 = {false};
        XMLReader parser = XMLHelper.newXMLReader();
        parser.setContentHandler(new DefaultHandler() {
            @Override
            public void startElement(String uri, String localName, String qName, Attributes attributes) {
                if ("workbookPr".equals(localName)) {
                    String value = attributes.getValue("date1904");
                    date1904[0] = "1".equals(value) || "true".equalsIgnoreCase(value);
                }
            }
        });
        try (InputStream properties = workbook.getWorkbookData()) {
            parser.parse(new InputSource(properties));
        }
        return date1904[0];
    }

    private static final class WorkbookFormatter extends DataFormatter {
        private final boolean date1904;

        WorkbookFormatter(boolean date1904) {
            super(Locale.ROOT);
            this.date1904 = date1904;
        }

        @Override
        public String formatRawCellContents(double value, int formatIndex, String formatString) {
            if (!Double.isFinite(value)) {
                return Double.toString(value);
            }
            if (value >= 1 && value < 2_958_466 && DateUtil.isADateFormat(formatIndex, formatString)) {
                double day = Math.floor(value);
                LocalDateTime midnight = DateUtil.getLocalDateTime(day, date1904);
                // Round display precision without moving a late timestamp into the next day.
                long millis = Math.min(MILLIS_PER_DAY - 1, Math.round((value - day) * MILLIS_PER_DAY));
                if (midnight.getYear() <= 9999) {
                    return midnight.plusNanos(millis * 1_000_000L).format(DateTimeFormatter.ISO_LOCAL_DATE_TIME);
                }
            }
            if (formatString.equalsIgnoreCase("General")) {
                return BigDecimal.valueOf(value).stripTrailingZeros().toPlainString();
            }
            return super.formatRawCellContents(value, formatIndex, formatString, date1904);
        }
    }

    private static final class SheetRows implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final RowConsumer consumer;
        private List<String> cells;
        private int nextColumn;

        SheetRows(RowConsumer consumer) {
            this.consumer = consumer;
        }

        @Override
        public void startRow(int index) {
            cells = new ArrayList<>();
            nextColumn = 0;
        }

        @Override
        public void cell(String reference, String value, XSSFComment comment) {
            int column = reference == null ? nextColumn : new CellReference(reference).getCol();
            if (column < 0 || column >= XLSX.getMaxColumns()) {
                throw new IllegalArgumentException("Invalid column reference: " + reference);
            }
            nextColumn = column + 1;
            if (value != null && !value.isEmpty()) {
                while (cells.size() <= column) {
                    cells.add("");
                }
                cells.set(column, value);
            }
        }

        @Override
        public void endRow(int index) {
            if (index < 0 || index >= XLSX.getMaxRows()) {
                throw new IllegalArgumentException("Invalid worksheet row: " + (index + 1));
            }
            try {
                consumer.accept(index + 1, Collections.unmodifiableList(cells));
            } catch (Exception error) {
                throw new RowConsumerException(error);
            }
        }
    }

    private static final class RowConsumerException extends RuntimeException {
        private static final long serialVersionUID = 1L;

        RowConsumerException(Exception cause) {
            super(cause);
        }

        Exception getOriginal() {
            return (Exception) getCause();
        }
    }
}
