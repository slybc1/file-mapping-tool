package com.filemapping.aix;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

public class FileMappingCliTest {
    @Rule
    public final TemporaryFolder temp = new TemporaryFolder();

    @Test
    public void previewsAndThenMovesGroupedFilesFromDeepNestedDirectories() throws IOException {
        Path source = temp.newFolder("source").toPath();
        Path firstReport = write(source.resolve("north/archive/old/layer4/Report.pdf"), "first");
        Path secondReport = write(source.resolve("south/backup/Report.pdf"), "second");
        Path notes = write(source.resolve("north/archive/old/layer4/Notes.txt"), "notes");
        Path caseMismatch = write(source.resolve("west/more/report.PDF"), "not an exact match");
        Path excel = temp.getRoot().toPath().resolve("input.xlsx");
        writeWorkbook(excel);

        Path target = temp.newFolder("target").toPath();
        Path wo2025 = target.resolve("2025/WO/77");
        Path pm2025 = target.resolve("2025/PM/77");
        Path wo2026 = target.resolve("2026/WO/77");
        Path existing = write(wo2025.resolve("Report.pdf"), "existing");
        Path result = temp.getRoot().toPath().resolve("result.xlsx");
        List<String> args = Arrays.asList(
            "--excel", excel.toString(), "--source", source.toString(),
            "--target", target.toString(), "--sheet", "Export",
            "--file-column", "Attachment", "--id-column", "CTX_ID",
            "--result", result.toString()
        );

        assertExitCode(args, 0);
        assertTrue(Files.isRegularFile(firstReport));
        assertTrue(Files.isRegularFile(secondReport));
        assertTrue(Files.isRegularFile(notes));
        assertFalse(Files.exists(result));
        assertFalse(Files.exists(pm2025));
        assertFalse(Files.exists(wo2026));
        assertEquals(set("Report.pdf"), fileNames(wo2025));

        List<String> apply = new ArrayList<>(args);
        apply.add("--apply");
        assertExitCode(apply, 0);

        assertEquals("existing", contents(existing));
        assertEquals(set("Report.pdf", "Report_2.pdf", "Report_3.pdf", "Notes.txt"), fileNames(wo2025));
        assertEquals(set("first", "second"), reportContents(wo2025, "Report_2.pdf", "Report_3.pdf"));
        for (Path folder : Arrays.asList(pm2025, wo2026)) {
            assertEquals(set("Report.pdf", "Report_2.pdf"), fileNames(folder));
            assertEquals(set("first", "second"), reportContents(folder, "Report.pdf", "Report_2.pdf"));
        }
        assertEquals("notes", contents(wo2025.resolve("Notes.txt")));
        assertFalse(Files.exists(firstReport));
        assertFalse(Files.exists(secondReport));
        assertFalse(Files.exists(notes));
        assertEquals("not an exact match", contents(caseMismatch));
        assertEquals(Arrays.asList("Mapped", "Mapped", "Mapped", "Mapped", "Mapped", "Not found"),
            rowStatuses(result));
    }

    @Test
    public void prefersStoredFileNameAndValidatesDb2Timestamps() throws IOException {
        Path root = temp.getRoot().toPath();
        Path source = temp.newFolder("source").toPath();
        Path stored = write(source.resolve("stored.csv"), "saved file");
        Path generic = write(source.resolve("generic.csv"), "alternate file");
        Path invalid = write(source.resolve("invalid.csv"), "invalid date");
        Path excel = root.resolve("db2.xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream output = Files.newOutputStream(excel)) {
            Sheet sheet = workbook.createSheet("Export");
            addRow(sheet, 0, "FILE_NAME", "STORED_FILE_NAME", "FILE_CREATE_DATETIME", "GC_CTX_TYPE_CD", "CTX_ID");
            addRow(sheet, 1, "generic.csv", "stored.csv", "2016-02-07-08.09.10.123456", "COC", "1001");
            addRow(sheet, 2, "invalid.csv", "invalid.csv", "2016-02-07-25.09.10.123456", "COC", "1002");
            workbook.write(output);
        }

        Path target = root.resolve("target");
        Path result = root.resolve("db2-result.xlsx");
        List<String> options = Arrays.asList("--excel", excel.toString(), "--source", source.toString(),
            "--target", target.toString(), "--apply");
        List<String> automatic = new ArrayList<>(options);
        automatic.addAll(Arrays.asList("--result", result.toString()));
        assertExitCode(automatic, 1);

        Path folder = target.resolve("2016/COC/1001");
        assertEquals("saved file", contents(folder.resolve("stored.csv")));
        assertFalse(Files.exists(stored));
        assertTrue(Files.isRegularFile(generic));
        assertTrue(Files.isRegularFile(invalid));
        assertFalse(Files.exists(target.resolve("2016/COC/1002")));
        assertEquals(Arrays.asList("Mapped", "Move failed"), rowStatuses(result));

        List<String> explicit = new ArrayList<>(options);
        explicit.addAll(Arrays.asList("--file-column", "FILE_NAME", "--result", root.resolve("override.xlsx").toString()));
        assertExitCode(explicit, 1);
        assertEquals("alternate file", contents(folder.resolve("generic.csv")));
        assertFalse(Files.exists(generic));
        assertTrue(Files.isRegularFile(invalid));
    }

    private static void writeWorkbook(Path path) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream output = Files.newOutputStream(path)) {
            workbook.createSheet("Other");
            Sheet sheet = workbook.createSheet("Export");
            addRow(sheet, 0, "GC_CTX_TYPE_CD", "Attachment", "CTX_ID", "FILE_CREATE_DATETIME");
            addRow(sheet, 1, "WO", "Report.pdf", "77", "2025-01-01 00:00:00");
            addRow(sheet, 2, "WO", "Report.pdf", "77", "2025-12-31 23:59:59.999");
            addRow(sheet, 3, "WO", "Notes.txt", "77", "2025-12-31 23:59:59");
            addRow(sheet, 4, "PM", "Report.pdf", "77", "2025-04-01 12:30:00");
            Row nextYear = addRow(sheet, 5, "WO", "Report.pdf", "77");
            Cell date = nextYear.createCell(3);
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
            date.setCellStyle(dateStyle);
            date.setCellValue(LocalDateTime.of(2026, 1, 1, 0, 0));
            addRow(sheet, 6, "WO", "report.pdf", "77", "2025-06-01 00:00:00");
            workbook.write(output);
        }
    }

    private static Row addRow(Sheet sheet, int index, String... values) {
        Row row = sheet.createRow(index);
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
        return row;
    }

    private static void assertExitCode(List<String> args, int expected) throws IOException {
        ByteArrayOutputStream diagnostics = new ByteArrayOutputStream();
        try (PrintStream stream = new PrintStream(diagnostics, true, "UTF-8")) {
            int exitCode = FileMappingCli.run(args.toArray(new String[0]), stream, stream);
            assertEquals(diagnostics.toString("UTF-8"), expected, exitCode);
        }
    }

    private static List<String> rowStatuses(Path result) throws IOException {
        try (InputStream input = Files.newInputStream(result); XSSFWorkbook workbook = new XSSFWorkbook(input)) {
            Sheet sheet = workbook.getSheet("Export");
            assertNotNull(sheet);
            Cell statusHeader = null;
            for (Cell header : sheet.getRow(0)) {
                if ("Mapping Status".equals(header.getStringCellValue())) {
                    statusHeader = header;
                    break;
                }
            }
            assertNotNull("Mapping Status header", statusHeader);
            List<String> statuses = new ArrayList<>();
            for (int row = 1; row <= sheet.getLastRowNum(); row++) {
                statuses.add(sheet.getRow(row).getCell(statusHeader.getColumnIndex()).getStringCellValue());
            }
            return statuses;
        }
    }

    private static Path write(Path path, String value) throws IOException {
        Files.createDirectories(path.getParent());
        return Files.write(path, value.getBytes(StandardCharsets.UTF_8));
    }

    private static String contents(Path path) throws IOException {
        return new String(Files.readAllBytes(path), StandardCharsets.UTF_8);
    }

    private static Set<String> fileNames(Path folder) throws IOException {
        Set<String> names = new HashSet<>();
        try (DirectoryStream<Path> entries = Files.newDirectoryStream(folder)) {
            for (Path entry : entries) {
                names.add(entry.getFileName().toString());
            }
        }
        return names;
    }

    private static Set<String> reportContents(Path folder, String first, String second) throws IOException {
        return set(contents(folder.resolve(first)), contents(folder.resolve(second)));
    }

    private static Set<String> set(String... values) {
        return new HashSet<>(Arrays.asList(values));
    }
}
