package com.filemapping.aix;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardOpenOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class FileMappingCli {
    private static final String MAPPED = "Mapped";
    private static final String FAILED = "Move failed";
    private static final String NOT_FOUND = "Not found";
    private static final String CONFLICT = "Mapping conflict";

    private static final String DATE_HEADER = "FILE_CREATE_DATETIME";
    private static final String TYPE_HEADER = "GC_CTX_TYPE_CD";
    private static final String ID_HEADER = "CTX_ID";
    private static final String STORED_NAME_HEADER = "STORED_FILE_NAME";
    private static final Set<String> FILE_HEADERS = new HashSet<>(Arrays.asList(
            "FILENAME", "ATTACHMENT", "ATTACHMENTNAME", "ATTACHMENTFILENAME",
            "DOCNAME", "DOCUMENTNAME", "文件名", "文件名称", "附件名称", "附件文件名"));
    private static final Set<String> ID_HEADERS = new HashSet<>(Arrays.asList(
            "CTXID", "WONO", "WORKORDERNUMBER", "WORKORDER", "ID", "编号", "工单号"));
    private static final Set<String> VALUE_OPTIONS = new HashSet<>(Arrays.asList(
            "--excel", "--sheet", "--file-column", "--id-column", "--source", "--target", "--result"));
    private static final Pattern DATE = Pattern.compile("([0-9]{4})([-/])([0-9]{1,2})\\2([0-9]{1,2})([ T].+)?");
    private static final Pattern DB2_TIMESTAMP = Pattern.compile(
            "([0-9]{4}-[0-9]{2}-[0-9]{2})-([0-9]{2})\\.([0-9]{2})\\.([0-9]{2})(\\.[0-9]{1,9})?");
    private static final Pattern INVALID_PATH_CHAR = Pattern.compile("[<>:\"/\\\\|?*\\x00-\\x1f]");
    private static final Pattern RESERVED_NAME = Pattern.compile(
            "(?i)(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\\..*)?");
    private static final DateTimeFormatter REPORT_TIME = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss-SSS");

    private FileMappingCli() {
    }

    public static void main(String[] args) {
        System.exit(run(args, System.out, System.err));
    }

    static int run(String[] args, PrintStream out, PrintStream err) {
        if (Arrays.asList(args).contains("--help")) {
            out.println("Usage: java -jar file-mapping-aix.jar --excel INPUT.xlsx --source SOURCE --target TARGET");
            out.println("       [--sheet NAME] [--file-column HEADER] [--id-column HEADER]");
            out.println("       [--result RESULT.xlsx] [--apply]");
            out.println("All nested source directories are scanned. File names are matched exactly.");
            out.println("Default: preview only. Use --apply to move files and write the result workbook and log.");
            return 0;
        }
        try {
            Config cfg = Config.parse(args);
            Plan plan = new Plan();
            WorkbookIO.read(cfg.excel, cfg.sheet, (number, cells) -> {
                if (number == 1) {
                    plan.columns = Columns.resolve(cells, cfg);
                } else if (plan.columns != null) {
                    plan.readRow(number, cells);
                }
            });
            if (plan.columns == null) {
                throw new IOException("The worksheet must have headers in its first row.");
            }
            plan.prepare();
            out.println("Scanning " + cfg.source + " and its subdirectories...");
            scan(cfg, plan);
            printPreview(plan, out);
            if (!cfg.apply) {
                out.println("Preview finished. No files were moved.");
                return 0;
            }
            return execute(cfg, plan, out, err);
        } catch (UsageException ex) {
            err.println(ex.getMessage());
            err.println("Use --help to view the command options.");
            return 2;
        } catch (Exception ex) {
            err.println("File mapping failed: " + ex.getMessage());
            return 1;
        }
    }

    private static void scan(Config cfg, Plan plan) throws IOException {
        Files.walkFileTree(cfg.source, new SimpleFileVisitor<Path>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                if (!attrs.isRegularFile()) {
                    return FileVisitResult.CONTINUE;
                }
                plan.scanned++;
                Match match = plan.files.get(file.getFileName().toString());
                if (match != null && !Files.isSameFile(file, cfg.excel)) {
                    match.sources.add(file);
                }
                return FileVisitResult.CONTINUE;
            }
        });
    }

    private static void printPreview(Plan plan, PrintStream out) {
        long matchedRows = 0;
        long missing = 0;
        long invalid = 0;
        long conflicts = 0;
        Set<Path> folders = new HashSet<>();
        for (MappingRow row : plan.rows.values()) {
            if (CONFLICT.equals(row.status)) {
                conflicts++;
            } else if (row.status != null) {
                invalid++;
            } else if (plan.files.get(row.filename).sources.isEmpty()) {
                missing++;
            } else {
                matchedRows++;
                folders.add(row.folder);
            }
        }
        long hits = 0;
        for (Match match : plan.files.values()) {
            hits += match.sources.size();
        }
        out.printf("Scanned %d files. Matched %d files across %d rows and %d destination folders.%n",
                plan.scanned, hits, matchedRows, folders.size());
        out.printf("Not found: %d rows. Invalid: %d rows. Conflicts: %d rows.%n", missing, invalid, conflicts);
    }

    private static int execute(Config cfg, Plan plan, PrintStream out, PrintStream err) throws Exception {
        Path log = cfg.result.resolveSibling(
                cfg.result.getFileName().toString().replaceFirst("(?i)\\.xlsx$", "") + "_log.csv");
        Files.createDirectories(cfg.result.getParent());
        if (Files.exists(cfg.result, LinkOption.NOFOLLOW_LINKS)) {
            throw new IOException("Result workbook already exists: " + cfg.result);
        }
        Path draft = Files.createTempFile(cfg.result.getParent(), ".file-mapping-", ".xlsx");
        boolean rowFailures;
        try {
            try (BufferedWriter writer = Files.newBufferedWriter(
                    log, StandardCharsets.UTF_8, StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE)) {
                csv(writer, "Time", "Identifier", "File Name", "Source", "Destination", "Status", "Reason");
                for (Map.Entry<String, Match> entry : plan.files.entrySet()) {
                    String filename = entry.getKey();
                    Match match = entry.getValue();
                    Collections.sort(match.sources);
                    for (Path source : match.sources) {
                        deliver(cfg, writer, filename, match, source);
                    }
                }
            }
            rowFailures = plan.rows.values().stream().anyMatch(row -> {
                String status = row.outcome()[0];
                return FAILED.equals(status) || CONFLICT.equals(status);
            });
            WorkbookIO.writeResult(cfg.excel, cfg.sheet, draft, number -> {
                MappingRow row = plan.rows.get(number);
                return row == null ? null : row.outcome();
            });
            Files.move(draft, cfg.result);
        } catch (Exception ex) {
            err.println("If file delivery started, check this log before retrying: " + log);
            throw ex;
        } finally {
            Files.deleteIfExists(draft);
        }
        out.println("Result: " + cfg.result);
        out.println("Log: " + log);
        return rowFailures ? 1 : 0;
    }

    private static void deliver(Config cfg, BufferedWriter log, String filename, Match match, Path source)
            throws IOException {
        boolean complete = true;
        for (Map.Entry<Path, DeliveryStatus> entry : match.destinations.entrySet()) {
            Path folder = cfg.target.resolve(entry.getKey());
            Path destination = folder.resolve(filename);
            String failure = "";
            DeliveryStatus status = entry.getValue();
            try {
                if (!Files.isRegularFile(source, LinkOption.NOFOLLOW_LINKS)) {
                    throw new IOException("Source file is no longer available: " + source);
                }
                createDestination(cfg, entry.getKey());
                destination = saveAttachment(source, folder, filename, match.destinations.size() == 1);
                status.mapped = true;
            } catch (IOException ex) {
                failure = ex.toString();
                status.fail(failure);
                complete = false;
            }
            csv(log, LocalDateTime.now(), folder.getFileName(), filename, source, destination,
                    failure.isEmpty() ? MAPPED : FAILED, failure);
        }
        log.flush();
        if (complete) {
            try {
                Files.delete(source);
            } catch (IOException ex) {
                String failure = "Destination saved but source could not be deleted: " + ex;
                for (DeliveryStatus status : match.destinations.values()) {
                    status.fail(failure);
                }
                csv(log, LocalDateTime.now(), "", filename, source, "", FAILED, failure);
                log.flush();
            }
        }
    }

    private static void createDestination(Config cfg, Path relative) throws IOException {
        if (cfg.target.resolve(relative).startsWith(cfg.source)) {
            throw new IOException("Destination would be inside the source directory.");
        }
        Path folder = cfg.target;
        Files.createDirectories(folder);
        for (Path component : relative) {
            folder = folder.resolve(component);
            try {
                Files.createDirectory(folder);
            } catch (FileAlreadyExistsException exists) {
                if (!Files.isDirectory(folder, LinkOption.NOFOLLOW_LINKS)) {
                    throw exists;
                }
            }
        }
    }

    private static Path saveAttachment(Path source, Path folder, String name, boolean singleDestination)
            throws IOException {
        int dot = name.lastIndexOf('.');
        String stem = dot > 0 ? name.substring(0, dot) : name;
        String suffix = dot > 0 ? name.substring(dot) : "";
        boolean tryLink = singleDestination;
        int index = 1;
        while (true) {
            Path candidate = folder.resolve(index == 1 ? name : stem + "_" + index + suffix);
            if (tryLink) {
                try {
                    return Files.createLink(candidate, source);
                } catch (FileAlreadyExistsException exists) {
                    index++;
                    continue;
                } catch (IOException | UnsupportedOperationException unavailable) {
                    tryLink = false;
                }
            }
            OutputStream stream;
            try {
                stream = Files.newOutputStream(candidate, StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE);
            } catch (FileAlreadyExistsException exists) {
                index++;
                continue;
            }
            try (OutputStream copy = stream) {
                Files.copy(source, copy);
            } catch (IOException ex) {
                try {
                    Files.deleteIfExists(candidate);
                } catch (IOException cleanup) {
                    ex.addSuppressed(cleanup);
                }
                throw ex;
            }
            return candidate;
        }
    }

    private static void csv(BufferedWriter log, Object... fields) throws IOException {
        for (int i = 0; i < fields.length; i++) {
            if (i > 0) {
                log.write(',');
            }
            log.write('"');
            log.write(fields[i].toString().replace("\"", "\"\""));
            log.write('"');
        }
        log.newLine();
    }

    private static String yearOf(String value) {
        Matcher db2 = DB2_TIMESTAMP.matcher(value);
        if (db2.matches()) {
            String fraction = db2.group(5);
            value = db2.group(1) + "T" + db2.group(2) + ":" + db2.group(3) + ":" + db2.group(4)
                    + (fraction == null ? "" : fraction);
        }
        Matcher match = DATE.matcher(value);
        if (!match.matches()) {
            return null;
        }
        try {
            LocalDate date = LocalDate.of(
                    Integer.parseInt(match.group(1)), Integer.parseInt(match.group(3)), Integer.parseInt(match.group(4)));
            String timestamp = match.group(5);
            if (timestamp != null) {
                DateTimeFormatter.ISO_DATE_TIME.parse(date + "T" + timestamp.substring(1));
            }
            return date.getYear() > 0 ? match.group(1) : null;
        } catch (DateTimeException ex) {
            return null;
        }
    }

    private static boolean validComponent(String value) {
        return !value.isEmpty() && !value.endsWith(".") && !value.endsWith(" ")
                && !INVALID_PATH_CHAR.matcher(value).find() && !RESERVED_NAME.matcher(value).matches();
    }

    private static String valueAt(List<String> cells, int col) {
        return col >= 0 && col < cells.size() ? cells.get(col).trim() : "";
    }

    private static String key(String value) {
        return value.trim().replaceAll("[\\s_-]+", "").toUpperCase(Locale.ROOT);
    }

    private static final class Plan {
        final Map<Integer, MappingRow> rows = new LinkedHashMap<>();
        final Map<String, Match> files = new LinkedHashMap<>();
        Columns columns;
        long scanned;

        void readRow(int rowNumber, List<String> cells) {
            String id = valueAt(cells, columns.id);
            String filename = valueAt(cells, columns.filename);
            String timestamp = valueAt(cells, columns.date);
            String type = valueAt(cells, columns.type);
            if (id.isEmpty() && filename.isEmpty() && timestamp.isEmpty() && type.isEmpty()) {
                return;
            }
            MappingRow row = new MappingRow(filename);
            rows.put(rowNumber, row);
            if (!validComponent(id)) {
                row.reject(FAILED, "Invalid or blank identifier.");
            } else if (filename.isEmpty() || filename.endsWith(".") || INVALID_PATH_CHAR.matcher(filename).find()) {
                row.reject(FAILED, "Invalid or blank file name.");
            } else if (columns.date < 0) {
                row.folder = Paths.get(id);
            } else {
                String year = yearOf(timestamp);
                if (year == null) {
                    row.reject(FAILED, DATE_HEADER + " is blank or not a valid date.");
                } else if (!validComponent(type)) {
                    row.reject(FAILED, TYPE_HEADER + " is blank or is not a valid folder name.");
                } else {
                    row.group = Paths.get(year, type);
                    row.folder = row.group.resolve(id);
                }
            }
        }

        void prepare() {
            Map<String, Path> first = new HashMap<>();
            Set<String> ambiguous = new HashSet<>();
            for (MappingRow row : rows.values()) {
                if (row.folder != null) {
                    notePath(first, ambiguous, row.group);
                    notePath(first, ambiguous, row.folder);
                }
            }
            for (MappingRow row : rows.values()) {
                if (row.folder == null) {
                    continue;
                }
                if (ambiguous.contains(pathKey(row.group)) || ambiguous.contains(pathKey(row.folder))) {
                    row.reject(CONFLICT, "Different classifications resolve to the same case-insensitive folder.");
                    continue;
                }
                Match match = files.computeIfAbsent(row.filename, ignored -> new Match());
                row.delivery = match.destinations.computeIfAbsent(row.folder, ignored -> new DeliveryStatus());
            }
        }

        private static void notePath(Map<String, Path> first, Set<String> ambiguous, Path path) {
            if (path != null) {
                String key = pathKey(path);
                Path previous = first.putIfAbsent(key, path);
                if (previous != null && !previous.toString().equals(path.toString())) {
                    ambiguous.add(key);
                }
            }
        }

        private static String pathKey(Path path) {
            return path == null ? "" : path.toString().toUpperCase(Locale.ROOT).toLowerCase(Locale.ROOT);
        }
    }

    private static final class Match {
        final List<Path> sources = new ArrayList<>();
        final Map<Path, DeliveryStatus> destinations = new LinkedHashMap<>();
    }

    private static final class DeliveryStatus {
        boolean mapped;
        String failure;

        void fail(String reason) {
            if (failure == null) {
                failure = reason;
            }
        }
    }

    private static final class MappingRow {
        final String filename;
        Path group;
        Path folder;
        DeliveryStatus delivery;
        String status;
        String reason;

        MappingRow(String filename) {
            this.filename = filename;
        }

        void reject(String value, String explanation) {
            status = value;
            reason = explanation;
        }

        String[] outcome() {
            if (status != null) {
                return new String[]{status, reason};
            }
            if (delivery.failure != null) {
                return new String[]{FAILED, delivery.failure};
            }
            return new String[]{delivery.mapped ? MAPPED : NOT_FOUND, ""};
        }
    }

    private static final class Columns {
        final int id;
        final int filename;
        final int date;
        final int type;

        private Columns(int id, int filename, int date, int type) {
            this.id = id;
            this.filename = filename;
            this.date = date;
            this.type = type;
        }

        static Columns resolve(List<String> headers, Config config) throws UsageException {
            int date = exact(headers, DATE_HEADER);
            int type = exact(headers, TYPE_HEADER);
            if ((date >= 0) != (type >= 0)) {
                throw new UsageException("Year and business-type grouping requires both " + DATE_HEADER + " and " + TYPE_HEADER);
            }
            int contextId = date >= 0 ? exact(headers, ID_HEADER) : -1;
            int id = contextId >= 0 ? contextId : choose(headers, config.idColumn, ID_HEADERS, "--id-column");
            int storedName = config.fileColumn == null ? exact(headers, STORED_NAME_HEADER) : -1;
            int filename = storedName >= 0
                    ? storedName : choose(headers, config.fileColumn, FILE_HEADERS, "--file-column");
            if (id == filename || id == date || id == type || filename == date || filename == type) {
                throw new UsageException("Identifier, file name, date and type must use distinct columns.");
            }
            return new Columns(id, filename, date, type);
        }

        private static int exact(List<String> headers, String desired) throws UsageException {
            int found = -1;
            for (int i = 0; i < headers.size(); i++) {
                if (desired.equalsIgnoreCase(headers.get(i).trim())) {
                    if (found >= 0) {
                        throw new UsageException("Duplicate column header: " + desired);
                    }
                    found = i;
                }
            }
            return found;
        }

        private static int choose(List<String> headers, String selection, Set<String> aliases, String option)
                throws UsageException {
            if (selection != null) {
                int exactColumn = exact(headers, selection);
                if (exactColumn >= 0) {
                    return exactColumn;
                }
                if (selection.matches("[A-Za-z]{1,3}")) {
                    int number = 0;
                    for (char c : selection.toUpperCase(Locale.ROOT).toCharArray()) {
                        number = number * 26 + c - 'A' + 1;
                    }
                    if (number <= headers.size()) {
                        return number - 1;
                    }
                }
                throw new UsageException("Column not found for " + option + ": " + selection);
            }
            int found = -1;
            for (int i = 0; i < headers.size(); i++) {
                if (aliases.contains(key(headers.get(i)))) {
                    if (found >= 0) {
                        throw new UsageException("More than one matching column. Specify " + option + " explicitly.");
                    }
                    found = i;
                }
            }
            if (found < 0) {
                throw new UsageException("Could not detect a column. Specify " + option + " explicitly.");
            }
            return found;
        }
    }

    private static final class Config {
        final Path excel;
        final Path source;
        final Path target;
        final Path result;
        final String sheet;
        final String fileColumn;
        final String idColumn;
        final boolean apply;

        private Config(Path excel, Path source, Path target, Path result, Map<String, String> flags, boolean apply) {
            this.excel = excel;
            this.source = source;
            this.target = target;
            this.result = result;
            this.sheet = flags.get("--sheet");
            this.fileColumn = flags.get("--file-column");
            this.idColumn = flags.get("--id-column");
            this.apply = apply;
        }

        static Config parse(String[] args) throws IOException, UsageException {
            Map<String, String> flags = new HashMap<>();
            boolean apply = false;
            for (int i = 0; i < args.length; i++) {
                String arg = args[i];
                if ("--apply".equals(arg) && !apply) {
                    apply = true;
                } else if (VALUE_OPTIONS.contains(arg) && !flags.containsKey(arg)
                        && i + 1 < args.length && !args[i + 1].startsWith("--")) {
                    flags.put(arg, args[++i]);
                } else {
                    throw new UsageException("Invalid or repeated argument: " + arg);
                }
            }
            for (String required : Arrays.asList("--excel", "--source", "--target")) {
                if (!flags.containsKey(required)) {
                    throw new UsageException("Missing required option: " + required);
                }
            }
            Path excel = Paths.get(flags.get("--excel")).toRealPath();
            Path source = Paths.get(flags.get("--source")).toRealPath();
            Path target = resolveFuturePath(Paths.get(flags.get("--target")));
            Path reportFolder = excel.startsWith(source) ? target : excel.getParent();
            Path result = flags.containsKey("--result")
                    ? resolveFuturePath(Paths.get(flags.get("--result")))
                    : reportFolder.resolve("file-mapping-result-" + LocalDateTime.now().format(REPORT_TIME) + ".xlsx");
            if (!Files.isDirectory(source) || !Files.isRegularFile(excel)) {
                throw new UsageException("Source must be a directory and Excel input must be an existing file.");
            }
            if (!isXlsx(excel) || !isXlsx(result)) {
                throw new UsageException("Excel input and result must have the .xlsx extension.");
            }
            if (target.startsWith(source)) {
                throw new UsageException("The target must be outside the source directory.");
            }
            if (result.startsWith(source) || result.equals(excel)) {
                throw new UsageException("Choose a result workbook outside the source directory and different from the input.");
            }
            return new Config(excel, source, target, result, flags, apply);
        }

        private static Path resolveFuturePath(Path path) throws IOException {
            Path absolute = path.toAbsolutePath().normalize();
            Path existing = absolute;
            while (!Files.exists(existing, LinkOption.NOFOLLOW_LINKS)) {
                existing = existing.getParent();
            }
            return existing.toRealPath().resolve(existing.relativize(absolute));
        }

        private static boolean isXlsx(Path path) {
            return path.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".xlsx");
        }
    }

    private static final class UsageException extends Exception {
        UsageException(String message) {
            super(message);
        }
    }
}
