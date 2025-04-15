import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.forenzix.common.Pair;
import com.forenzix.common.Slot;
import com.forenzix.excel.ReferenceType;
import com.forenzix.interpreter.Interpreter;
import com.forenzix.interpreter.TokenType;
import com.forenzix.interpreter.Tokeniser;
import com.forenzix.interpreter.Interpreter.MemberAccessor;
import com.forenzix.word.Extractor;
import com.forenzix.word.Preprocessor;
import com.forenzix.word.Replacer;
import com.forenzix.word.Replacers;

/**
 * The main entry point of the excel reporter program. It uses a few libraries
 * made for other projects to work properly
 * 
 * @author SMG
 * @see Interpreter
 * @see Preprocessor
 */
public class Main {

    public static final String VERSION = "1.1";
    public static final String ANSI_RESET = "\u001B[0m",
            ANSI_BLACK = "\u001B[30m",
            ANSI_RED = "\u001B[31m",
            ANSI_GREEN = "\u001B[32m",
            ANSI_YELLOW = "\u001B[33m",
            ANSI_BLUE = "\u001B[34m",
            ANSI_PURPLE = "\u001B[35m",
            ANSI_CYAN = "\u001B[36m",
            ANSI_WHITE = "\u001B[37m";

    public static int MAX_THREAD_COUNT = 0; // Set to 0 to disable multi-threading

    public static void main(String[] args) throws IOException {
        final Date start = new Date();
        try {
            mainProcedure(args);
        } catch (Exception e) {
            System.out.println(ANSI_RED + "Something went wrong. Please inspect the following error.");
            e.printStackTrace();
        }

        System.out.println("Total time taken: %.2f seconds".formatted((new Date().getTime() - start.getTime()) / 1000.0D));
        System.out.print("Press Enter to close this window.");
        System.out.print(ANSI_RESET);
        final Scanner scan = new Scanner(System.in);
        scan.nextLine();
        scan.close();
        System.exit(0);
    }

    public static Map<String, List<String>> parseArgs(String[] args) {
        final Map<String, List<String>> argmap = new HashMap<>();
        String state = "a";
        for (String arg : args) {
            if (arg.startsWith("-")) {
                switch (arg) {
                    case "-t":
                    case "-template":
                        state = "t";
                        break;
                    case "-m":
                    case "-multi-thred":
                        state = "m";
                        break;
                    case "-c":
                    case "-calc":
                        state = "c";
                        break;
                    case "-o":
                    case "-out":
                    case "-output":
                        state = "o";
                        break;
                    case "-v":
                    case "-version":
                        System.out.println("excelr v" + VERSION);
                        System.exit(0);
                        break;
                    default:
                        throw new IllegalArgumentException("Unknown argument: %s".formatted(arg));
                }
                continue;
            }
            
            switch (state) {
                case "t":
                    argmap.computeIfAbsent(state, _k -> new ArrayList<String>()).add(arg);
                    state = "a";
                    break;
                case "c":
                case "o":
                    argmap.computeIfAbsent(state, _k -> new ArrayList<String>()).add(arg);
                    break;
                case "m":
                    argmap.computeIfAbsent(state, _k -> new ArrayList<String>()).add(arg);
                    state = "a";
                    break;
                default:
                    throw new IllegalArgumentException();
            }
        }

        /* if (argmap.get("c").size() != argmap.get("o").size()) {
            throw new IllegalArgumentException("The input size does not match the output size");
        }
        else */ if (argmap.get("t").size() != 1) {
            throw new IllegalArgumentException("One template must be provided at most");
        }

        if (argmap.containsKey("m") && argmap.get("m").size() >= 1) {
            MAX_THREAD_COUNT = Integer.parseInt(argmap.get("m").get(0));
        }

        return argmap;
    }

    public static void mainProcedure(String[] args) throws IOException, InterruptedException {

        final String docfile;
        final List<String> wbfiles, outfiles;
        final Map<String, List<String>> argmap = parseArgs(args);

        if ((docfile = argmap.get("t").get(0)) == null || (wbfiles = argmap.get("c")).size() == 0) {
            System.out.println(
                ANSI_RED +
                "The template and workbook files were not supplied correctly\n." +
                "Please use the format: java app -t <template> -c <calculator(s)> [-o <destination(s)>]" +
                ANSI_RESET
            );

            throw new IllegalArgumentException("One or more required files was null");
        }

        if (argmap.get("o") == null) {
            outfiles = wbfiles.stream()
                .map((wb) -> {
                    final String name = Paths.get(wb).getFileName().toString();
                    final String basename = name.substring(0, name.lastIndexOf("."));
                    final String parent = wb.substring(0, wb.length() - name.length());
                    return parent + basename + " output.docx";
                })
                .collect(Collectors.toList());
        }
        else {
            outfiles = argmap.get("o");
        }

        final List<Thread> threads = new ArrayList<>();
        final boolean singleReport = wbfiles.size() == 1;
        if (!singleReport) System.out.println("Generating %d reports.".formatted(wbfiles.size()));
        if (MAX_THREAD_COUNT > 0) System.out.println("Using %d threads.".formatted(MAX_THREAD_COUNT));
        for (int i = 0; i < wbfiles.size(); i += 1) {
            final String wb = wbfiles.get(i), out = outfiles.get(i);
            final Supplier<Void> runner = () -> {
                try {
                    produceReport(wb, out, docfile, singleReport);
                    System.out.println(ANSI_GREEN + "Report '%s' generated successfully.".formatted(out) + ANSI_RESET);
                }
                catch (IOException e) {
                    if (!singleReport) System.out.println(ANSI_RED + "Report '%s' failed. Run independently to get more details.".formatted(out) + ANSI_RESET);
                    else e.printStackTrace();
                }
                return null;
            };

            if (MAX_THREAD_COUNT != 0) {
                while (threads.size() >= MAX_THREAD_COUNT) {
                    Thread thread = threads.removeFirst();
                    thread.join();
                }
                final Thread thread = new Thread(new Runnable() { public void run() { runner.get(); } });
                thread.start();
                threads.add(thread);
            } 
            else {
                runner.get();
            }
        }

        for (Thread thread : threads)
            thread.join();
    }

    
    final static Object 
    TOSTRINGIFY = new Object(),
    STRINGIFY = new Object(),
    DATEIFY = new Object(),
    INTIFY = new Object(),
    NUMIFY = new Object(),
    BOOLIFY = new Object(),
    ROW = new Object(),
    COLUMN = new Object(),
    FORMULIFY = new Object(),
    DUPLICATE_NAME = new Object();

    @SuppressWarnings({ "rawtypes", "unchecked" })
    final static Function<Slot<Interpreter>, MemberAccessor<Object, String, Object>> makeMaccess = (in) -> (obj, member) -> {
        if (obj instanceof Pair) {
            final Pair<Sheet, Map<String, XSSFCell>> sheetPair = (Pair) obj;

            // There are multiple ways to access data in an Excel sheet
            // 1. It is a named range
            if (sheetPair.value.containsKey(member)) {
                return sheetPair.value.get(member);
            }

            // 2. The given member is a variable that contains a valid address
            final Object memVal = in.value().findVariable(member).orElse(Map.of()).get(member);
            if (memVal != null && memVal instanceof String && reftype((String) memVal) == ReferenceType.CELL) {
                return cell((XSSFSheet) sheetPair.key, (String) memVal);
            }

            // 3. It is a valid cell reference (eg. Home.A1)
            if (reftype(member) == ReferenceType.CELL) {
                return cell((XSSFSheet) sheetPair.key, member);
            }

            throw new IllegalArgumentException("Sheet member \"" + member +
                    "\" is not a defined name, or a valid cell address (like A1).");
        }

        final Object thing = in.value().getVariable(member);
        if (obj == TOSTRINGIFY) {
            return String.valueOf(thing);
        }

        if (thing instanceof XSSFCell) {
            final XSSFCell cell = (XSSFCell) thing;

            if (obj == INTIFY
                    && (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)) {
                return ((Double) cell.getNumericCellValue()).intValue();
            } else if (obj == NUMIFY
                    && (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)) {
                return cell.getNumericCellValue();
            } else if (obj == DATEIFY
                    && (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)) {
                return cell.getDateCellValue();
            } else if (obj == BOOLIFY
                    && (cell.getCellType() == CellType.BOOLEAN || cell.getCellType() == CellType.FORMULA)) {
                return cell.getBooleanCellValue();
            } else if (obj == STRINGIFY) {
                return cell.getStringCellValue();
            } else if (obj == FORMULIFY && cell.getCellType() == CellType.FORMULA) {
                return cell.getCellFormula();
            } else if (obj == ROW)
                return cell.getRowIndex() + 1;
            else if (obj == COLUMN)
                return cell.getColumnIndex() + 1;
        }

        else if (thing instanceof Double && obj == INTIFY) {
            return ((Double) thing).intValue();
        }

        return null;
    };

    public static void produceReport(String wbfile, String outfile, String docfile, boolean printLogs) throws FileNotFoundException, IOException {

        final XSSFWorkbook workbook;
        final XWPFDocument template;
        final List<XSSFName> xssfnames;
        final Iterator<Sheet> sheetItr;
        final List<String> tags;
        
        final Map<String, Object> vars = new HashMap<>();
        final Slot<Interpreter> in = Slot.of("interpreter", null);
        final List<Replacer> replacers = new LinkedList<>();

        // final MemberUpdater<Object, String, Object> mupdate; // No mupdating here.
        final var maccess = makeMaccess.apply(in);

        template = new XWPFDocument(Files.newInputStream(Paths.get(docfile)));
        workbook = new XSSFWorkbook(wbfile);
        xssfnames = workbook.getAllNames();
        sheetItr = workbook.sheetIterator();

        while (sheetItr.hasNext()) {
            final Pair<Sheet, Map<String, XSSFCell>> sheetPair = Pair.of(sheetItr.next(), new HashMap<>());

            vars.put(codename(sheetPair.key), sheetPair);
            if (validName(sheetPair.key.getSheetName())) {
                vars.put(sheetPair.key.getSheetName(), sheetPair);
            }
        }

        for (XSSFName xname : xssfnames) {
            if (reftype(xname.getRefersToFormula()) == ReferenceType.CELL) {

                final CellReference ref = ref(xname.getRefersToFormula());
                final String sname = ref.getSheetName(), name = xname.getNameName();
                final int row = ref.getRow(), col = ref.getCol();

                // I'll get rid of the warning. Eventually... -SMG
                @SuppressWarnings("unchecked")
                final Map<String, XSSFCell> sheet = (((Pair<Sheet, HashMap<String, XSSFCell>>) vars.get(sname))).value;
                final XSSFCell cell = workbook.getSheet(sname).getRow(row).getCell(col);

                vars.put(name, vars.containsKey(name) ? DUPLICATE_NAME : cell);
                sheet.put(name, cell);
            }
        }

        for (String key : vars.keySet()) {
            if (vars.get(key) == DUPLICATE_NAME) {
                vars.remove(key);
            }
        }

        workbook.close();

        vars.put("str", STRINGIFY);
        vars.put("tostr", TOSTRINGIFY);
        vars.put("date", DATEIFY);
        vars.put("int", INTIFY);
        vars.put("num", NUMIFY);
        vars.put("bool", BOOLIFY);
        vars.put("row", ROW);
        vars.put("col", COLUMN);
        vars.put("formula", FORMULIFY);
        vars.put("f", FORMULIFY);
        vars.put("Now", Date.from(Instant.now()));

        // Find the number of contracts according to the template
        final List<String> specialTags = Extractor.extractSpecialTags(template);

        int NumberOfContracts = 0;
        for (String tag : specialTags) {
            // Assume tags are of the form <<...>> or <<<...>>>
            String prog = tag.substring(3, tag.length() - 3);

            in.value(new Interpreter(prog, vars));
            in.value().setMemberAccessCallback(maccess);

            try {
                in.value().interpret();
                replacers.add(replacer(tag, ""));
            } catch (Exception e) {
                report(e, tag);
            }

            if (in.value().findVariable("NUMBER_OF_CONTRACTS").isPresent()) {
                NumberOfContracts = ((Double) in.value().getVariable("NUMBER_OF_CONTRACTS")).intValue();
            }
        }

        Replacers.orderedReplace(template, replacers);
        replacers.clear();

        final int nContracts = NumberOfContracts;
        Preprocessor.repeatSections(template, nContracts);
        for (int i = 1; i <= nContracts; i += 1) {
            vars.put("C" + i + "_Ord", i);
        }

        tags = Extractor.extractTags(template);
        for (String tag : tags) {
            // Assume tags are of the form <<...>>
            String prog = tag.substring(2, tag.length() - 2);

            final int colon = prog.lastIndexOf(':');
            final int quote = Math.max(prog.lastIndexOf('\"'), prog.lastIndexOf('\''));
            final String spec;
            if (colon != -1 && quote < colon) {
                spec = prog.substring(colon + 1, prog.length()).strip();
                prog = prog.substring(0, colon);
            } else {
                spec = null;
            }

            in.value(new Interpreter(prog, vars));
            in.value().setMemberAccessCallback(maccess);

            try {
                final Object output = in.value().interpret();
                final String result = output == null ? "" : format(output, spec);

                replacers.add(replacer(tag, result));

                if (tag.contains("\n")) {
                    tag = tag.substring(0, Math.min(tag.indexOf('\n'), 16)) + "...\\n>>";
                }
                if (printLogs) System.out.println(tag + " -> " + result.trim());
            } catch (Exception e) {
                if (printLogs) report(e, tag);
            }

            // Update our variables so we can have persistance across tag executions.
            vars.putAll(in.value().getGlobalScopeVariables());
        }

        // Replace tags
        Replacers.orderedReplace(template, replacers);
        
        // Remove unwanted sections
        Preprocessor.removeSections(template);
        
        // Save to file
        try (final FileOutputStream fs = new FileOutputStream(outfile)) {
            template.write(fs);
            template.close();
        }
    }
 
    static Replacer replacer(String bookmark, String replacement) {
        return Replacer.of(bookmark, replacement);
    }

    static void report(Exception e, String program) {
        if (program.contains("\n")) {
            program = program.substring(0, Math.min(program.indexOf('\n'), 16)) + "...\\n>>";
        }

        final String msg = e.getMessage() == null ? "[" + e.getClass().getSimpleName() + "]"
                : e.getMessage().replace("\n", " ");

        System.out.println(program + " -> " + ANSI_RED + "Error: " + msg + ANSI_RESET);
    }

    private static final DateFormat shortDate = new SimpleDateFormat("dd/MM/yyyy");

    static String format(Object thing, String spec) {

        if (thing instanceof BigDecimal)
            return format(((BigDecimal) thing).doubleValue(), spec);
        else if (spec == null)
            return String.valueOf(thing);
        else if (spec.contains("%"))
            return String.format(spec, thing);
        else
            switch (spec) {
                case "short_date":
                case "sdate":
                    return shortDate.format(thing);
                case "currency":
                case "curr":
                    return format(thing, "£ %,.2f");

                // Legacy stuff. I want to phase it out someday.
                case "currp":
                    return String.format("%s p", NumberFormat.getInstance().format((Double) thing * 100));
                case "currpx":
                    return String.format("%s p", NumberFormat.getInstance().format((Double) thing));
                default:
                    throw new IllegalArgumentException("Unrecognised format spec: " + spec);
            }

    }

    private static ReferenceType reftype(String address) {

        if (ref(address) != null) {
            return ReferenceType.CELL;
        }

        String[] parts = address.replace("$", "").split(":");
        if (parts.length != 2)
            return ReferenceType.INVALID;

        try {
            CellReference.convertColStringToIndex(parts[1]);
            // No error means valid column
            return ReferenceType.COLUMN;
        } catch (Exception e) {
        }

        try {
            Integer.parseInt(parts[1]);
            // No error means valid row
            return ReferenceType.ROW;
        } catch (Exception e) {
        }

        return ReferenceType.INVALID;
    }

    private static final String r1c1pattern = "R\\d+C\\d+";

    private static CellReference refR1C1(String ref) {
        if (!(ref = ref.trim()).matches(r1c1pattern))
            return null;

        final int cindex = ref.indexOf('C');
        return new CellReference(
                Integer.parseInt(ref.substring(1, cindex)) - 1,
                Short.parseShort(ref.substring(cindex + 1)) - 1);
    }

    private static CellReference ref(String address) {
        try {
            return new CellReference(address);
        } catch (Exception e) {
            return refR1C1(address);
        }

    }

    private static XSSFCell cell(XSSFSheet sheet, String loc) {
        final CellReference ref = ref(loc);

        final XSSFRow row = sheet.getRow(ref.getRow());
        if (row == null)
            return null;

        return row.getCell(ref.getCol());
    }

    private static String codename(Sheet sheet) {
        return ((XSSFSheet) sheet).getCTWorksheet().getSheetPr().getCodeName();
    }

    private static boolean validName(String name) {
        return new Tokeniser(name).nextToken().isAny(TokenType.Qualifier);
    }
}
