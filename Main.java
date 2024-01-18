import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.input.InfiniteCircularInputStream;
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

    public static final String ANSI_RESET = "\u001B[0m";
    public static final String ANSI_BLACK = "\u001B[30m";
    public static final String ANSI_RED = "\u001B[31m";
    public static final String ANSI_GREEN = "\u001B[32m";
    public static final String ANSI_YELLOW = "\u001B[33m";
    public static final String ANSI_BLUE = "\u001B[34m";
    public static final String ANSI_PURPLE = "\u001B[35m";
    public static final String ANSI_CYAN = "\u001B[36m";
    public static final String ANSI_WHITE = "\u001B[37m";

    
    public static void main(String[] args) throws IOException {
        /*
         * Declarations
         */

        final String wbfile, docfile;
        final XSSFWorkbook workbook;
        final XWPFDocument template;
        final List<XSSFName> xssfnames;
        final Iterator<Sheet> sheetItr;
        final List<String> tags;
        final MemberAccessor<Object, String, Object> maccess;
        // final MemberUpdater<Object, String, Object> mupdate; // No mupdating here.
        
        final Map<String, Object> vars = new HashMap<>();
        final Slot<Interpreter> in = Slot.of("interpreter", null);
        final Object TOSTRINGIFY, STRINGIFY, DATEIFY, INTIFY, NUMIFY, BOOLIFY, 
            ROW, COLUMN, FORMULIFY, DUPLICATE_NAME = new Object();
		final List<Replacer> replacers = new LinkedList<>();
        
        
        wbfile = "C:\\Users\\SaifeldinMohamed\\Desktop\\BEC Calculator v2.4.2 (SMG).xlsm";
        docfile = "C:\\Users\\SaifeldinMohamed\\Desktop\\S2 BEC Report Template v1.23 - SMG.docx";

        template = new XWPFDocument(Files.newInputStream(Paths.get(docfile)));
        workbook = new XSSFWorkbook(wbfile);
        xssfnames = workbook.getAllNames();
        sheetItr = workbook.sheetIterator();

        final int nContracts = 2;
        Preprocessor.repeatSections(template, nContracts);
        for (int i = 1; i <= nContracts; i += 1) {
            vars.put("C" + i + "_Ord", i);
        }

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

        vars.put("str", STRINGIFY = new Object());
        vars.put("tostr", TOSTRINGIFY = new Object());
        vars.put("date", DATEIFY = new Object());
        vars.put("int", INTIFY = new Object());
        vars.put("num", NUMIFY = new Object());
        vars.put("bool", BOOLIFY = new Object());
        vars.put("row", ROW = new Object());
        vars.put("col", COLUMN = new Object());
        vars.put("formula", FORMULIFY = new Object());
        vars.put("f", FORMULIFY);

        maccess = (obj, member) -> {
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

                throw new IllegalArgumentException("Sheet member \"" + member + "\" is not a defined name, or a valid cell address (like A1).");
            }
            
            final Object thing = in.value().getVariable(member);
            if (obj == TOSTRINGIFY) {
                return String.valueOf(thing);
            } 
            
            if (thing instanceof XSSFCell) {
                final XSSFCell cell = (XSSFCell) thing;
              
                if (obj == INTIFY && (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)) {
                    return ((Double) cell.getNumericCellValue()).intValue();
                }
                else if (obj == NUMIFY && (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)) {
                    return cell.getNumericCellValue();
                }
                else if (obj == DATEIFY && (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA)) {
                    return cell.getDateCellValue();
                }
                else if (obj == BOOLIFY && (cell.getCellType() == CellType.BOOLEAN || cell.getCellType() == CellType.FORMULA)) {
                    return cell.getBooleanCellValue();
                }
                else if (obj == STRINGIFY) {
                    return cell.getStringCellValue();
                }
                else if (obj == FORMULIFY && cell.getCellType() == CellType.FORMULA) {
                    return cell.getCellFormula();
                }
                else if (obj == ROW) return cell.getRowIndex() + 1;
                else if (obj == COLUMN) return cell.getColumnIndex() + 1;
            }
            
            else if (thing instanceof Double && obj == INTIFY) {
                return ((Double) thing).intValue();
            }

            return null;
        };

        tags = Extractor.extractTags(template);
		for (String tag : tags) {
			// Assume tags are of the form <<...>> or <<<...>>>
			String prog;
			if (tag.startsWith("<<<") && tag.endsWith(">>>")) {
				// We can some day use triple quoted tags for something. 
				// For now, they are treated the same. -SMG
				prog = tag.substring(3, tag.length() - 3);
			}
			else if (tag.startsWith("<<") && tag.endsWith(">>")) {
				prog = tag.substring(2, tag.length() - 2);
			}
			else {
				throw new IllegalArgumentException("Tag should start and end with double or triple angle brackets (<>)." 
				+ "\nInstead it looks like: " + tag);
			}

			final int colon = prog.lastIndexOf(':');
			final int quote = Math.max(prog.lastIndexOf('\"'), prog.lastIndexOf('\''));
			final String spec;
			if (colon != -1 && quote < colon) {
				spec = prog.substring(colon + 1, prog.length()).strip();
				prog = prog.substring(0, colon);
			}
			else {
				spec = null;
			}

            // if (true) {
            //     prog =  
            //         "    let i = %d; \r\n\r" + //
            //         "    let siteCell = \"R\" + (row.CSiteCol + i) + \"C\" + col.CSiteCol; \r\n" + //
            //         "    siteCell = Contracts.siteCell;\r\n" + //
            //         "    f.siteCell"
            //             // "    (f.siteCell == \"\" or f.siteCell == empty) and str.siteCell != \"\";"
            //         //  + "    let siteCell = \"R\" + CurrentRow + \"C\" + col.CSiteCol; "
            //            ;

            //     for (int i = 1; i <= 10; i += 1) {
            //         in.value(new Interpreter(String.format(prog, i), vars));
            //         in.value().setMemberAccessCallback(maccess);
            //         final Object value = in.value().interpret();
            //         System.out.println(i + " " + value);
            //     }
            //     return;
            // }


			// Date and number formats should be dealt with here
			// final Interpreter interpreter = new Interpreter(prog, vars);
            in.value(new Interpreter(prog, vars));
			in.value().setMemberAccessCallback(maccess);

			try {
				final Object output = in.value().interpret();
                final String result = output == null ? "" : format(output, spec);

				replacers.add(replacer(tag, result));

                if (tag.contains("\n")) {
                    tag = tag.substring(0, Math.min(tag.indexOf('\n'), 16)) + "...\\n>>";
                }
                System.out.println(tag + " -> " + result.trim());
			}
			catch (Exception e) {
				report(e, tag);
                // e.printStackTrace();
                // return;
			}

			// Update our variables so we can have persistance across tag executions.
			vars.putAll(in.value().getGlobalScopeVariables());
		}

        // Save to file
        Replacers.orderedReplace(template, replacers);
        Path parent = Paths.get(docfile).getParent();
        template.write(new FileOutputStream(parent.toString() + "/output.docx"));
        template.close();
    }

    
	static Replacer replacer(String bookmark, String replacement) {
		return Replacer.of(bookmark, replacement);
	}

	static void report(Exception e, String program) {
		// StringWriter sw = new StringWriter();
		// e.printStackTrace(new PrintWriter(sw));

		// Core.getLogger("Tag Extractor").error(new StringBuilder("Something went wrong processing the following program:")
		// 	.append("\n").append(program).append("\n").append("Error Details:\n")
		// 	.append(String.format("Type: %s \n", e.getClass().getSimpleName()))
		// 	.append(String.format("Message: %s \n", e.getMessage()))
		// 	.append(String.format("Stack Trace: %s \n", sw.toString())).toString());

        if (program.contains("\n")) {
            program = program.substring(0, Math.min(program.indexOf('\n'), 16)) + "...\\n>>";
        }
        
        final String msg = e.getMessage() == null ? 
            "[" + e.getClass().getSimpleName() + "]" : 
            e.getMessage().replace("\n", " ");

        System.out.println(program + " -> " + ANSI_RED + "Error: " + msg + ANSI_RESET);
	}

    
	private static final DateFormat shortDate = new SimpleDateFormat("dd/MM/yyyy");
 	static String format(Object thing, String spec) {

		if (thing instanceof BigDecimal) return format(((BigDecimal) thing).doubleValue(), spec);
		else if (spec == null) return String.valueOf(thing);
		else if (spec.contains("%")) return String.format(spec, thing);
		else switch (spec) {
			case "short_date":
			case "sdate": return shortDate.format(thing);
			case "currency":
			case "curr": return format(thing, "£ %,.2f");

			// Legacy stuff. I want to phase it out someday.
			case "currp": return String.format("%s p", NumberFormat.getInstance().format((Double) thing * 100));
			case "currpx": return String.format("%s p", NumberFormat.getInstance().format((Double) thing));
			default: throw new IllegalArgumentException("Unrecognised format spec: " + spec);
		}

	}
    

    private static ReferenceType reftype(String address) {
        
        if (ref(address) != null) {
            return ReferenceType.CELL;
        }

        String[] parts = address.replace("$", "").split(":");
        if (parts.length != 2) return ReferenceType.INVALID;
        
        try {
            CellReference.convertColStringToIndex(parts[1]);
            // No error means valid column
            return ReferenceType.COLUMN;
        }
        catch (Exception e) {}
        
        try {
            Integer.parseInt(parts[1]);
            // No error means valid row
            return ReferenceType.ROW;
        }
        catch (Exception e) {}

        return ReferenceType.INVALID;
    }

    private static final String r1c1pattern = "R\\d+C\\d+";
    private static CellReference refR1C1(String ref) {
        if (!(ref = ref.trim()).matches(r1c1pattern)) return null;

        final int cindex = ref.indexOf('C');
        return new CellReference(
            Integer.parseInt(ref.substring(1, cindex)) - 1, 
            Short.parseShort(ref.substring(cindex + 1)) - 1
        );
    }

    private static CellReference ref(String address) {        
        try {
            return new CellReference(address);
        }
        catch (Exception e) {
            return refR1C1(address);
        }
        
    }

    // private static XSSFCell cell(XSSFWorkbook wb, String loc) {
    //     final CellReference ref = ref(loc);

    //     final XSSFSheet sheet = wb.getSheet(ref.getSheetName());
    //     if (sheet == null) return null;
        
    //     final XSSFRow row = sheet.getRow(ref.getRow());
    //     if (row == null) return null;
        
    //     return row.getCell(ref.getCol());
    // }
    
    private static XSSFCell cell(XSSFSheet sheet, String loc) {
        final CellReference ref = ref(loc);

        final XSSFRow row = sheet.getRow(ref.getRow());
        if (row == null) return null;
        
        return row.getCell(ref.getCol());
    }

    private static String codename(Sheet sheet) {
        return ((XSSFSheet) sheet).getCTWorksheet().getSheetPr().getCodeName();
    }

    private static boolean validName(String name) {
        return new Tokeniser(name).nextToken().isAny(TokenType.Qualifier);
    }
}
