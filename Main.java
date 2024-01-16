import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;

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
import com.forenzix.excel.NameReader;
import com.forenzix.excel.ReferenceType;
import com.forenzix.interpreter.Interpreter;
import com.forenzix.interpreter.Interpreter.MemberAccessor;
import com.forenzix.word.Extractor;
import com.forenzix.word.Preprocessor;

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
        XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\SaifeldinMohamed\\Desktop\\BEC Calculator v2.4.2 (SMG).xlsm");
        XWPFDocument template = new XWPFDocument();
        final List<XSSFName> xssfnames = workbook.getAllNames();

        final Map<String, Object> vars = new HashMap<>();
        Iterator<Sheet> sheetItr = workbook.sheetIterator();
        while (sheetItr.hasNext()) {
            final Sheet sheet = sheetItr.next();
            vars.put(sheet.getSheetName(), Pair.of(sheet, new HashMap<String, XSSFCell>()));
        }

        final List<String> tags = Extractor.extractTags(template);
        for (XSSFName xname : xssfnames) {
            if (reftype(xname.getRefersToFormula()) == ReferenceType.CELL) {
                
                final CellReference ref = ref(xname.getRefersToFormula());
                final String sname = ref.getSheetName(), name = xname.getNameName();
                final int row = ref.getRow(), col = ref.getCol();
                
                // I'll get rid of the warning... eventually. -SMG
                final Map<String, XSSFCell> sheet = (((Pair<Sheet, HashMap<String, XSSFCell>>) vars.get(sname))).value;

                sheet.put(name, workbook.getSheet(sname).getRow(row).getCell(col));
            }
        }

        final Slot<Interpreter> in = Slot.of("interpreter", null);

        final Object STRINGIFY, DATEIFY, INTIFY, NUMIFY, BOOLIFY;

        vars.put("str", STRINGIFY = new Object());
        vars.put("date", DATEIFY = new Object());
        vars.put("int", INTIFY = new Object());
        vars.put("num", NUMIFY = new Object());
        vars.put("bool", BOOLIFY = new Object());

        final SimpleDateFormat df = new SimpleDateFormat("dd/MM/yyyy");

        final MemberAccessor<Object, String, Object> maccess = (obj, member) -> {
            if (obj instanceof Pair) {
                final Pair<Sheet, HashMap<String, XSSFCell>> sheetPair = (Pair) obj;
                
                // There are multiple ways to access data in an Excel sheet
                // 1. It is a named range
                if (sheetPair.value.containsKey(member)) {
                    return sheetPair.value.get(member);
                }

                // 2. The given member is a variable that contains a valid address
                Object memVal = in.value().findVariable(member).orElse(Map.of()).get(member);
                if (memVal != null && memVal instanceof String && reftype((String) memVal) == ReferenceType.CELL) {
                    return cell((XSSFSheet) sheetPair.key, (String) memVal);
                }

                // 3. It is a valid cell reference (eg. Home.A1)
                if (reftype(member) == ReferenceType.CELL) {
                    return cell((XSSFSheet) sheetPair.key, member);
                }
                

                throw new IllegalArgumentException("Sheet member \"" + member + "\" is not a defined name, or a valid cell address (like A1).");
                
            }
            else if (obj == STRINGIFY) return ((XSSFCell) in.value().getVariable(member)).getStringCellValue();
            else if (obj == DATEIFY) return df.format(((XSSFCell) in.value().getVariable(member)).getDateCellValue());
            else if (obj == INTIFY) return (int) ((XSSFCell) in.value().getVariable(member)).getNumericCellValue();
            else if (obj == NUMIFY) return ((XSSFCell) in.value().getVariable(member)).getNumericCellValue();
            else if (obj == BOOLIFY) return ((XSSFCell) in.value().getVariable(member)).getBooleanCellValue();

            return null;
        };

        for (String prog : List.of(
            "let ha = Home.Assessor; str.ha + \" Yo!\"",
            "Bills.Assessor",
            "Assessor",
            "Home.B2",
            "let x = \"C4\"; Home.x;"
        )) {

            in.value(new Interpreter(prog, vars));
            in.value().setMemberAccessCallback(maccess);

            try {
                System.out.println(in.value().interpret());
            }
            catch (Exception e) {
                String msg = e.getMessage() == null ? 
                    "[" + e.getClass().getSimpleName() + "]" : 
                    e.getMessage().replace("\n", " ");

                System.out.println(ANSI_RED + "Error: " + msg + ANSI_RESET);
            }
        }

        try {
            workbook.close();
        }
        catch (Exception e) {}


        /*
         * Interpreter Preparation
         * 
         * 1. Add all named ranges as variables
         * This will allow <<ClaimDate>> -> 31/12/2023 (Done)
         * 
         * 2. Named ranges that represent single columns or rows as indices
         * This will allow <<SDateCol>> -> 3
         * 
         * 3. Add sheets as variables. Tweak maccess to interpret sheet members 
         * as cell addresses.
         * This will allow <<shtHome.A1>> -> "Yo" (Done)
         * 
         * 4. Add dynamic sheet accessor for contract data. Unsure if effective.
         * This will allow <<shtHome[CX, 1 + 3]>> -> 21
         * 
         * 5. Add special variables that take in a range and produce information.
         * Unsure if useful.
         * This will allow <<Row.CNameCell>> -> 5
         */
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

    private static CellReference ref(String address) {        
        try {
            return new CellReference(address);
        }
        catch (Exception e) {
            return null;
        }
        
    }

    private static XSSFCell cell(XSSFWorkbook wb, String loc) {
        final CellReference ref = ref(loc);

        final XSSFSheet sheet = wb.getSheet(ref.getSheetName());
        if (sheet == null) return null;
        
        final XSSFRow row = sheet.getRow(ref.getRow());
        if (row == null) return null;
        
        return row.getCell(ref.getCol());
    }
    
    private static XSSFCell cell(XSSFSheet sheet, String loc) {
        final CellReference ref = ref(loc);

        final XSSFRow row = sheet.getRow(ref.getRow());
        if (row == null) return null;
        
        return row.getCell(ref.getCol());
    }
}
