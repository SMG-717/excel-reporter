import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.forenzix.common.Pair;
import com.forenzix.excel.NameReader;
import com.forenzix.interpreter.Interpreter;
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
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook("C:\\Users\\SaifeldinMohamed\\Desktop\\BEC Calculator v2.4.2 (SMG).xlsm");
        List<Pair<String, String>> names = NameReader.extractNames(workbook);
        final List<XSSFName> xssfnames = workbook.getAllNames();
        // System.out.println(names);


        // System.out.println(CellReference.convertColStringToIndex("Home!$X:$X"));
        // System.out.println(CellReference.convertColStringToIndex("$X:$X"));
        // System.out.println(CellReference.convertColStringToIndex("X:X")); // 16091
        System.out.println(CellReference.convertColStringToIndex("$X")); // 23
        
        
        // final CellReference ref = new CellReference("Home!$X:$X");
        // System.out.println(ref);

        // final Map<String, Object> vars = new HashMap<>();
        // for (XSSFName name : xssfnames) {
        //     vars.put(name.getNameName(), cell(workbook, name.getRefersToFormula()));
        // }

        // System.out.println(vars);
        // System.out.println(cell(workbook.getSheetAt(0), "Home!$B$2"));

        final Interpreter interpreter = new Interpreter(null);

        /*
         * Interpreter Preparation
         * 
         * 1. Add all named ranges as variables
         * This will allow <<ClaimDate>> -> 31/12/2023
         * 
         * 2. Named ranges that represent single columns or rows as indices
         * This will allow <<SDateCol>> -> 3
         * 
         * 3. Add sheets as variables. Tweak maccess to interpret sheet members 
         * as cell addresses.
         * This will allow <<shtHome.A1>> -> "Yo"
         * 
         * 4. Add dynamic sheet accessor for contract data. Unsure if effective.
         * This will allow <<shtHome[CX, 1 + 3]>> -> 21
         * 
         * 5. Add special variables that take in a range and produce information.
         * Unsure if useful.
         * This will allow <<Row.CNameCell>> -> 5
         */
    }

    private static CellReference ref(String address) {
        return new CellReference(address);
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
        final CellReference ref = new CellReference(loc);
        
        final XSSFRow row = sheet.getRow(ref.getRow());
        if (row == null) return null;
        
        return row.getCell(ref.getCol());
    }
}
