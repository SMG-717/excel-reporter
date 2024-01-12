package com.forenzix.excel;

import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.forenzix.common.Pair;

public class NameReader {

    public static List<Pair<String, String>> extractNames(XSSFWorkbook workbook) {
        List<XSSFName> names = workbook.getAllNames();
        
        // for (XSSFName name : names) System.out.println(name);

        return names.stream()
            .map(n -> Pair.of(n.getNameName(), n.getRefersToFormula()))
            .collect(Collectors.toList());
    }
}
