package org.tarasov;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class ColorResolver {
    private Map<String, String> colors = new HashMap<String, String>();
    private String patternsPath;

    public ColorResolver() {
    }

    public ColorResolver(String patternsPath)
    {
        this.patternsPath = patternsPath;
        try {
            init();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void init() throws IOException {
        File excelSource = new File(patternsPath);
        FileInputStream fileInputStream = new FileInputStream(excelSource);
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String key = row.getCell(0).getStringCellValue();
            String value = row.getCell(1).getStringCellValue();
            colors.put(key, value);
        }
        workbook.close();
        fileInputStream.close();
    }

    public String resolveColor(String color) {
        String resolvedColor = colors.get(color);
        if (resolvedColor == null)
            return null;
        return resolvedColor;
    }

    public String getPatternsPath() {
        return patternsPath;
    }

    public void setPatternsPath(String patternsPath) {
        this.patternsPath = patternsPath;
    }
}
