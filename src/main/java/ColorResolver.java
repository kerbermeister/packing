import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;

public class ColorResolver {
    private static String pattern;

    static void init() throws FileNotFoundException , IOException {
        File excelSource = new File("C:\\Users\\tarasov.a\\Desktop\\parser\\excel.xls");
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

    static Map<String, String> colors = new TreeMap<String, String>();

    static {
//       colors.put("темно-серый", "Dark Grey");
//       colors.put("черный", "Black");
//       colors.put("белый", "White");
       try {
           init();
       } catch (FileNotFoundException e) {
           e.printStackTrace();
       } catch (IOException e) {
           e.printStackTrace();
       }
    }

    public static String resolveColor(String color) {
        String resolvedColor = colors.get(color.toLowerCase());
        if (resolvedColor == null)
            return color;
        return resolvedColor;
    }
}
