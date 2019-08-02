package merger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class Parser {
    public static void main(String[] args) throws IOException {

        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\tarasov.a\\Desktop\\pars\\ch.xls");

        Workbook workbook = new HSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = sheet.rowIterator();

        Map<String, String> dictionary = new HashMap<>();

        while (iterator.hasNext()) {
            Row row = iterator.next();

            Cell eng = row.getCell(1);
            Cell rus = row.getCell(2);
            eng.setCellType(CellType.STRING);
            rus.setCellType(CellType.STRING);
            dictionary.put(eng.getStringCellValue(), rus.getStringCellValue());
        }






        FileInputStream fileInputStream2 = new FileInputStream("C:\\Users\\tarasov.a\\Desktop\\pars\\cultraview.xls");
        Workbook done = new HSSFWorkbook();
        Sheet doneSheet = done.createSheet();


        Workbook workbook2 = new HSSFWorkbook(fileInputStream2);

        Sheet sheet2 = workbook2.getSheetAt(0);

        Iterator<Row> rowIterator = sheet2.rowIterator();

        int counter = 0;
        while (rowIterator.hasNext()) {
            if (counter == 0) {
                rowIterator.next();
                counter++;
                continue;
            }



            Row row = rowIterator.next();
            row.getCell(1).setCellType(CellType.STRING);
            String word = row.getCell(1).getStringCellValue();
            String translate;

            if (dictionary.containsKey(word)) {
                translate = dictionary.get(word);
                Row doneRow = createDoneRow(doneSheet, word, translate);
                Cell doneCell = doneRow.getCell(1);
                CellStyle cellStyle = done.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                doneCell.setCellStyle(cellStyle);

            } else {
                row.getCell(4).setCellType(CellType.STRING);
                translate = row.getCell(4).getStringCellValue();
                createDoneRow(doneSheet, word, translate);
            }
            counter++;
        }

        File output  = new File("C:\\Users\\tarasov.a\\Desktop\\pars\\done.xls");
        FileOutputStream fileOutputStream = new FileOutputStream(output);
        ((HSSFWorkbook) done).write(fileOutputStream);
        fileOutputStream.close();

    }


    public static Row createDoneRow(Sheet sheet, String eng, String rus) {
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(lastRowNum+1);
        row.createCell(0).setCellType(CellType.STRING);
        row.createCell(1).setCellType(CellType.STRING);
        row.getCell(0).setCellValue(eng);
        row.getCell(1).setCellValue(rus);
        return row;
    }
}
