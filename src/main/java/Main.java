import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) throws InvalidFormatException, IOException {
        String path;
        Scanner scanner = new Scanner(System.in);
        System.out.println("enter file path: ");
        path = scanner.nextLine();
        Workbook doneWorkbook;

        File directory = new File(path);

        File[] files = directory.listFiles();

        for (File file : files) {
            FileInputStream fileInputStream = new FileInputStream(file);
            Workbook workbook = new HSSFWorkbook(fileInputStream);
            Iterator<Row> rows = ExcelReader.getExcelList(workbook);
            doneWorkbook = new HSSFWorkbook();
            Sheet doneSheet = doneWorkbook.createSheet();

            int rowNum = 0;
            Row firstDoneRow = doneSheet.createRow(0);
            firstDoneRow.createCell(0).setCellValue("Model");
            firstDoneRow.createCell(1).setCellValue("Quantity");
            firstDoneRow.createCell(2).setCellValue("Container");
            firstDoneRow.createCell(3).setCellValue("Date");
            firstDoneRow.createCell(4).setCellValue("Serial number");
            firstDoneRow.createCell(5).setCellValue("Bar Code");
            firstDoneRow.createCell(6).setCellValue("Version");
            firstDoneRow.createCell(7).setCellValue("Color");
            firstDoneRow.createCell(8).setCellValue("Order");

            while (rows.hasNext()) {
                if (rowNum == 0){
                    rows.next();
                    rowNum++;
                    continue;
                }

                Row row = rows.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Row doneRow = doneSheet.createRow(rowNum);

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    cell.setCellType(CellType.STRING);
                }

                for (int i = 0; i < 9; i++) {
                    doneRow.createCell(i).setCellType(CellType.STRING);
                }

                doneRow.getCell(0).setCellValue(row.getCell(6).getStringCellValue());
                doneRow.getCell(1).setCellValue(row.getCell(0).getStringCellValue());
                doneRow.getCell(2).setCellValue(row.getCell(5).getStringCellValue());
                doneRow.getCell(3).setCellValue(row.getCell(1).getStringCellValue());
                doneRow.getCell(4).setCellValue(row.getCell(2).getStringCellValue());
                doneRow.getCell(5).setCellValue(row.getCell(3).getStringCellValue());
                doneRow.getCell(6).setCellValue(row.getCell(10).getStringCellValue());
                doneRow.getCell(7).setCellValue(row.getCell(4).getStringCellValue());
                doneRow.getCell(8).setCellValue(row.getCell(7).getStringCellValue());

                rowNum++;
            }

            File outputFile = new File("C:\\Users\\tarasov.a\\Desktop\\files\\1.xls");
            FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
            doneWorkbook.write(fileOutputStream);
            fileOutputStream.close();
            fileInputStream.close();
        }
    }
}
