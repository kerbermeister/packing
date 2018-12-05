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
            Iterator<Row> excelList = ExcelReader.getExcelList(workbook);
            doneWorkbook = new HSSFWorkbook();
            Sheet doneSheet = doneWorkbook.createSheet();

            int rowNum = 0;
            while (excelList.hasNext()) {
                if (rowNum == 0){
                    excelList.next();
                    rowNum++;
                    continue;
                }

                Row row = excelList.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Row doneRow = doneSheet.createRow(rowNum);

                int cellNum = 0;
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    cell.setCellType(CellType.STRING);
                    Cell doneCell = doneRow.createCell(cellNum);
                    doneCell.setCellType(CellType.STRING);
                    doneCell.setCellValue(cell.getStringCellValue());
                    cellNum++;
                }
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
