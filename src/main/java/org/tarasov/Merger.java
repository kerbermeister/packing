package org.tarasov;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


class Merger {
    private File[] getFilesToMerge(String path) {
        File directory = new File(path);
        return directory.listFiles();
    }

    private Workbook merge(File[] files) throws IOException {
        Workbook mergedWorkbook = new HSSFWorkbook();

        for (File file : files) {
            if (file.getName().contains("xls")) {
                List<Sheet> sheets = getSheets(file);

                for (Sheet sheet : sheets) {
                    insertSheet(mergedWorkbook, sheet);
                }
            }
        }

        return mergedWorkbook;
    }

    private List<Sheet> getSheets(File file) throws IOException {
        FileInputStream fio = new FileInputStream(file);
        Workbook workbook = new HSSFWorkbook(fio);
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        List<Sheet> sheetList = new ArrayList<>();

        while (sheetIterator.hasNext()) {
            sheetList.add(sheetIterator.next());
        }
        return sheetList;
    }

    private void insertSheet(Workbook workbook, Sheet sheet) {
        int numberOfSheets = workbook.getNumberOfSheets();
        Sheet insertedSheet = workbook.createSheet("(" + numberOfSheets + ")" + sheet.getSheetName());

        Iterator<Row> rowIterator = sheet.rowIterator();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat((short) 14);

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Row doneRow = insertedSheet.createRow(row.getRowNum());

            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                Cell doneCell = doneRow.createCell(cell.getColumnIndex());

                if (cell.getCellType() == 0) {
                    doneCell.setCellType(cell.getCellType());
                    doneCell.setCellValue(cell.getDateCellValue());
                    doneCell.setCellStyle(cellStyle);
                } else {
                    doneCell.setCellValue(cell.getStringCellValue());
                }
            }
        }
    }

    private void saveMergedFile(Workbook workbook, String pathToSave) throws FileNotFoundException, IOException {
        File file = new File(pathToSave + "!merged.xls");
        FileOutputStream fio = new FileOutputStream(file);
        workbook.write(fio);
        fio.close();
    }

    public void process(String filesToMergePath, String pathToSave) throws IOException {
        File[] filesToMerge = getFilesToMerge(filesToMergePath);
        Workbook merged = merge(filesToMerge);
        saveMergedFile(merged, pathToSave);
    }
}
