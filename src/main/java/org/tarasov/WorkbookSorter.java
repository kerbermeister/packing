package org.tarasov;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

public class WorkbookSorter {
    private static void insertHeadRow(Sheet sheet, Row headRow) {
        Row doneRow = sheet.createRow(0);
        Iterator<Cell> iterator = headRow.iterator();
        int cellNum = 0;
        while (iterator.hasNext()) {
            Cell cell = iterator.next();
            Cell doneCell = doneRow.createCell(cellNum);
            doneCell.setCellType(CellType.STRING);
            doneCell.setCellValue(cell.getStringCellValue());
            cellNum++;
        }
    }

    public static Workbook sortWorkbook(Workbook workbook) {
        Workbook doneWorkbook = new HSSFWorkbook();
        CellStyle cellStyle = doneWorkbook.createCellStyle();
        cellStyle.setDataFormat((short) 14);

        Sheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.rowIterator();
        List<Row> list = new ArrayList<Row>();

        Row headRow = sheet.getRow(0);

        int rowNum = 0;
        while(rowIterator.hasNext()) {
            if (rowNum == 0) {
                rowNum++;
                rowIterator.next();
                continue;
            }
            Row row = rowIterator.next();
            if (isRowEmpty(row))
                continue;

            list.add(row);
            rowNum++;
        }

        list.sort(new Comparator<Row>() {
            public int compare(Row o1, Row o2) {
                String o1str = o1.getCell(4).getStringCellValue();
                String o1str2 = o1.getCell(6).getStringCellValue();
                String o2str = o2.getCell(4).getStringCellValue();
                String o2str2 = o2.getCell(6).getStringCellValue();
                return  ((o1str2 + o1str).compareTo(o2str2 + o2str));
            }
        });

        Sheet doneSheet = doneWorkbook.createSheet();
        insertHeadRow(doneSheet, headRow);

        for (int i = 0; i < list.size(); i++) {
            Row doneRow = doneSheet.createRow(doneSheet.getLastRowNum()+1);
            Iterator<Cell> cellIterator = list.get(i).cellIterator();
            int cellNum = 0;
            while (cellIterator.hasNext()) {
                if (cellNum == 1) {
                    Cell cell = cellIterator.next();
                    Cell doneCell = doneRow.createCell(cellNum);
                    doneCell.setCellType(CellType.NUMERIC);
                    doneCell.setCellValue(cell.getDateCellValue());
                    doneCell.setCellStyle(cellStyle);
                    cellNum++;
                    continue;
                }
                Cell cell = cellIterator.next();
                cell.setCellType(CellType.STRING);
                Cell doneCell = doneRow.createCell(cellNum);
                doneCell.setCellType(CellType.STRING);
                doneCell.setCellValue(cell.getStringCellValue());
                cellNum++;
            }
        }
        return doneWorkbook;
    }

    private static boolean isCellEmpty(Cell cell) {
        cell.setCellType(CellType.STRING);
        if (cell == null) return true;
        if (cell.getCellType() == cell.CELL_TYPE_BLANK) return true;
        if (cell.getStringCellValue().trim().length() == 0) return true;
        return false;
    }

    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        Iterator<Cell> iterator = row.iterator();
        while (iterator.hasNext()) {
            Cell cell = iterator.next();
            if (!isCellEmpty(cell))
                return false;
        }
        return true;
    }
}
