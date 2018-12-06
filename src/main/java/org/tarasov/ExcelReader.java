package org.tarasov;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;

public class ExcelReader {
    private int mainSheetIndex;

    public ExcelReader() {
    }

    public ExcelReader(int mainSheetIndex) {
        this.mainSheetIndex = mainSheetIndex;
    }

    public Iterator<Row> getExcelList(Workbook workBook) {
        return workBook.getSheetAt(mainSheetIndex).rowIterator();
    }
}
