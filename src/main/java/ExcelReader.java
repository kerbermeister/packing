import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;

public class ExcelReader {

    public static Iterator<Row> getExcelList(Workbook workBook) {
        return workBook.getSheetAt(0).rowIterator();
    }
}
