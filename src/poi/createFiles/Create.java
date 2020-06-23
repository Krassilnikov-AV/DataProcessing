package poi.createFiles;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class Create {

    public void create() throws FileNotFoundException, IOException {
        // создание экземпляра книг
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("List");  // 
        // cоздание строки
        Row row = sheet.createRow(1);
        //создание ячейки
        Cell cell = row.createCell(2);
        cell.setCellValue("cell");
// создание листа с белебердой через класс WorkbookUtil
        Sheet sheet1 = wb.createSheet(WorkbookUtil.createSafeSheetName("!%^&@^$@^"));  // 
        
        Row row1 = sheet1.createRow(1);
        Cell cell1 = row1.createCell(2);
        cell1.setCellValue("Ok!");
        
        FileOutputStream fos = new FileOutputStream("test.xls");  // создание файла для записи книги

        wb.write(fos);
        fos.close();
    }
}
