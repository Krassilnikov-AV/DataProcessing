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

public class CreateExcelFiles {

    public void createTest() throws FileNotFoundException, IOException {
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
    // создание файла для работы с техсостоянием и ремонтом авто
     public void createAuto() throws FileNotFoundException, IOException {
        // создание экземпляра книг
        Workbook wbAuto = new HSSFWorkbook();
        Sheet sheet = wbAuto.createSheet("List");  
    
        FileOutputStream fosAuto = new FileOutputStream("WorkOnTheCar.xls");
        wbAuto.write(fosAuto);
        fosAuto.close();
    }
     // создание файла для работы с английским текстом dictionary 
       public void createDictionary() throws FileNotFoundException, IOException {
        // создание экземпляра книг
        Workbook wbDic = new HSSFWorkbook();
        Sheet sheet = wbDic.createSheet("AN-RUS");  
    
        FileOutputStream fos = new FileOutputStream("EngRusDict.xls");
        wbDic.write(fos);
        fos.close();
    }
}
