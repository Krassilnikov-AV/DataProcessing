package poi.createFiles;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateContact {

    // создание файла для работы с данными Person
    public void createPersonFile() throws FileNotFoundException, IOException {
        // создание экземпляра книг:
        // для xls
        // Workbook wbDic = new HSSFWorkbook();
      

        //      Sheet sheet = wbDic.createSheet("Contact");
        //        FileOutputStream fos = new FileOutputStream("ContactPerson.xls");
//        wbDic.write(fos);
//        fos.close();
  // для xlsx с применением библиотек:
  // poi-ooxml-3.11-20141221.jar; xmlbeans-2.6.0.jar; poi-ooxml-scemas-3.11-20141221.jar
        Workbook wbDicx = new XSSFWorkbook();
        Sheet sheetx = wbDicx.createSheet("Contactx");

        FileOutputStream fosx = new FileOutputStream("ContactPersonx.xlsx");
        wbDicx.write(fosx);
        fosx.close();

    }
}
