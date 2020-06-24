
package poi.createFiles;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class CreateContact {
    // создание файла для работы с данными Person
       public void createPersonFile() throws FileNotFoundException, IOException {
        // создание экземпляра книг
        Workbook wbDic = new HSSFWorkbook();
        Sheet sheet = wbDic.createSheet("Contact");  
    
        FileOutputStream fos = new FileOutputStream("ContactPerson.xls");
        wbDic.write(fos);
        fos.close();
    }
}
