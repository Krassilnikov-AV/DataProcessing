
package poi.readExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ReadExcelFiles {
    
    FileInputStream fis;
    String nameFile;
    Workbook wb;
    String result;
    
    public void readFiles() throws FileNotFoundException, IOException {
    nameFile = "D:/LEARNING/JAVA_DEVELOP/MyProject/Repositories/ApplicExcelPars/test.xls";
        fis = new FileInputStream(nameFile);
 //       fis = new FileInputStream("WorkOnTheCar.xlsx");
        wb = new HSSFWorkbook(fis);
        result = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        System.out.println(result);
        fis.close();
    }       
}
