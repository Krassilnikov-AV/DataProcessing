
package poi.main;

import java.io.FileNotFoundException;
import java.io.IOException;

import poi.createFiles.CreateExcelFiles;

public class Main {

     public static void main(String[] args) throws FileNotFoundException, IOException {
     
         CreateExcelFiles cf = new CreateExcelFiles();
         cf.create();
    }
    
}
