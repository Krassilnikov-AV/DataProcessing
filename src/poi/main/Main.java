package poi.main;

import java.io.FileNotFoundException;
import java.io.IOException;

import poi.createFiles.CreateExcelFiles;
import poi.readExcel.ReadExcelFiles;

public class Main {

    public static void main(String[] args) throws FileNotFoundException, IOException {

//         CreateExcelFiles cf = new CreateExcelFiles();
//         cf.createTest();
//         CreateExcelFiles cfauto = new CreateExcelFiles();
//         cfauto.createAuto();
//        CreateExcelFiles cfauto = new CreateExcelFiles();
//        cfauto.createDictionary();
        System.out.println("file content:");
// file created
         ReadExcelFiles rf = new ReadExcelFiles();
         rf.readFiles();
    }
}
