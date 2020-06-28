package poi.main;

import java.io.FileNotFoundException;
import java.io.IOException;
import poi.createFiles.CreateContact;

import poi.createFiles.*;
import poi.readExcel.*;

public class Main {

    public static void main(String[] args) throws FileNotFoundException, IOException {

//         CreateExcelFiles cf = new CreateExcelFiles();
//         cf.createTest();
//         CreateExcelFiles cfauto = new CreateExcelFiles();
//         cfauto.createAuto();
//        CreateExcelFiles cfauto = new CreateExcelFiles();
//        cfauto.createDictionary();
        CreateContact cp = new CreateContact();
        cp.createPersonFile();
        System.out.println("file content:");
// file created

//        ReadExcelFiles rf = new ReadExcelFiles();
//        rf.readFiles();

//        Person pers = new Person();
//        pers.getContacts();
    }
}
