package poi.readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

public class Person {

    private String name; //ФИО
    private String address; //Адрес
    private String phoneNumber; //Номер телефона

    public String getName() {
        return name;
    }

    public String getAddress() {
        return address;
    }

    public String getPhoneNumber() {
        return phoneNumber;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public void setPhoneNumber(String phoneNumber) {
        this.phoneNumber = phoneNumber;
    }

    public static final int NAME_COLUMN_NUMBER = 0; //ФИО
    public static final int ADDRESS_COLUMN_NUMBER = 1; //Адрес
    public static final int PHONE_COLUMN_NUMBER = 2; //Телефон
    
   // String path = "ContactPerson.xls";

    public List<Person> getContacts(String path) throws IOException {
        List<Person> contacts = new ArrayList<Person>(); //Создаём пустой список контактов

       File addressDB = new File(path); //Переменная path содержит путь к документу в ФС
//       FileInputStream fis;
//          fis = new FileInputStream(path);
        POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(addressDB)); //Открываем документ
        HSSFWorkbook workBook = new HSSFWorkbook(fileSystem); // Получаем workbook
        HSSFSheet sheet = workBook.getSheetAt(0); // Проверяем только первую страницу

        Iterator<Row> rows = sheet.rowIterator(); // Перебираем все строки

        // Пропускаем "шапку" таблицы
        if (rows.hasNext()) {
            rows.next();
        }

        // Перебираем все строки начиная со второй до тех пор, пока документ не закончится 
        while (rows.hasNext()) {
            HSSFRow row = (HSSFRow) rows.next();
            //Получаем ячейки из строки по номерам столбцов
            HSSFCell nameCell = row.getCell(NAME_COLUMN_NUMBER); //ФИО
            HSSFCell addressCell = row.getCell(ADDRESS_COLUMN_NUMBER); //Адрес
            HSSFCell phoneCell = row.getCell(PHONE_COLUMN_NUMBER); //Номер телефона
            // Если в первом столбце нет данных, то контакт не создаём 
            if (nameCell != null) {
                Person person = new Person();
                person.setName(nameCell.getStringCellValue()); //Получаем строковое значение из ячейки

                person.setAddress(""); //Адрес может не быть задан
                if (addressCell != null && !"".equals(addressCell.getStringCellValue())) {
                    person.setAddress(addressCell.getStringCellValue()); //Адрес - строка
                }

                person.setPhoneNumber(""); //Телефон тоже может не быть задан
                if (phoneCell != null && !"".equals(phoneCell.getStringCellValue())) {
                    person.setPhoneNumber(phoneCell.getStringCellValue()); // Телефон - тоже строка
                }

                contacts.add(person); //Добавляем контакт в список
            }
        }
        return contacts;
    }

}
