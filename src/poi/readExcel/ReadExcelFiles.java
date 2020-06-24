package poi.readExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ReadExcelFiles {

    FileInputStream fis;
    String nameFile;
    Workbook wb;
    String result;

    public void readFiles() throws FileNotFoundException, IOException {
        nameFile = "EngRusDict.xls";
        fis = new FileInputStream(nameFile);
        //       fis = new FileInputStream("WorkOnTheCar.xlsx");
        wb = new HSSFWorkbook(fis);
// чтение данных итератором
        Sheet sheet = wb.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();

                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + ".....";
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        result += "[" + cell.getNumericCellValue() + ")";
                        break;

                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }

//        result = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        System.out.println(result);
        fis.close();
    }

    // метод для самостоятельного определения формата считываемого значения
    public static String getCelltext(Cell cell) {

        String result = "";

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula().toString();
                break;
            default:
                break;
        }
        return result;
    }
}
