import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class main {
    public static void main(String[] args) {
        try {
            // Отваряне на XLSX файл
            FileInputStream file = new FileInputStream(new File("src/main/java/example.xlsx"));

            // Създаване на workbook обект
            Workbook workbook = WorkbookFactory.create(file);

            // Обхождане на всички листове във файла
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                // Избиране на текущия лист
                Sheet sheet = workbook.getSheetAt(i);
                System.out.println("Reading sheet: " + sheet.getSheetName());

                // Обхождане на редовете на листа
                for (Row row : sheet) {
                    // Обхождане на клетките в реда
                    for (Cell cell : row) {
                        // Печат на съдържанието на клетките
                        switch (cell.getCellType()) {
                            case STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue() + "\t");
                                break;
                            case BOOLEAN:
                                System.out.print(cell.getBooleanCellValue() + "\t");
                                break;
                            default:
                                System.out.print("\t");
                        }
                    }
                    System.out.println();
                }
            }

            // Затваряне на файла
            file.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
