import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


 public class Main {

     private static final String FILE_NAME = "testik.xls";

     public static void main(String[] args) {

         XSSFWorkbook workbook = new XSSFWorkbook();
         XSSFSheet sheet = workbook.createSheet("Норникель");
         Object[][] datatypes = {
                 {"Металл", "Доля на рынке", "Год"},
                 {"Палладий", "40 %", 2019},
                 {"Никель", "12 %", 2019},
                 {"Кобальт", "5 %", 2019},
                 {"Медь", "2 %", 2019},
         };

         int rowNum = 0;
         System.out.println("Creating excel");

         for (Object[] datatype : datatypes) {
             Row row = sheet.createRow(rowNum++);
             int colNum = 0;
             for (Object field : datatype) {
                 Cell cell = row.createCell(colNum++);
                 if (field instanceof String) {
                     cell.setCellValue((String) field);
                 } else if (field instanceof Integer) {
                     cell.setCellValue((Integer) field);
                 }
             }
         }
         for (int i = 0;i < 3; i++) {
             sheet.autoSizeColumn(i);
         }

         try {
             FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
             workbook.write(outputStream);
             workbook.close();
         } catch (FileNotFoundException e) {
             e.printStackTrace();
         } catch (IOException e) {
             e.printStackTrace();
         }
         System.out.println("Done");
     }
 }