import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

//import statements
public class ReadExcel
{
    public static void main(String[] args)
    {
        try (Workbook workbook = WorkbookFactory.create(new File("input.xlsx"))) {
            Sheet sheet = workbook.getSheet("Sheet1");


            Workbook workbook1 = WorkbookFactory.create(true);

            // Create a new worksheet
            Sheet sheet1 = workbook.createSheet("Sheet1");

            // Create default headers
            Row headerRow1 = sheet1.createRow(0);
            headerRow1.createCell(0).setCellValue("Project");
            headerRow1.createCell(1).setCellValue("Testcase");




            Row headerRow= sheet1.getRow(0);
            // Iterate through each cell in the header row
            for (Cell cell : headerRow) {
                String columnHeader = cell.getStringCellValue();

                // Check if the column header matches "Project" or "Test Case"
                if (columnHeader.equals("Project")) {
                    // Read the values in the column
                    for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                        Row dataRow = sheet.getRow(rowIndex);
                        Row row = sheet.createRow(rowIndex);

                        Cell dataCell = dataRow.getCell(cell.getColumnIndex());
                        String cellValue = dataCell.getStringCellValue();
                        row.createCell(0).setCellValue("cellValue");
                        System.out.println(cellValue);
                    }
                } else if (columnHeader.equals("Testcase")) {

                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}