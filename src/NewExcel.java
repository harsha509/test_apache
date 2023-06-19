import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class NewExcel {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a new sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        // Headers array
        String[] headers = {"project", "Test Case", "Type", "Parent", "Assignee", "priority", "labels", "seal ld", "description"};

        // Create the header row
        Row headerRow = sheet.createRow(0);

        // Write headers into cells
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Read data from another Excel file
        try (FileInputStream fileInputStream = new FileInputStream("input.xlsx")) {
            Workbook inputWorkbook = WorkbookFactory.create(fileInputStream);
            Sheet inputSheet = inputWorkbook.getSheetAt(0);

            StringBuilder stepsString = new StringBuilder();
            stepsString.append("||Steps||Description||Expected Result||\n");


            // Iterate over the rows in the input sheet
            for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
                Row inputRow = inputSheet.getRow(i);
                Row newRow = sheet.createRow(i + 1);  // Create the row only once

                // Read the value from the 4th column (index 3) in the input sheet
                Cell inputValueCell = inputRow.getCell(4);
                if (inputValueCell != null && inputValueCell.getCellType() == CellType.STRING) {
                    String projectValue = inputValueCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(0);
                    outputValueCell.setCellValue(projectValue);
                }

                // Read the value from the 1st column (index 0) in the input sheet
                Cell testCaseCell = inputRow.getCell(0);
                if (testCaseCell != null && testCaseCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testCaseCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(1);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // type
                Cell testTypeCell = inputRow.getCell(5);
                if (testTypeCell != null && testTypeCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testTypeCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(2);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // parent
                Cell testParentCell = inputRow.getCell(6);
                if (testParentCell != null && testParentCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testParentCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(3);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // assignee
                Cell testAssigneeCell = inputRow.getCell(7);
                if (testAssigneeCell != null && testAssigneeCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testAssigneeCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(4);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // priority
                Cell testPriorityCell = inputRow.getCell(8);
                if (testPriorityCell != null && testPriorityCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testPriorityCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(5);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // labels
                Cell testlabelsCell = inputRow.getCell(9);
                if (testlabelsCell != null && testlabelsCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testlabelsCell.getStringCellValue();


                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(6);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // sealid
                Cell testSealidCell = inputRow.getCell(10);
                if (testSealidCell != null && testSealidCell.getCellType() == CellType.NUMERIC) {
                    Number testCaseValue = testSealidCell.getNumericCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(7);
                    outputValueCell.setCellValue((Double) testCaseValue);
                }


                // description
                Cell testStepsCell = inputRow.getCell(1);
                Cell  testDescCell= inputRow.getCell(2);
                Cell testExoectedCell = inputRow.getCell(3);
                if (
                        testStepsCell != null && testStepsCell.getCellType() == CellType.STRING &&
                        testDescCell != null && testDescCell.getCellType() == CellType.STRING &&
                        testExoectedCell != null && testExoectedCell.getCellType() == CellType.STRING
                ) {


                    stepsString.append(testStepsCell.getStringCellValue())
                            .append("|")
                            .append(testDescCell.getStringCellValue())
                            .append("|")
                            .append(testExoectedCell.getStringCellValue())
                            .append("\n");

                    Cell outputValueCell = newRow.createCell(8);
                    outputValueCell.setCellValue(stepsString.toString());
                }
            }


        } catch (IOException e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
        }

        // Save the workbook to a file
        try (FileOutputStream outputStream = new FileOutputStream("output.xlsx")) {
            workbook.write(outputStream);
            System.out.println("Excel file created successfully!");
        } catch (IOException e) {
            System.err.println("Error creating Excel file: " + e.getMessage());
        } finally {
            // Close the workbook
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
