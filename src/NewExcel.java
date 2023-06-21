import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class NewExcel {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a new sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        // Headers array
        String[] headers = {"project", "Test Case", "Type", "Parent", "Assignee", "priority", "labels", "seal ld", "description"};

        // Create the header row
        Row headerRow_ = sheet.createRow(0);

        // Write headers into cells
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow_.createCell(i);
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

            // Find the header row in the input sheet
            Row headerRow = inputSheet.getRow(0);

            // Find the column indices based on header matching
            int projectColumnIndex = -1;
            int testCaseColumnIndex = -1;
            int testTypeColumnIndex = -1;
            int testParentColumnIndex = -1;
            int testAssigneeColumnIndex = -1;
            int testPriorityColumnIndex = -1;
            int testLabelsColumnIndex = -1;
            int testSealidColumnIndex = -1;
            int testStepsColumnIndex = -1;
            int testDescColumnIndex = -1;
            int testExpectedColumnIndex = -1;

            for (Cell headerCell : headerRow) {
                String headerCellValue = headerCell.getStringCellValue().toLowerCase();

                switch (headerCellValue.toLowerCase()) {
                    case "project":
                        projectColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "testcase":
                        testCaseColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "type":
                        testTypeColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "parent":
                        testParentColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "asignee":
                        testAssigneeColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "priority":
                        testPriorityColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "label":
                        testLabelsColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "sealid":
                        testSealidColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "step name":
                        testStepsColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "step description":
                        testDescColumnIndex = headerCell.getColumnIndex();
                        break;
                    case "expected result":
                        testExpectedColumnIndex = headerCell.getColumnIndex();
                        break;
                }
            }

            if (projectColumnIndex == -1 || testCaseColumnIndex == -1 || testTypeColumnIndex == -1 ||
                    testParentColumnIndex == -1 || testAssigneeColumnIndex == -1 ||
                    testPriorityColumnIndex == -1 || testLabelsColumnIndex == -1 ||
                    testSealidColumnIndex == -1 || testStepsColumnIndex == -1 ||
                    testDescColumnIndex == -1 || testExpectedColumnIndex == -1) {
                System.err.println("One or more headers not found in the input sheet.");
                return; // or handle the error as per your requirement
            }
            StringBuilder subStepString = new StringBuilder();
            List<StringBuilder> stepList = new ArrayList<>();
            for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
                Row inputRow = inputSheet.getRow(i);


                Row newRow = sheet.createRow(i + 1);  // Create the row only once

                // Read the values based on the column indices
                // project
                Cell projectCell = inputRow.getCell(projectColumnIndex);
                if (projectCell != null && projectCell.getCellType() == CellType.STRING) {
                    String projectValue = projectCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(0);
                    outputValueCell.setCellValue(projectValue);
                }

                // testcase
                Cell testCaseCell = inputRow.getCell(testCaseColumnIndex);
                if (testCaseCell != null && testCaseCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testCaseCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(1);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // type
                Cell testTypeCell = inputRow.getCell(testTypeColumnIndex);
                if (testTypeCell != null && testTypeCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testTypeCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(2);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // parent
                Cell testParentCell = inputRow.getCell(testParentColumnIndex);
                if (testParentCell != null && testParentCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testParentCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(3);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // assignee
                Cell testAssigneeCell = inputRow.getCell(testAssigneeColumnIndex);
                if (testAssigneeCell != null && testAssigneeCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testAssigneeCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(4);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // priority
                Cell testPriorityCell = inputRow.getCell(testPriorityColumnIndex);
                if (testPriorityCell != null && testPriorityCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testPriorityCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(5);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // labels
                Cell testLabelsCell = inputRow.getCell(testLabelsColumnIndex);
                if (testLabelsCell != null && testLabelsCell.getCellType() == CellType.STRING) {
                    String testCaseValue = testLabelsCell.getStringCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(6);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // sealid
                Cell testSealidCell = inputRow.getCell(testSealidColumnIndex);
                if (testSealidCell != null && testSealidCell.getCellType() == CellType.NUMERIC) {
                    double testCaseValue = testSealidCell.getNumericCellValue();

                    // Write the value to the corresponding column in the new sheet
                    Cell outputValueCell = newRow.createCell(7);
                    outputValueCell.setCellValue(testCaseValue);
                }

                // steps, description, expected result
                Cell testStepsCell = inputRow.getCell(testStepsColumnIndex);
                Cell testDescCell = inputRow.getCell(testDescColumnIndex);
                Cell testExpectedCell = inputRow.getCell(testExpectedColumnIndex);
                if (testStepsCell != null && testStepsCell.getCellType() == CellType.STRING &&
                        testDescCell != null && testDescCell.getCellType() == CellType.STRING &&
                        testExpectedCell != null && testExpectedCell.getCellType() == CellType.STRING) {
                    if (testCaseCell != null) {
                        stepList.add(subStepString);
                        subStepString = new StringBuilder();
                    }
                    subStepString.append(testStepsCell.getStringCellValue())
                            .append("|")
                            .append(testDescCell.getStringCellValue())
                            .append("|")
                            .append(testExpectedCell.getStringCellValue())
                            .append("\n");
                }
            }
            if (!subStepString.toString().equals("")) {
                stepList.add(stepsString.append(subStepString));
            }
            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                Row sheetRow = sheet.getRow(i);
                if (sheetRow != null && sheetRow.getCell(1) != null) {
                    for (int j = 1; j < stepList.size(); j++) {
                        Cell outputValueCell = sheetRow.createCell(8);
                        outputValueCell.setCellValue(stepList.get(j).toString());
                    }
                }
                else{
                    if(sheetRow != null) {
                        sheet.removeRow(sheetRow);
                    }
                }
            }
            // Iterate over the rows in the input sheet
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
