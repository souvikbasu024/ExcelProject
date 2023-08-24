package souvik;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DeleteAndShiftWithFormating {

    public static void main(String[] args) {
        // Get input folder path and column name from the user
        String folderPath = "C:\\Users\\souvi\\OneDrive\\Documents\\aps-oracle-cloud-erp\\src\\main\\resources\\Data";
        String columnNameToDelete = "MYNAME";

        // Get a list of Excel files in the folder
        File folder = new File(folderPath);
        File[] excelFiles = folder.listFiles((dir, name) -> name.endsWith(".xlsx"));

        if (excelFiles == null || excelFiles.length == 0) {
            System.out.println("No Excel files found in the specified folder.");
            return;
        }

        // Process each Excel file
        for (File excelFile : excelFiles) {
            try {
                boolean columnDeleted = processExcelFile(excelFile, columnNameToDelete);
                if (columnDeleted) {
                    System.out.println("Column deleted from " + excelFile.getName());
                } else {
                    System.out.println("Column not found in " + excelFile.getName());
                }
            } catch (IOException e) {
                System.out.println("Error processing " + excelFile.getName() + ": " + e.getMessage());
            }
        }

        System.out.println("Column deletion process completed.");
    }

    private static boolean processExcelFile(File excelFile, String columnNameToDelete) throws IOException {
        FileInputStream inputStream = new FileInputStream(excelFile);
        Workbook workbook = new XSSFWorkbook(inputStream);
        boolean columnDeleted = false;

        // Process each sheet in the workbook
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Row headerRow = sheet.getRow(0);
            int columnIndexToDelete = -1;

            // Find the column index to delete
            for (Cell cell : headerRow) {
                if (cell.getStringCellValue().equalsIgnoreCase(columnNameToDelete)) {
                    columnIndexToDelete = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndexToDelete != -1) {
                // Delete the column and shift cells to the left
                for (Row row : sheet) {
                    Cell cellToDelete = row.getCell(columnIndexToDelete);
                    if (cellToDelete != null) {
                        row.removeCell(cellToDelete);
                    }
                    shiftCellsToLeft(row, columnIndexToDelete);
                }
                columnDeleted = true;
            }
        }

        if (columnDeleted) {
            // Write the changes back to the file
            FileOutputStream outputStream = new FileOutputStream(excelFile);
            workbook.write(outputStream);
            outputStream.close();
        }

        // Close resources
        inputStream.close();
        workbook.close();

        return columnDeleted;
    }

    private static void shiftCellsToLeft(Row row, int columnIndex) {
        for (int i = columnIndex + 1; i <= row.getLastCellNum(); i++) {
            Cell cellToShift = row.getCell(i);
            if (cellToShift != null) {
                Cell newCell = row.createCell(i - 1, cellToShift.getCellType());
                copyCellStyling(cellToShift, newCell);
                copyCellValue(cellToShift, newCell);
                row.removeCell(cellToShift);
            }
        }
    }

    private static void copyCellStyling(Cell sourceCell, Cell targetCell) {
        targetCell.setCellStyle(sourceCell.getCellStyle());
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell) {
        CellType cellType = sourceCell.getCellType();
        switch (cellType) {
            case NUMERIC:
                targetCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
               
        }
    }
}
