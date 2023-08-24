package souvik;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class DeleteOnlyCellValue {
    public static void main(String[] args) {
        String folderPath = "C:\\Users\\souvi\\OneDrive\\Documents\\aps-oracle-cloud-erp\\src\\main\\resources\\Data";
        String columnName = "MYNAME";

        File folder = new File(folderPath);
        File[] excelFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));

        if (excelFiles != null) {
            for (File file : excelFiles) {
                try {
                    FileInputStream fis = new FileInputStream(file);
                    Workbook workbook = new XSSFWorkbook(fis);

                    boolean fileModified = false; // To track if any changes were made to the file

                    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                        Sheet sheet = workbook.getSheetAt(sheetIndex);

                        // Find the column index by searching for the column name in the header row
                        int columnIndex = -1;
                        Row headerRow = sheet.getRow(0);
                        for (Cell cell : headerRow) {
                            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                                columnIndex = cell.getColumnIndex();
                                break;
                            }
                        }

                        if (columnIndex != -1) {
                            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                                Row row = sheet.getRow(rowIndex);
                                if (row != null) {
                                    Cell cell = row.getCell(columnIndex);
                                    if (cell != null) {
                                        // Preserve formatting of the deleted cell
                                        CellStyle cellStyle = cell.getCellStyle();
                                        row.removeCell(cell);
                                        Cell newCell = row.createCell(columnIndex);
                                        newCell.setCellStyle(cellStyle);
                                        fileModified = true;
                                    }
                                }
                            }
                        }
                    }

                    if (fileModified) {
                        // Save the changes back to the file
                        FileOutputStream fos = new FileOutputStream(file);
                        workbook.write(fos);
                        fos.close();
                        System.out.println("File modified: " + file.getName());
                    }

                    fis.close();
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
