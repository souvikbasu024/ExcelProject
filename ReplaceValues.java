package souvik;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReplaceValues {
    public static void main(String[] args) {
    	
       
        String folderPath ="C:\\Users\\souvi\\OneDrive\\Documents\\aps-oracle-cloud-erp\\src\\main\\resources\\Data";
        String searchValue ="SOUVIKBASU" ;
        String replaceValue = "Souvik_Basu012345";

        File folder = new File(folderPath);
        if (!folder.exists() || !folder.isDirectory()) {
            System.out.println("Invalid folder path.");
            return;
        }

        File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx"));
        if (files == null || files.length == 0) {
            System.out.println("No Excel files found in the folder.");
            return;
        }

        for (File file : files) {
            try {
                FileInputStream inputStream = new FileInputStream(file);
                Workbook workbook = new XSSFWorkbook(inputStream);
                boolean modified = false;

                for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                    Sheet sheet = workbook.getSheetAt(sheetIndex);
                    boolean sheetModified = false;

                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            if (cell.getCellType() == CellType.STRING &&
                                cell.getStringCellValue().equals(searchValue)) {
                                cell.setCellValue(replaceValue);
                                sheetModified = true;
                                modified = true;
                            }
                        }
                    }

                    if (sheetModified) {
                        System.out.println("Replaced values in sheet " + sheet.getSheetName() +
                                           " of file: " + file.getName());
                    }
                }

                inputStream.close();

                if (modified) {
                    FileOutputStream outputStream = new FileOutputStream(file);
                    workbook.write(outputStream);
                    workbook.close();
                    outputStream.close();
                    System.out.println("Saved changes to file: " + file.getName());
                }

            } catch (IOException e) {
                System.out.println("Error processing file: " + file.getName());
                e.printStackTrace();
            }
        }
    }
}


