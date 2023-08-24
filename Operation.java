package souvik;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Operation {
	
	//this code takes input from user and  do the replace operation on specific value
	
	static void replace(String Path, String Cellvalue, String Replacevalue) {
		String folderPath = Path;
		String searchValue = Cellvalue;
		String replaceValue = Replacevalue;

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
							if (cell.getCellType() == CellType.STRING
									&& cell.getStringCellValue().equals(searchValue)) {
								cell.setCellValue(replaceValue);
								sheetModified = true;
								modified = true;
							}
						}
					}

					if (sheetModified) {
						System.out.println(
								"Replaced values in sheet " + sheet.getSheetName() + " of file: " + file.getName());
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

	/*----------------------------------------------------------------------------------------------------*/

	static void DeleteColumn(String Path, String Cellvalue) {
		String folderPath = Path;
		String columnNameToDelete = Cellvalue;

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
	/*------------------------------------------------------------------------------------------------------------------------------*/
	// Method to delete the respective values from cell by taking column header as input
	static void Deletevalue(String Path, String Colname) {

		String folderPath = Path;
		String columnName = Colname;

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