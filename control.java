package souvik;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;

public class control extends Operation {

	@FXML
	private Button btnDelete;

	@FXML
	private Button btnDelete1;

	@FXML
	private Button btnreplace;

	@FXML
	private TextField tfCellvalue;

	@FXML
	private TextField tfPath;

	@FXML
	private TextField tfreplace;

	// String FilePath =tfPath.getText();
	// String Cellvalue = tfCellvalue.getText();
	// String ReplaceValue =tfreplace.getText();
	@FXML
	void btnDeleteClick(ActionEvent event) {
		String Path = tfPath.getText();
		String Cellvalue = tfCellvalue.getText();
		Operation.DeleteColumn(Path, Cellvalue);
	}

	@FXML
	void btnDeleteValueClick(ActionEvent event) {
		String Path = tfPath.getText();
		String Colheader = tfCellvalue.getText();
		Operation.Deletevalue(Path, Colheader);
	}

	@FXML
	void btnreplaceClick(ActionEvent event) {
		String Path = tfPath.getText();
		String Cellvalue = tfCellvalue.getText();
		String ReplaceValue = tfreplace.getText();
		Operation.replace(Path, Cellvalue, ReplaceValue);
	}

}
