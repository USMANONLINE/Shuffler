package shuffler;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Optional;
import java.util.ResourceBundle;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextInputDialog;
import javafx.stage.FileChooser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import shuffler.model.Xlsx;

public class FXMLDocumentController implements Initializable {

    @FXML
    private Button uploadBtn;
    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // Creates the working directory
        File workingDir = new File("C:\\Users\\User\\Documents\\Shuffled Excel Files");
        if (!workingDir.exists())
            workingDir.mkdir();
    }        

    @FXML
    private void handleFileUpload(ActionEvent event) throws IOException, InvalidFormatException {
        FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel files (*.xls)", "*.xlsx");
        fileChooser.getExtensionFilters().add(extFilter);
        File selectedFile = fileChooser.showOpenDialog(null);
        int groupInput = 0;
        
        if (selectedFile != null) {
            TextInputDialog inputDialog = new TextInputDialog();
            inputDialog.setTitle("Grouping");
            inputDialog.setHeaderText("Enter the number of people per group after shuffling");
            inputDialog.setContentText("Number of people");
            uploadBtn.setText("loading...");
            Optional<String> grouping = inputDialog.showAndWait();
            
            if (grouping.isPresent()) {
                groupInput = Integer.parseInt(grouping.get());
            } else {
                Alert invalidInput = new Alert(Alert.AlertType.ERROR);
                invalidInput.setTitle("Invalid Grouping");
                invalidInput.setContentText("Please enter a valid grouping number");
                invalidInput.setHeaderText(null);
                invalidInput.showAndWait();
            }
        }
        groupInput += 1;
        
       Xlsx xlsxDoc = new Xlsx();
       xlsxDoc.setData(selectedFile.getAbsolutePath(), groupInput);
       xlsxDoc.writeData("C:\\Users\\User\\Documents\\Shuffled Excel Files\\" +selectedFile.getName());
       uploadBtn.setText("Done");
    }
}
