package shuffler.model;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import javafx.scene.control.Alert;
import javafx.scene.control.TextArea;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @Usman Abubakar {usmanabubakar0014@yahoo.com}
 */
public class Xlsx {
    
    private ArrayList<ArrayList<String>> data;

    public void setData (String filePath, int grouping) throws InvalidFormatException, IOException {
        OPCPackage pkg = OPCPackage.open(new File(filePath));
        try (XSSFWorkbook workbook = new XSSFWorkbook(pkg)) {
            Sheet workSheet = workbook.getSheetAt(0);
            
            ArrayList<ArrayList<String>> store = new ArrayList<>();
            ArrayList<String> record = new ArrayList<>();
            
            for (Row row : workSheet) {
                for (Cell cell : row) {
                    record.add(cell.toString());
                }
                store.add(record);
                record = new ArrayList<>();
            }
            
            ArrayList<String> Header = store.get(0);
            ArrayList<String> mainHeader = new ArrayList<>();
            Header.forEach((o) -> {
                mainHeader.add((String) o);
            });
            store.remove(0);
            
            // Create empty list
            ArrayList<String> emptyList = Header;
            Collections.fill(emptyList, "");
            Collections.shuffle(store);
            
            int counter = 0;
            for (int index = 0; index < store.size(); index++) {
                counter++;
                if (counter % grouping == 0) {
                    store.add(index, emptyList);
                    counter = 0;
                }
            }
            
            store.add(0, mainHeader);
            store.add(1, emptyList);
            data = store;
        } catch (Exception ex) {
            Alert exceptionAlert = new Alert(Alert.AlertType.ERROR);
            exceptionAlert.setTitle("File Reading Error");
            exceptionAlert.setHeaderText("Unable to read file");
            exceptionAlert.setContentText("Sorry unable to read file uploaded. The file uploaded might contain invalid text format or graph");
            TextArea errorText = new TextArea(ex.getMessage());
            exceptionAlert.getDialogPane().setExpandableContent(errorText);
            exceptionAlert.showAndWait();
        }
    }
    
    public void writeData (String filename) throws FileNotFoundException, IOException {
        XSSFWorkbook randomXlsx = new XSSFWorkbook();
        XSSFSheet randomXlsxSheet = randomXlsx.createSheet("Random Sheet");
        XSSFRow randomXlsxRow;
        int randomRowId = 0;
        for (ArrayList<String> randomData : data) {
            randomXlsxRow = randomXlsxSheet.createRow(randomRowId++);
            int randomCellId = 0;
            for (String randomCellValue : randomData) {
                Cell cell = randomXlsxRow.createCell(randomCellId++);
                cell.setCellValue(randomCellValue);
            }
        }
        try (FileOutputStream out = new FileOutputStream(new File(filename))) {
            randomXlsx.write(out);
            Alert successAlert = new Alert(Alert.AlertType.INFORMATION);
            successAlert.setTitle("Success");
            successAlert.setContentText("Done ! Check out Documents / Shuffled Excel Files for the shuffled document");
            successAlert.showAndWait();
        } catch(Exception ex) {
            Alert exceptionAlert = new Alert(Alert.AlertType.ERROR);
            exceptionAlert.setTitle("File Writing Error");
            exceptionAlert.setHeaderText("Unable to write file");
            exceptionAlert.setContentText("Sorry unable to write randomized data to new file.");
            TextArea errorText = new TextArea(ex.getMessage());
            exceptionAlert.getDialogPane().setExpandableContent(errorText);
            exceptionAlert.showAndWait();
        }
    }
}