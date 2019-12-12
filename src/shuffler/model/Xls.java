package shuffler.model;

/**
 *
 * @Usman Abubakar {usmanabubakar0014@yahoo.com}
 */

import java.io.File;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;

public class Xls {
    public Xls (String filePath, int grouping) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(new File(filePath));
        HSSFWorkbook workbook = new HSSFWorkbook(fs.getRoot(), true);
        HSSFSheet xlsSheet = workbook.getSheetAt(0);
        workbook.close();
    }
}
