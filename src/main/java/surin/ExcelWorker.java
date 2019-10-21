package surin;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ExcelWorker {
    private final String pathToFile;
    private static final Logger log = Logger.getLogger(ExcelWorker.class);

    public ExcelWorker(String pathToFile){
        this.pathToFile = pathToFile;
    }
    public XSSFWorkbook readWorkbook() {
        try {
            ZipSecureFile.setMinInflateRatio(0);
            OPCPackage pkg = OPCPackage.open(new File(pathToFile));
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            return wb;
        }
        catch (Exception e) {
            log.error("Error", e);
            e.printStackTrace();
            return null;
        }
    }

    public String save(XSSFWorkbook wb, String pathToNewFile) {
        String pathToNewFileGenerated;
        if(pathToNewFile == null) {
            pathToNewFileGenerated = pathToFile.substring(0, pathToFile.length() - 5) + "_" +
                    LocalDateTime.now().format(DateTimeFormatter.ofPattern("ss_HH_mm_dd_MM")) + ".xlsx";
        }else{
            pathToNewFileGenerated = pathToNewFile;
        }
        try {

            //FileOutputStream fileOut = new FileOutputStream(pathToNewFileGenerated);


            FileOutputStream fileOut;
            if(pathToNewFile != null) {
                fileOut = new FileOutputStream(pathToNewFileGenerated + ".new");
                wb.write(fileOut);
                wb.close();
                fileOut.close();
                new File(pathToNewFileGenerated).renameTo(new File("tmpName.xlsx"));
                Files.delete(Paths.get("tmpName.xlsx"));
                Files.move(Paths.get(pathToNewFileGenerated + ".new"), Paths.get(pathToNewFileGenerated));
            }else {
                fileOut = new FileOutputStream(pathToNewFileGenerated);
                wb.write(fileOut);
                wb.close();
                fileOut.close();

            }

            //wb.write(fileOut);
        }
        catch (Exception e) {
            log.error("Error", e);
            e.printStackTrace();
        }
        return pathToNewFileGenerated;
    }
}
