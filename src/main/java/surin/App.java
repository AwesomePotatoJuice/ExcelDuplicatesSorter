package surin;


import org.apache.log4j.Logger;

import java.io.File;
import java.util.*;

/**
 * Hello world!
 *
 */
public class App 
{
    private static final Logger log = Logger.getLogger(App.class);
//    private static final String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
//    private static final String propsPath = rootPath + "/properties.properties";

    public static void main( String[] args )
    {
        Scanner sc = new Scanner(System.in);
        String path = null;
        int totalRowsRead = 0;

        System.out.println("Path to file: ");
        String pathToFile = sc.nextLine();
        log.info("Filepath before convertion: " + pathToFile);
        pathToFile = checkRemoveBrackets(pathToFile);
        log.info("Filepath after convertion: " + pathToFile);
        boolean validExtension = false;
        if(pathToFile.length() > 4 && pathToFile.substring(pathToFile.length()-5).equals(".xlsx")) {
            validExtension = true;
        }
        while(!new File(pathToFile).isFile() || !validExtension){
            System.out.println("No such file or invalid path! Repeat: ");
            pathToFile = sc.nextLine();
            pathToFile = checkRemoveBrackets(pathToFile);
            if(pathToFile.length() > 4 && pathToFile.substring(pathToFile.length()-5).equals(".xlsx")) {
                validExtension = true;
            }
        }
        System.out.println("Input start row: ");
        String rowToStart = sc.nextLine();
        while(!isNumeric(rowToStart)){
            System.out.println("Input digits only: ");
            rowToStart = sc.nextLine();
        }

        System.out.println("How much structures today? ");
        String chunkSize = sc.nextLine();
        while(!isNumeric(chunkSize)){
            System.out.println("Input digits only: ");
            rowToStart = sc.nextLine();
        }

        log.info("Reading file...");
        System.out.println("Reading file...");
        DataWorker dataWorker = new DataWorker(pathToFile, false);
        log.info("Filtering file...");
        System.out.println("Filtering data...");
        List<Integer> integers = dataWorker.filterData(Integer.parseInt(chunkSize), Integer.parseInt(rowToStart));
        int currentRow = 0;
        if (integers != null) {
            int readChunks = integers.get(0);
            currentRow = integers.get(1);
            while (readChunks < Integer.parseInt(chunkSize)) {
                System.out.println("Saving changes every 500 rows...");
                path = dataWorker.saveChanges(path);
                System.out.println("Continue...");
                dataWorker = new DataWorker(path, true);
                integers = dataWorker.filterData(Integer.parseInt(chunkSize) - readChunks, currentRow);
                if (integers == null)
                    break;
                readChunks += integers.get(0);
                currentRow = integers.get(1);
            }
            totalRowsRead = currentRow;
        }

        System.out.println("Saving changes...");
        dataWorker.saveChanges(path);
        System.out.println("Successfully ended. Last row read: " + totalRowsRead + "\n");
        sc.nextLine();

    }

    private static String checkRemoveBrackets(String pathToFile) {
        if (pathToFile.substring(0, 1).equals("\"") && pathToFile.substring(pathToFile.length() - 1).equals("\""))
            return pathToFile.substring(1, pathToFile.length() - 1);
        return pathToFile;
    }

    private static boolean isNumeric(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch(NumberFormatException e){
            log.error("null", e);
            return false;
        }
    }
}
