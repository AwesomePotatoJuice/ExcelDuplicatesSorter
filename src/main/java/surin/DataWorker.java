package surin;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class DataWorker {
    private ExcelWorker excelWorker;
    private XSSFWorkbook wb;
    private XSSFSheet sheetInitial;
    private XSSFSheet sheet1;
    private XSSFSheet sheet2;
    private XSSFSheet sheet3;
    private boolean endOfFile = false;

    public DataWorker(String pathToFile, boolean reopened){
        this.excelWorker = new ExcelWorker(pathToFile);
        this.wb = excelWorker.readWorkbook();
        sheetInitial = wb.getSheetAt(0);
        if(wb.getNumberOfSheets() < 4) {
            wb.createSheet();
            wb.createSheet();
            wb.createSheet();
        }
        sheet1 = wb.getSheetAt(1);
        sheet2 = wb.getSheetAt(2);
        sheet3 = wb.getSheetAt(3);
    }

    public List<Integer> filterData(int chunkSize, int rowToStart) {
        List<XSSFRow> rowsOnB;
        List<List<XSSFRow>> rowsOnL;
        int currentRow = rowToStart;
        int sessionRows = 0;
        for (int i = 0; i < chunkSize; i++) {
            if(endOfFile)
                return null;
            if (sessionRows >= 500 && sessionRows%500 < 5 ) {
                ArrayList<Integer> tempVals = new ArrayList<>();
                tempVals.add(i);
                tempVals.add(currentRow);
                return tempVals;
            }
            rowsOnB = getBs(currentRow); //Читаем очередные дубликаты по B
            int sizeOfB = rowsOnB.size();
            currentRow+= sizeOfB;
            sessionRows+= sizeOfB;

            if (sizeOfB == 1) {
                sendToFile(rowsOnB, false, 1);
                continue; //Оставляем в 1 если по B не повторяется
            }

            rowsOnL = getEqualsOnCell(11, rowsOnB);
            int sizeOfL = rowsOnL.get(0).size();
            if (sizeOfL < sizeOfB) { //Если есть хотя бы одна уникальная по L, то убираем полные дубликаты и переносим в 3
                List<XSSFRow> rowsOnLReady = removeDoubleLG(rowsOnL);
                sendToFile(rowsOnLReady, false, 3);
            } else if (sizeOfL == sizeOfB) { //Если все L одинаковые, то убираем полные дубликаты, кроме одной. Одну запись оставляем в 1, остальные все в 2
                List<XSSFRow> rowsOnLReady = removeDoubleLG(rowsOnL);
                sendToFile(rowsOnLReady, true, 2);
            }
        }
        return null;
    }

    private List<XSSFRow> removeDoubleLG(List<List<XSSFRow>> listsOfRows) {
        List<XSSFRow> rows = new ArrayList<>();
        for (List<XSSFRow> listOfRows: listsOfRows) {
            boolean skip = true;
            XSSFRow row1 = listOfRows.get(0);
            List<XSSFRow> markedToRemoveRows = new ArrayList<>();
            for (XSSFRow row:listOfRows) {
                if (skip) {
                    skip = false;
                    continue;
                }
                if (row1.getCell(6).getCellType() == CellType.NUMERIC) {
                    if (row1.getCell(6).getNumericCellValue() == row.getCell(6).getNumericCellValue()){
                        markedToRemoveRows.add(row);
                    }
                }else if(row1.getCell(6).getCellType() == CellType.STRING){
                    if (row1.getCell(6).getStringCellValue().equals(row.getCell(6).getStringCellValue())){
                        markedToRemoveRows.add(row);
                    }
                }
            }
            for (XSSFRow row:markedToRemoveRows) {
                listOfRows.remove(row);
            }
            rows.addAll(listOfRows);
        }
        return rows;
    }

    private List<XSSFRow> getBs(int rowToRead) {
        boolean condition = true;
        List<XSSFRow> rows = new ArrayList<>();
        XSSFRow row1 = sheetInitial.getRow(rowToRead);
        rows.add(row1);
        XSSFRow row2;
        for(int i = rowToRead; condition; i++){
            row2 = sheetInitial.getRow(i + 1);
            if (row2 == null) {
                endOfFile = true;
                return rows;
            }
            if (row1.getCell(1).getCellType() == CellType.NUMERIC) {
                if(row1.getCell(1).getNumericCellValue() == row2.getCell(1).getNumericCellValue()) {
                    rows.add(row2);
                    row1 = row2;
                }else{
                    condition = false;
                }
            }else if (row1.getCell(1).getCellType() == CellType.STRING){
                if(row1.getCell(1).getStringCellValue().equals(row2.getCell(1).getStringCellValue())) {
                    rows.add(row2);
                    row1 = row2;
                }else{
                    condition = false;
                }
            }

        }
        return rows;
    }

    private List<List<XSSFRow>> getEqualsOnCell(int cellIndex, List<XSSFRow> rowsOnB){
        List<List<XSSFRow>> listsOfRows = new ArrayList<>();
        XSSFRow row1 = rowsOnB.get(0);
        List<XSSFRow> list1 = new ArrayList<>();
        list1.add(row1);
        listsOfRows.add(list1);
        boolean skip = true;
        for (XSSFRow row:rowsOnB) {
            if (skip) {
                skip = false;
                continue;
            }
            if (row1.getCell(cellIndex).getCellType() == CellType.NUMERIC) {
                if (row1.getCell(cellIndex).getNumericCellValue() == row.getCell(cellIndex).getNumericCellValue()){
                    list1.add(row);
                    row1 = row;
                }else{
                    list1 = new ArrayList<>();
                    list1.add(row);
                    listsOfRows.add(list1);
                    row1 = row;
                }
            }else if(row1.getCell(cellIndex).getCellType() == CellType.STRING){
                if (row1.getCell(cellIndex).getStringCellValue().equals(row.getCell(cellIndex).getStringCellValue())){
                    list1.add(row);
                    row1 = row;
                }else{
                    list1 = new ArrayList<>();
                    list1.add(row);
                    listsOfRows.add(list1);
                    row1 = row;
                }
            }
        }
        return listsOfRows;
    }

    public String saveChanges(String path) {
        return excelWorker.save(wb, path);
    }

    private void sendToFile(List<XSSFRow> rowsOnLReady, boolean skipFirst, int dedicatedFile) {
        if(dedicatedFile == 3){
            int lastRowToAppend = sheet3.getLastRowNum() + 1;
            int rowsCount = rowsOnLReady.size();
            for(int i = 0; i < rowsCount; i++){
                copyRow(sheet3, rowsOnLReady.get(i), lastRowToAppend);
                lastRowToAppend++;
            }
        }
        if(dedicatedFile == 2){
            int lastRowToAppend2 = sheet2.getLastRowNum() + 1;
            int lastRowToAppend1 = sheet1.getLastRowNum() + 1;
            int rowsCount = rowsOnLReady.size();
            for(int i = 0; i < rowsCount; i++){
                if(skipFirst){
                    skipFirst = false;
                    copyRow(sheet1, rowsOnLReady.get(i), lastRowToAppend1);
                    continue;
                }
                copyRow(sheet2, rowsOnLReady.get(i), lastRowToAppend2);
                lastRowToAppend2++;
            }
        }
        if(dedicatedFile == 1){
            int lastRowToAppend1 = sheet1.getLastRowNum() + 1;
            copyRow(sheet1, rowsOnLReady.get(0), lastRowToAppend1);
        }
    }
    private void copyRow(XSSFSheet worksheet, XSSFRow sourceRowNum, int destinationRowNum) {
        // Get the source / new row

        XSSFRow  newRow = worksheet.createRow(destinationRowNum);

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRowNum.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRowNum.getCell(i);
            XSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
//            XSSFCellStyle newCellStyle = wb.createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//            ;
//            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }
    }
}
