package com.excelmerge.Excel_Merge.Config;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

public class ReadNWriteExcel {
    static final String path1 = "D:/Scaler/Java Programs/Excel-Merge/src/Excel Merge/SANJAY_GSTR1_NOV_2024.xlsx";
    static final String path2 = "D:/Scaler/Java Programs/Excel-Merge/src/Excel Merge/BM_GSTR1_NOV_2024.xlsx";
    public void test(){
        try {
            FileInputStream file = new FileInputStream(new File(path1));
            FileInputStream file2 = new FileInputStream(new File(path2));
            Workbook workbook = new XSSFWorkbook(file);
            Workbook newWorkBook = new XSSFWorkbook();
            for(Sheet sheet : workbook){
                Sheet newSheet = newWorkBook.createSheet(sheet.getSheetName());
                int rowIdx = 0;
                for(Row row : sheet){
                    Row newRow = newSheet.createRow(rowIdx++);
                    int colIdx = 0;
                    //System.out.println("Hello");
                    for(int i=0; i<row.getLastCellNum(); i++){
                        Cell cell = row.getCell(i);
                        if(cell == null){
                            colIdx++;
                            continue;
                        }
                        newRow.createCell(colIdx++).setCellValue(getCellValue(cell));
                    }
                }
            }
            workbook.close();

            workbook =  new XSSFWorkbook(file2);
            for(Sheet sheet : workbook){
                Sheet newSheet = newWorkBook.getSheet(sheet.getSheetName());
                int rowIdx = newSheet.getPhysicalNumberOfRows();
                for(Row row : sheet){
                    //skipping header of the 2nd file
                    if(row.getRowNum() == sheet.getFirstRowNum())    continue;

                    Row newRow = newSheet.createRow(rowIdx++);
                    int colIdx = 0;
                    //System.out.println("Hello");
                    for(int i=0; i<row.getLastCellNum(); i++){
                        Cell cell = row.getCell(i);
                        if(cell == null){
                            colIdx++;
                            continue;
                        }
                        newRow.createCell(colIdx++).setCellValue(getCellValue(cell));
                    }
                }
                //if sheet is not blank call this fun to auto adjust cell width
                if(newSheet.getFirstRowNum() != -1) {
                    setAutoAdjustedWidthToAllCell(newSheet);
                }
            }
            workbook.close();


            FileOutputStream fileOutputStream = new FileOutputStream("D:/Scaler/Java Programs/Excel-Merge/src/Excel Merge/output.xlsx");
            newWorkBook.write(fileOutputStream);
            newWorkBook.close();
            workbook.close();


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e){
            throw new RuntimeException(e);
        }
    }

    private void setAutoAdjustedWidthToAllCell(Sheet sheet) {
        for(int i=0; i<sheet.getRow(0).getLastCellNum(); i++){
            sheet.autoSizeColumn(i);
        }
    }

    private String getCellValue(Cell cell) {
        String pattern = "d-MMM-yy";
        DateFormat df = new SimpleDateFormat(pattern);
        try{
            if(DateUtil.isCellDateFormatted(cell)){
                String date = df.format(cell.getDateCellValue());
                return date;
            }
            else{
                return getCellValueOtherThanDate(cell);
            }
        }
        catch (Exception e){
            return getCellValueOtherThanDate(cell);
        }
    }
    private String getCellValueOtherThanDate(Cell cell){
        CellType cellType = cell.getCellType();
        switch (cellType){
            case STRING:
                return cell.getRichStringCellValue().getString();
            case NUMERIC:
                double numericCellValue = cell.getNumericCellValue();
                if(numericCellValue%1 == 0){
                    return (int)numericCellValue+"";
                }
                else{
                    return numericCellValue+"";
                }
            default :
                return "";
        }
    }

}
