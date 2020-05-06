package com.harp.excelPlay;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadFromExcel {
    //String filePath = "src/test/resources";
    public void readData(String filePath, String sheetName) throws IOException {
        // Create object of the File class to open xlsx and create object of FileInputStream to read xlsx.
        // File file = new File(filePath+"/"+fileName);
        // FileInputStream inputStream = new FileInputStream(file);

        FileInputStream inputStream = new FileInputStream(new File(filePath));

        // create object of the workbook class
        Workbook myworkbook = new XSSFWorkbook(inputStream);
        // Get number of sheets using method
        int noSheet = myworkbook.getNumberOfSheets();
        System.out.println("Number of sheets => " + noSheet);
        for (int i = 0; i < noSheet; i++) {
            System.out.println("Sheet "+i + 1 + " ====>  " + myworkbook.getSheetName(i));
        }

        // create object of XSSFSheet to read particular sheet

        Sheet mysheet = myworkbook.getSheet(sheetName);
        // get the rowcount of sheet
        System.out.println("total number of rows in " + sheetName + " ===> " + mysheet.getLastRowNum());

        // to get column count of a particular row
        for (int i = 0; i <= mysheet.getLastRowNum(); i++) {
            Row row = mysheet.getRow(i);
            // System.out.println("Number of columns ==>" + row.getLastCellNum());
            for (int j = 0; j < row.getLastCellNum(); j++) {
                System.out.print(row.getCell(j).getStringCellValue() + "\t\t\t");

            }
            System.out.println();
        }
    }

    public static void main(String[] args) throws IOException {
        ReadFromExcel myread = new ReadFromExcel();
        String fileLocation = System.getProperty("user.dir")+"/src/test/resources/testFile.xlsx";
        myread.readData(fileLocation, "Demographics");
    }


}