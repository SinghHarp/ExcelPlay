package com.harp.excelPlay;

import jdk.nashorn.internal.ir.CatchNode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class WriteToExcel {

    //String filePath = "src/test/resources";
    public void writeData(String filePath, String sheetName) throws IOException {
        // Create object of the File class to open xlsx and create object of FileInputStream to read xlsx.
        // File file = new File(filePath+"/"+fileName);
        // FileInputStream inputStream = new FileInputStream(file);

        FileInputStream inputStream = new FileInputStream(new File(filePath));

        // create object of the workbook class
        Workbook myworkbook = new XSSFWorkbook(inputStream);

        Sheet mysheet = myworkbook.getSheet(sheetName);
        // get row count of the sheet
        int rowCount = mysheet.getLastRowNum() - mysheet.getFirstRowNum();
        System.out.println("Total Number of rows ==> " + rowCount);

        // Creating an object array of data that needs to be added
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[]{"10", "fnam1", "Lnam1", "male", "USA", "34", "13/10/2001", "1454"});
        data.put("2", new Object[]{"11", "fnam2", "Lnam2", "male", "USA", "32", "13/10/2001", "1457"});
        data.put("3", new Object[]{"12", "fnam3", "Lnam3", "male", "USA", "37", "13/10/2001", "1459"});
        // Iterate over the data
        Set<String> keySet = data.keySet();
        for (String key : keySet) {
            Row rowNew1 = mysheet.createRow(rowCount++);
            Object[] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj : objArr) {
                Cell cell = rowNew1.createCell(cellNum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }

        }
        // Create a new row at end of the rows and add data

        // Close the input stream
        inputStream.close();
        // create object of FileOutputStream to write to excel file
        try {
            FileOutputStream outputStream = new FileOutputStream(filePath);
            myworkbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        WriteToExcel mywrite = new WriteToExcel();
        String fileLocation = System.getProperty("user.dir") + "/src/test/resources/testFile.xlsx";
        mywrite.writeData(fileLocation, "Demographics");
    }

}