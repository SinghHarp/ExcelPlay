package com.harp.excelPlay;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

public class UtilsReadFromExcelV2 {
        public static File file;
        public static FileInputStream inputStream;
        public static Workbook myWorkBook;
        public static Sheet mySheet;


        public String readCellData( int rowNum , int colNum) throws IOException {

            String cellData = null;
            // Row row = mySheet.getRow(rowNum);
            String cell = mySheet.getRow(rowNum).getCell(colNum).toString();
            // cellData = row.getCell(colNum).getStringCellValue();

            return cell;
        }

        /**
         *
         * @param methodName
         * @param sheetName
         * @param filePath
         * @throws IOException
         * the following method is use to find the match of test method passed in excel sheet and then adding
         * key(Column Name)-value (Row data) pair to the hashmap.
         */

        public void createHashMap( String methodName, String sheetName , String filePath) throws IOException {

            HashMap<String,String> testData = new HashMap<String , String>();

            inputStream = new FileInputStream(new File(filePath));

            // create object of the workbook class
            myWorkBook = new XSSFWorkbook(inputStream);

            // Get number of sheets using method
            int noSheet = myWorkBook.getNumberOfSheets();
            System.out.println(" Excel Name ===> "+filePath.substring(filePath.lastIndexOf("/")+1));
            System.out.println("Number of sheets => " + noSheet);
            for (int i = 0; i < noSheet; i++) {
                System.out.println("Sheet "+(i + 1) + " ====>  " + myWorkBook.getSheetName(i));

            }

            // create object of XSSFSheet to read particular sheet

            mySheet = myWorkBook.getSheet(sheetName);
            //mysheet.getRow(i).getCell(j);
            // get the rowcount of sheet

            int rowCount = mySheet.getLastRowNum();
            System.out.println("total number of rows in " + sheetName + " ===> " + rowCount+"\n");

            for( int i =0; i <=rowCount; i++){
                Row row = mySheet.getRow(i);
               // System.out.println(readCellData( i, 1)+"\t");

                if(readCellData( i, 1).toString().equalsIgnoreCase(methodName)){

                    int colCount = row.getLastCellNum();
                    for (int j=0;j < colCount; j++){
                        String cellKeyValue = null;
                        String cellValue = null;
                        cellKeyValue = readCellData(0,j);
                        //System.out.print(cellKeyValue+"\t\t");
                        cellValue = readCellData(  i, j);
                        testData.put(cellKeyValue,cellValue);
                        //System.out.println( cellValue);

                    }

                }

            }
            System.out.println(testData);

        }

        public static void main(String[] args) throws IOException {
           UtilsReadFromExcelV2 myread = new UtilsReadFromExcelV2();
            String fileLocation = System.getProperty("user.dir")+"/src/test/resources/testFile.xlsx";
            System.out.println("=========>>>>>>>> "+fileLocation);
            myread.createHashMap( "testMethod05","Demographics", fileLocation);

        }



}
