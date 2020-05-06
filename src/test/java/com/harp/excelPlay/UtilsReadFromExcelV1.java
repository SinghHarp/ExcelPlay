package com.harp.excelPlay;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * 1)Lesson to learn from this program that in case of fileinput stream I cannot execute both the methods at same time because fileinput stream then would have conflicts
 *  I was eariler trying to pass file, inputtream , myWorkBook ... object reference in readCellData method but ideally it should be in HashMap create method because in that
 *   method all the required object references( inputStream, myWorkBook....) will be initialized and values would be passed and then we make call to readCellData method
 *   which just required mySheet and that already had been initialized in HashMap create Method.
 *
 * 2)  We need to be cautious of the order we are calling the method because that in itself decides which variables would of the method called would be initialized first
 * subsequently then can be used in upcoming methods. as in example below we made call to HashMapCreate method in which first all the required objects are initialized and
 * then in the same method there is call to readCellData method that needs mysheet but it has been already been provided value in HashMapCreate method.
 */

public class UtilsReadFromExcelV1 {
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
     * the following method is use to find the match of test method passed in excel sheet and printing values
     * of the particular row
     */

    public void findTestMethodMatch( String methodName, String sheetName , String filePath) throws IOException {

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
        System.out.println("total number of rows in " + sheetName + " ===> " + mySheet.getLastRowNum()+"\n");

        int rowCount = mySheet.getLastRowNum();
        for( int i =0; i <=rowCount; i++){
            Row row = mySheet.getRow(i);
            if(readCellData( i, 1).toString().equalsIgnoreCase(methodName)){

                int colCount = row.getLastCellNum();
                for (int j=0;j < colCount; j++){
                    String cellValue = null;
                    cellValue = readCellData(  i, j);
                    System.out.print( cellValue+"\t");

                }

            }

        }
        System.out.println(" total number of rows ==>"+ rowCount);

    }

    public static void main(String[] args) throws IOException {
        UtilsReadFromExcelV1 myread = new UtilsReadFromExcelV1();
        String fileLocation = System.getProperty("user.dir")+"/src/test/resources/testFile.xlsx";
        System.out.println("=========>>>>>>>> "+fileLocation);
        myread.findTestMethodMatch( "testMethod01","Demographics", fileLocation);

    }


}


