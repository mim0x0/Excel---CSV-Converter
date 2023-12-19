package ConverterWebpage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ConvertExcelToCSV {

    public static void main(String[] args){

        //creating an inputfile object with specific file path
        File inputFile = new File("D:/yr2 sem1/cat201/ass 1/ConvertExcelToCSV/Excel-CSV-Converter-main/Files/Excel/Book1.xlsx");

        //creating an outputfile object to write excel data to csv
        File outputFile = new File("D:/yr2 sem1/cat201/ass 1/ConvertExcelToCSV/Excel-CSV-Converter-main/Files/CSV/Book1.csv");

        //for storing data into csv files
        StringBuffer data = new StringBuffer();

        try {
            URL url = outputFile.toURI().toURL();
            System.out.println("\nURL: \n" + url);
        } catch (MalformedURLException e) {
            throw new RuntimeException(e);
        }

        try{

            //creating input stream
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = null;

            //get the workbook object for excel file based on file format
            if (inputFile.getName().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            }
            else if (inputFile.getName().endsWith(".xls")){
                workbook = new HSSFWorkbook(fis);
            }
            else{
                fis.close();
                throw new Exception("File not supported!");
            }

            //get first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            //iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()){
                Row row = rowIterator.next();
                //for each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()){
                        case BOOLEAN:
                            data.append(cell.getBooleanCellValue() + ",");
                            break;

                        case NUMERIC:
                            data.append(cell.getNumericCellValue() + ",");
                            break;

                        case STRING:
                            data.append(cell.getStringCellValue() + ",");
                            break;

                        case BLANK:
                            data.append("" + ",");
                            break;

                        default:
                            data.append(cell + ",");
                    }
                }
                //appending new line after each row
                data.append('\n');
            }

            FileOutputStream fos = new FileOutputStream(outputFile);
            fos.write(data.toString().getBytes());
            fos.close();

        }

        catch (Exception e){
            e.printStackTrace();
        }
        System.out.println("Conversion from Excel file to CSV file is done!");


    }



}
