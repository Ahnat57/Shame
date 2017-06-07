
package csv_compiler;

import java.io.FileInputStream;
import java.io.*;
import java.util.*;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;


public class CSV_Compiler {

    private static String file;
    
    public static String NEW_EXCEL_FILE_LOCATION = "C:\\Users\\Alex\\Desktop\\Whaaaaa\\EBAY\\UpdatedItems.xls";
    public static String EXCEL_FILE_LOCATION = "";
    
    public static void main(String[] args) throws WriteException, IOException {
        if (args.length > 0) {
            EXCEL_FILE_LOCATION = args[0];
        }
        
       
        
    //  1. Create an Excel file            
        Workbook workbook = null;

        try {   
            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));

            int row = 0;
            Sheet sheet = workbook.getSheet(0);
            Cell[] cellColumn = sheet.getColumn(row);
            int NumItems = cellColumn.length;

            WritableWorkbook myFirstWbook = null;
            try {     

            myFirstWbook = Workbook.createWorkbook(new File(NEW_EXCEL_FILE_LOCATION));

            // create an Excel sheet
            WritableSheet excelSheet = myFirstWbook.createSheet("Import", 0);

            // add something into the Excel sheet
            Label label = new Label(0, 0, "Action(SiteID=US|Country=US|Currency=USD|Version=585|CC=ISO-8859-1)");
            excelSheet.addCell(label);

            label = new Label(2,0, "Quantity");
            excelSheet.addCell(label);

            //adding Revise to all A cells
            for(int i = 1, j = 1; i < NumItems; i++, j++) {
                label = new Label(0, j, "Revise");
                excelSheet.addCell(label);
            }

            //creating cells for Item#
            for(int i = 0, k = 0, j = 0; i < NumItems; i++, k++, j++) {
                Cell CELL1 = sheet.getCell(0, k);
                label = new Label(1, j, CELL1.getContents());
                excelSheet.addCell(label);
            }

            label = new Label(1, 0, "ItemID");
            excelSheet.addCell(label);

            for(int wow = 0, k = 1, j = 1; k < NumItems; wow++, k++, j++) {
                int total = 0, nr1 = 0, nr2 = 0, nr3 = 0;

                Cell CELL1 = sheet.getCell(2, k);
                Cell CELL2 = sheet.getCell(3, k);
                Cell CELL3 = sheet.getCell(4, k);

                nr1 = Integer.parseInt(CELL1.getContents());

                nr2 = Integer.parseInt(CELL2.getContents());;

                nr3 = Integer.parseInt(CELL3.getContents());;

                total = nr1 - nr2 - nr3;

                label = new Label(2, j, (Integer.toString(total)));
                excelSheet.addCell(label);
            }
            myFirstWbook.write();

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (myFirstWbook != null) {
                myFirstWbook.close();
            }
        }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    } 
}