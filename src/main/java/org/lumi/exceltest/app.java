package org.lumi.exceltest;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by John Tsantilis on 29/5/2018.
 *
 * @author John Tsantilis <i.tsantilis [at] ubitech [dot] com>
 */

public class app {
    public static void main(String[] args) throws IOException {
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());

        }

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

    }

    private static final String FILE_NAME = "/home/lumi/Dropbox/unipi/paper_cryptocurrencies_forecasting/" +
            "final_cryptocurrencies_fixed_data_pop.xlsx";

}
