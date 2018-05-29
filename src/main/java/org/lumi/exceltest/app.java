package org.lumi.exceltest;

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

    }

    private static final String FILE_NAME = "/home/lumi/Dropbox/unipi/Cryptocurrencies_forecasting/" +
            "final_cryptocurrencies_fixed_data_pop.xlsx";

}
