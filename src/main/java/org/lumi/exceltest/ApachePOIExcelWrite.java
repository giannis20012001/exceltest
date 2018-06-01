package org.lumi.exceltest;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

/**
 * Created by John Tsantilis on 29/5/2018.
 *
 * @author John Tsantilis <i.tsantilis [at] ubitech [dot] com>
 */

public class ApachePOIExcelWrite {
    @SuppressWarnings("Duplicates")
    public static void modifyExistingWorkbook(String fileName, String sheetName,
                                              ArrayList<Cryptocurrency> cryptocurrencyList)
            throws IOException, InvalidFormatException {
        //Get a Workbook from an Excel file (.xls or .xlsx)
        FileInputStream excelFile = new FileInputStream(new File(fileName));
        Workbook workbook = new XSSFWorkbook(excelFile);
        //Getting the Sheet with the specified name
        Sheet sheet = workbook.getSheet(sheetName);

        //Max 30 (zero-based)
        for (int columnIndex = 1; columnIndex < 31; columnIndex++) {
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                Row row = CellUtil.getRow(rowIndex, sheet);
                Cell cell = CellUtil.getCell(row, columnIndex);

                //Create the cell if it doesn't exist
                if (cell == null) {
                    cell = row.createCell(columnIndex);

                }

                String name = (String) ApachePOIExcelRead.getCellValue(
                        CellUtil.getCell(
                                CellUtil.getRow(0, sheet),
                                columnIndex));
                Date date = (Date) ApachePOIExcelRead.getCellValue(
                        CellUtil.getCell(
                                CellUtil.getRow(rowIndex, sheet),
                                0));

                for (Cryptocurrency cryptocurrency : cryptocurrencyList) {
                    if ((name.contentEquals(cryptocurrency.getName())) &&
                            (date.compareTo(cryptocurrency.getDate()) == 0)) {
                        //Enter the cell's value
                        cell.setCellValue(cryptocurrency.getPrice());

                    }

                }

            }

        }

        // Write the output to the file
        FileOutputStream fileOut = new FileOutputStream(fileName);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

    }

}