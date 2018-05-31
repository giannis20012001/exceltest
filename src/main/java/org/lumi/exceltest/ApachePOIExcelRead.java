package org.lumi.exceltest;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

/**
 * Created by John Tsantilis on 29/5/2018.
 *
 * @author John Tsantilis <i.tsantilis [at] ubitech [dot] com>
 */

public class ApachePOIExcelRead {
    public static ArrayList<Cryptocurrency> readExistingWorkbook(String fileName, String sheetName,
                                                          ArrayList<Cryptocurrency> cryptocurrencyList)
            throws IOException {
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
                //Create the data
                String name = (String) getCellValue(
                        CellUtil.getCell(
                                CellUtil.getRow(0, sheet),
                                columnIndex));
                Double price = (Double) getCellValue(cell);
                Date date = (Date) getCellValue(
                        CellUtil.getCell(
                                CellUtil.getRow(rowIndex, sheet),
                                0));
                //Add cryptocurrency data to List
                cryptocurrencyList.add(new Cryptocurrency(name, price, date));

            }

        }

        return cryptocurrencyList;

    }

    public static Object getCellValue(Cell cell) {
        Object value;

        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();

                }
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case BLANK:
                value = 0.0;
                break;
            default:
                value = null;

        }

        return value;

    }

}