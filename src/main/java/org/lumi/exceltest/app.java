package org.lumi.exceltest;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by John Tsantilis on 29/5/2018.
 *
 * @author John Tsantilis <i.tsantilis [at] ubitech [dot] com>
 */

public class app {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        cryptocurrencyList = ApachePOIExcelRead.readExistingWorkbook(FILE_NAME, "Sheet2", cryptocurrencyList);
        System.out.println(cryptocurrencyList.size());
        System.out.println(cryptocurrencyList.get(54779));

        ApachePOIExcelWrite.modifyExistingWorkbook(FILE_NAME, "Sheet3", cryptocurrencyList);

    }


    //==================================================================================================================
    //Class variables
    //==================================================================================================================
    private static ArrayList<Cryptocurrency> cryptocurrencyList = new ArrayList<>();

    private static final String FILE_NAME = "/home/lumi/Dropbox/unipi/paper_cryptocurrencies_forecasting/" +
            "final_cryptocurrencies_fixed_data_pop.xlsx";

}
