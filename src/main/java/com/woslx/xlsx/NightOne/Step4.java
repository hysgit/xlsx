package com.woslx.xlsx.NightOne;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

public class Step4 {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        File ding = new File("/home/hy/tmp/nightone/dingout.xlsx");
        File yunnan = new File("/home/hy/tmp/nightone/yunnanout.xlsx");
        File shanghai = new File("/home/hy/tmp/nightone/shanghaiout.xlsx");

        Workbook wbding = new XSSFWorkbook(ding);

    }
}
