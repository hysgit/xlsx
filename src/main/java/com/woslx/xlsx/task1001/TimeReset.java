package com.woslx.xlsx.task1001;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

/**
 * Created by hy on 10/1/17.
 */
public class TimeReset {
    public static void main(String[] args) throws Exception {
        String fileinput = "/home/hy/tmp/sing/2017-07-23周密.xlsx";
        String fileoutput = "/home/hy/tmp/sing/2017-07-23周密out.xlsx";

        FileInputStream fis = new FileInputStream(fileinput);
        Workbook wb = new XSSFWorkbook(fis);

        Sheet sheet0 = wb.getSheet("Sheet0");
        int lastRowNum = sheet0.getLastRowNum();
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        Calendar calendar = Calendar.getInstance();
        calendar.set(2017, Calendar.MONTH, 23, 21, 18, 0);
        for (int i = 1; i <= lastRowNum; i++) {
            Cell cell = sheet0.getRow(i).getCell(3);
            cell.setCellValue(sdf.format(calendar.getTime()));
            calendar.add(Calendar.SECOND, 30);
        }

        fis.close();
        FileOutputStream fos = new FileOutputStream(fileoutput);
        wb.write(fos);
        fos.close();
    }
}
