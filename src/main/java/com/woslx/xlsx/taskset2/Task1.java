package com.woslx.xlsx.taskset2;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by hy on 8/21/17.
 */
public class Task1 {
    public static void main(String[] args) throws IOException {
        //加载txt文件
        String input = "/home/hy/tmp/newdata/杭州七院（20份数据睡眠分期）.xlsx";
        String output = "/home/hy/tmp/newdata/杭州七院20份数据睡眠分期out.xlsx";

        try {
            FileInputStream fis = new FileInputStream(input);
            Workbook wb = new XSSFWorkbook(fis);

            Iterator<Sheet> sheetIterator = wb.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();
                for (int i = firstRowNum; i <= lastRowNum; i++) {
                    Row row = sheet.getRow(i);

                    Cell cell = row.getCell(1);
                    if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        String stringCellValue = cell.getStringCellValue().trim();
                        if (("清醒").equals(stringCellValue)) {
                            cell.setCellValue(5);
                        } else if (("N1").equals(stringCellValue)) {
                            cell.setCellValue(1);
                        } else if (("N2").equals(stringCellValue)) {
                            cell.setCellValue(2);
                        } else if (("N3").equals(stringCellValue)) {
                            cell.setCellValue(3);
                        } else if (("REM").equals(stringCellValue)) {
                            cell.setCellValue(4);
                        }
                    }
                }
            }
            fis.close();
            FileOutputStream fos = new FileOutputStream(output);
            wb.write(fos);
            fos.flush();
            fos.close();
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
