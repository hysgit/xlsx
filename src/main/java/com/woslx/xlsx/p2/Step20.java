package com.woslx.xlsx.p2;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;

/**
 * 把丁数据转换
 */
public class Step20 {
    public static void main(String[] args) throws Exception {
        FileInputStream file1 = new FileInputStream("/home/hy/tmp/p2/ding/T1/第二批数据汇总T1.xlsx");
        FileInputStream file2 = new FileInputStream("/home/hy/tmp/p2/ding/T4/第二批数据汇总T4.xlsx");

        XSSFWorkbook wb1 = new XSSFWorkbook(file1);
        XSSFWorkbook wb2 = new XSSFWorkbook(file2);

        FileOutputStream fos1 = new FileOutputStream("/home/hy/tmp/p2/ding/t1out.xlsx");
        FileOutputStream fos4 = new FileOutputStream("/home/hy/tmp/p2/ding/t4out.xlsx");

        Iterator<Sheet> iterator = wb1.sheetIterator();
        while (iterator.hasNext()) {
            Sheet sheet = iterator.next();
            formatSheet(sheet);
        }
        wb1.write(fos1);


        Iterator<Sheet> iterator2 = wb2.sheetIterator();
        while (iterator2.hasNext()) {
            Sheet sheet = iterator2.next();
            formatSheet(sheet);
        }
        wb2.write(fos4);



    }

    private static void formatSheet(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(5);
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                String stringCellValue = cell.getStringCellValue().trim();
                if (("n拺").equals(stringCellValue)) {
                    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                    cell.setCellValue(5);
                } else if (("N1").equals(stringCellValue)) {
                    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                    cell.setCellValue(1);
                } else if (("N2").equals(stringCellValue)) {
                    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                    cell.setCellValue(2);
                } else if (("N3").equals(stringCellValue)) {
                    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                    cell.setCellValue(3);
                } else if (("REM").equals(stringCellValue)) {
                    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                    cell.setCellValue(4);
                }
                else if(("清醒").equals(stringCellValue)) {
                    cell.setCellValue(5);
                }
            }
        }
    }
}
