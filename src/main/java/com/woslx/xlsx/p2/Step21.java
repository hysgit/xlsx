package com.woslx.xlsx.p2;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

/**
 * 上海数据转换
 */
public class Step21 {
    public static void main(String[] args) throws Exception {
        FileInputStream file1 = new FileInputStream("/home/hy/tmp/p2/shanghai/8-28-9-18T1第二批原始数据.xlsx");
        FileInputStream file2 = new FileInputStream("/home/hy/tmp/p2/shanghai/8-28-9-18原始数据T4.xlsx");

        XSSFWorkbook wb1 = new XSSFWorkbook(file1);
        XSSFWorkbook wb2 = new XSSFWorkbook(file2);

        FileOutputStream fos1 = new FileOutputStream("/home/hy/tmp/p2/shanghai/shanghait1out.xlsx");
        FileOutputStream fos4 = new FileOutputStream("/home/hy/tmp/p2/shanghai/shanghait4out.xlsx");

        Iterator<Sheet> iterator = wb1.sheetIterator();
        while (iterator.hasNext()) {
            Sheet sheet = iterator.next();
            formatSheet(sheet);
        }
        wb1.write(fos1);
        Iterator<Sheet> iterator2 = wb1.sheetIterator();
        while (iterator2.hasNext()) {
            Sheet sheet = iterator2.next();
            formatSheet(sheet);
        }
        wb2.write(fos4);



    }

    private static void formatSheet(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 2; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(1);
            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                String stringCellValue = cell.getStringCellValue().trim();
                if (("未评分").equals(stringCellValue)) {
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
                else if(("清醒").equals(stringCellValue)) {
                    cell.setCellValue(5);
                }
                else{
                    System.out.println("error");
                    System.exit(1);
                }
            }
        }
    }
}
