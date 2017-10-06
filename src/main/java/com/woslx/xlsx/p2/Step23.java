package com.woslx.xlsx.p2;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Iterator;

/**
 * 上海数据转换
 */
public class Step23 {
    public static void main(String[] args) throws Exception {
        File path = new File("/home/hy/tmp/p2/yunnan/src");
        File[] files = path.listFiles();
        for (File file : files) {
            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 22; i <= lastRowNum; i++) {
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(1);
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
                    } else if (("清醒").equals(stringCellValue)) {
                        cell.setCellValue(5);
                    } else {
                        System.out.println("error");
                        System.exit(1);
                    }
                }
            }
            OutputStream ops = new FileOutputStream("/home/hy/tmp/p2/yunnan/res/" + file.getName().substring(0, file.getName().length() - 5) + ".xlsx");
            wb.write(ops);
        }


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
                } else if (("清醒").equals(stringCellValue)) {
                    cell.setCellValue(5);
                } else {
                    System.out.println("error");
                    System.exit(1);
                }
            }
        }
    }
}
