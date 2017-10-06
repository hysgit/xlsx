package com.woslx.xlsx.sub;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.crypto.Cipher;
import java.io.*;
import java.util.Iterator;

/**
 * Created by hy on 8/21/17.
 */
public class Task2 {
    public static void main(String[] args) throws IOException {
        //加载txt文件
        File dir = new File("/home/hy/tmp/xlsx/");
        File[] files = dir.listFiles();
        for (File file : files) {
            try {
                if(file.getName().startsWith(".~"))
                    continue;
                FileInputStream fis = new FileInputStream(file);
                Workbook wb = new XSSFWorkbook(fis);

                Iterator<Sheet> sheetIterator = wb.sheetIterator();
                while(sheetIterator.hasNext()){
                    Sheet sheet = sheetIterator.next();
                    int firstRowNum = sheet.getFirstRowNum();
                    int lastRowNum = sheet.getLastRowNum();
                    for (int i = firstRowNum; i <= lastRowNum; i++) {
                        Row row = sheet.getRow(i);
                        for (int j = 4; j <= 5; j++) {
                            Cell cell = row.getCell(j);
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
                        if (i != 0) {
                            Cell cell4 = row.getCell(4);
                            Cell cell5 = row.getCell(5);
                            Cell cell6 = row.getCell(6);

                            if (cell4 == null || cell5 == null || cell6 == null) {
                                System.out.println();
                            } else {
                                if (cell4.getCellType() == Cell.CELL_TYPE_NUMERIC &&
                                        cell5.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                                    if (cell4.getNumericCellValue() == cell5.getNumericCellValue()) {
                                        double numericCellValue = cell4.getNumericCellValue();
                                        cell6.setCellValue(numericCellValue);
                                    } else {
                                        cell6.setCellValue(0);
                                    }
                                }
                            }
                        }
                    }
                }
                fis.close();
                FileOutputStream fos = new FileOutputStream(file);
                wb.write(fos);
                fos.flush();
                fos.close();
                wb.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
