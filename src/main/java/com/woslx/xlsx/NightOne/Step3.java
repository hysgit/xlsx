package com.woslx.xlsx.NightOne;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

/**
 * 上海数据
 */


public class Step3 {
    public static void main(String[] args) throws Exception {
        File[] files = new File[]{new File("/home/hy/tmp/nightone/shanghai/T1第一晚数据.xlsx"),
                new File("/home/hy/tmp/nightone/shanghai/T4第一晚数据.xlsx")};
        Workbook wb = new XSSFWorkbook();
        for (File file : files) {
            mergeFile(file, wb);
        }
        OutputStream os = new FileOutputStream("/home/hy/tmp/nightone/shanghaiout.xlsx");
        wb.write(os);
        wb.close();
        os.close();
    }

    private static void mergeFile(File file, Workbook wb) throws Exception {
        Workbook wbx = new XSSFWorkbook(file);
        Iterator<Sheet> sheetIterator = wbx.sheetIterator();
        while (sheetIterator.hasNext()){
            Sheet sheetx = sheetIterator.next();
            String name = sheetx.getSheetName();
            name = name.replaceAll("\\.", "").replaceAll("[0123456789]", "").replaceAll(" ", "");
            int lastRowNum = sheetx.getLastRowNum();
            Sheet sheet = wb.createSheet(name);
            int x = 0;
            for (int i = 2; i <= lastRowNum; i++) {
                Row row = sheet.createRow(x);
                Cell cell0 = row.createCell(0);
                Cell cell1 = row.createCell(1);

                cell0.setCellValue(x + 1);
                cell1.setCellValue(getIntFromCell(sheetx.getRow(i).getCell(1)));
                x++;
            }
        }


    }

    public static int getIntFromCell(Cell cell) {
        int cellType = cell.getCellType();
        if (cellType == Cell.CELL_TYPE_STRING) {
            String stringCellValue = cell.getStringCellValue();
            if (stringCellValue.contains("n拺")) {
                return 5;
            } else if (stringCellValue.contains("清醒")) {
                return 5;
            }else if (stringCellValue.contains("未评分")) {
                return 5;
            } else if (stringCellValue.contains("N1")) {
                return 1;
            } else if (stringCellValue.contains("N2")) {
                return 2;
            } else if (stringCellValue.contains("N3")) {
                return 3;
            } else if (stringCellValue.contains("REM")) {
                return 4;
            } else {
                System.out.println("未知字符串:" + stringCellValue);
                System.exit(1);
                return 0;
            }
        } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
            return (int) cell.getNumericCellValue();
        } else {
            System.out.println("类型:" + cellType);
            System.exit(1);
            return 0;
        }
    }
}
