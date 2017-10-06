package com.woslx.xlsx.task1001;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by hy on 8/21/17.
 * 把T1和T4进行转换
 */
public class Task2 {
    public static void main(String[] args) throws IOException {
        //加载txt文件
        String input[] = {"/home/hy/tmp/sing/云南数据.xlsx"};
        String output[] = {"/home/hy/tmp/sing/云南数据out.xlsx"};

        try {
            for(int x = 0;x < input.length;x++) {
                FileInputStream fis = new FileInputStream(input[x]);
                Workbook wb = new XSSFWorkbook(fis);

                Iterator<Sheet> sheetIterator = wb.sheetIterator();
                while (sheetIterator.hasNext()) {
                    Sheet sheet = sheetIterator.next();
                    int firstRowNum = sheet.getFirstRowNum();
                    int lastRowNum = sheet.getLastRowNum();
                    for (int i = 22; i <= lastRowNum; i++) {
                        Row row = sheet.getRow(i);

                        Cell cell = row.getCell(1);
                        hcell(cell);
                        Cell cell2 = row.getCell(2);
                        hcell(cell2);
                    }
                }
                fis.close();
                FileOutputStream fos = new FileOutputStream(output[x]);
                wb.write(fos);
                fos.flush();
                fos.close();
                wb.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void hcell(Cell cell) {
        if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
            String stringCellValue = cell.getStringCellValue().trim();
            if (("n拺").equals(stringCellValue)) {
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
        }
    }
}
