package com.woslx.xlsx.taskset2;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by hy on 9/3/17.
 */
public class Task3 {
    public static void main(String[] args) throws Exception {
        String fileinut1 = "/home/hy/tmp/newdata/杭州七院20份数据睡眠分期out.xlsx";
        String fileinut2 = "/home/hy/tmp/newdata/T1out.xlsx";
        String fileinut3 = "/home/hy/tmp/newdata/T4out.xlsx";

        String filemerge = "/home/hy/tmp/newdata/merge.xlsx";

        InputStream ins = new FileInputStream(fileinut1);
        Workbook wb = new XSSFWorkbook(ins);

        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            InputStream t1 = new FileInputStream(fileinut2);
            InputStream t2 = new FileInputStream(fileinut3);

            Workbook wbt1 = new XSSFWorkbook(t1);
            Workbook wbt2 = new XSSFWorkbook(t2);

            //复制时间列到后3列
            copycolumn(sheet, 2, 5, wb);
            Sheet t = getSheetFromT1T2(wbt1, wbt2, sheet.getSheetName());

            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();

            int firstRowNumt = t.getFirstRowNum();
            int lastRowNumt = t.getLastRowNum();

            for (int i = 2; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                Cell cell1 = row.getCell(1);
                Cell cell2 = row.getCell(2);
                if (cell2 == null)
                    cell2 = row.createCell(2);

                Cell cell3 = row.getCell(3);
                if (cell3 == null)
                    cell3 = row.createCell(3);
                Row row1 = t.getRow(i + 20);
                Cell cellt1 = row1.getCell(1);
                Cell cellt2 = row1.getCell(2);
//                copyCellStyle(cell1, cell2, wb);
//                copyCellStyle(cell1, cell3, wb);
                cell2.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell2.setCellStyle(null);
                cell3.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell2.setCellValue(cellt1.getNumericCellValue());
                cell3.setCellValue(cellt2.getNumericCellValue());
            }
        }

        OutputStream ops = new FileOutputStream(filemerge);
        wb.write(ops);
        ops.flush();
        ops.close();

    }

    private static void copycolumn(Sheet sheet, int i, int i1, Workbook wb) {
        Row row = sheet.getRow(2);
        Cell cell1 = row.createCell(i1);
        Cell cell = row.getCell(i);
        if (cell != null) {
            copyValue(cell, cell1);
            copyCellStyle(cell, cell1, wb);
        }
    }

    private static void copyCellStyle(Cell cell, Cell cellout, Workbook wbout) {
        CellStyle cellStyleout = wbout.createCellStyle();
        cellStyleout.cloneStyleFrom(cell.getCellStyle());
        cellout.setCellStyle(cellStyleout);
    }

    private static Sheet getSheetFromT1T2(Workbook wbt1, Workbook wbt2, String sheetName) {
        int numberOfSheetst1 = wbt1.getNumberOfSheets();
        int numberOfSheetst2 = wbt2.getNumberOfSheets();
        for (int i = 0; i < numberOfSheetst1; i++) {
            Sheet sheetAt = wbt1.getSheetAt(i);
            String sheetName1 = sheetAt.getSheetName();
            if (sheetName1.contains(sheetName)) {
                return sheetAt;
            }
        }
        for (int i = 0; i < numberOfSheetst2; i++) {
            Sheet sheetAt = wbt2.getSheetAt(i);
            String sheetName1 = sheetAt.getSheetName();
            if (sheetName1.contains(sheetName)) {
                return sheetAt;
            }
        }
        throw new RuntimeException("未找到sheet相关的其他sheet:" + sheetName);
    }

    private static void copyValue(Cell formCell, Cell toCell) {
        switch (formCell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                toCell.setCellValue(formCell.getRichStringCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(formCell)) {
                    toCell.setCellValue(formCell.getDateCellValue());
                } else {
                    toCell.setCellValue(formCell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                toCell.setCellValue(formCell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                toCell.setCellValue(formCell.getCellFormula());
                break;
            default:

        }
    }
}
