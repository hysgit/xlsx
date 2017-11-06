package com.woslx.xlsx.NightOne;


/**
 * 云南36人数据
 */

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;

public class Step2 {
    public static void main(String[] args) throws Exception {
        File path = new File("/home/hy/tmp/nightone/yunnan36人数据");
        File[] files = path.listFiles(pathname -> {
            String name = pathname.getName();
            if (!pathname.isFile()) {
                return false;
            }
            if (name.startsWith(".")) {
                return false;
            }
            if (name.endsWith(".xlsx")) {
                return true;
            }
            return false;
        });

        Workbook wb = new XSSFWorkbook();
        for (File ftemp : files) {
            file2sheet(ftemp, wb);
        }

        OutputStream os = new FileOutputStream("/home/hy/tmp/nightone/yunnanout.xlsx");
        wb.write(os);
        wb.close();
        os.close();


    }

    private static void file2sheet(File ftemp, Workbook wb) throws IOException, InvalidFormatException {
        String name = ftemp.getName();

        name = name.replaceAll(" ", "").replaceAll("[0123456789]", "").replaceAll("-", "").replaceAll(".xlsx", "");
        System.out.println(name);

        Sheet sheet = wb.createSheet(name);     //新sheet
        Workbook wbx = new XSSFWorkbook(ftemp);

        Sheet sheetx = wbx.getSheetAt(0);       //被转换的文件

        int lastRowNum = sheetx.getLastRowNum();
        int x = 0;
        for (int i = 22; i <= lastRowNum; i++) {
            Row rowx = sheetx.getRow(i);
            Row row = sheet.createRow(x);
            if (i == 22) {
                Date javaDate = HSSFDateUtil.getJavaDate(rowx.getCell(3).getNumericCellValue());
                javaDate.setYear(2017);
                System.out.println(javaDate);
                CellStyle cellStyle = wb.createCellStyle();
                CreationHelper createHelper = wb.getCreationHelper();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
                Cell cell = row.createCell(2);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(javaDate);
            }
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            cell0.setCellValue(x + 1);
            cell1.setCellValue(getIntFromCell(rowx.getCell(1)));
            x++;
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
