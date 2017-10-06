package com.woslx.xlsx.p2;

import com.sun.jndi.toolkit.ctx.StringHeadTail;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.java2d.opengl.GLXSurfaceData;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class Step24 {
    public static void main(String[] args) throws Exception {
        //以丁为不变

        File shanghai1 = new File("/home/hy/tmp/p2/shanghai/shanghait1out.xlsx");
        Workbook wbshanghait1 = new XSSFWorkbook(shanghai1);
        File shanghai4 = new File("/home/hy/tmp/p2/shanghai/shanghait4out.xlsx");
        Workbook wbshanghait4 = new XSSFWorkbook(shanghai4);

        File[] files = new File("/home/hy/tmp/p2/yunnan/res/").listFiles();

        Workbook wb1 = new XSSFWorkbook("/home/hy/tmp/p2/ding/t1out.xlsx");
        Iterator<Sheet> iterator = wb1.sheetIterator();
        while (iterator.hasNext()) {
            Sheet sheet = iterator.next();
            Sheet shanghaiSheet = findSheetFromShangHai(sheet.getSheetName(), wbshanghait1, wbshanghait4);
            File file = findFileFromYunNan(sheet.getSheetName(), files);
            Sheet sheetYun = new XSSFWorkbook(file).getSheetAt(0);
            merge(sheet, shanghaiSheet, sheetYun);
        }

        Workbook wb4 = new XSSFWorkbook("/home/hy/tmp/p2/ding/t4out.xlsx");
        iterator = wb4.sheetIterator();
        while (iterator.hasNext()) {
            Sheet sheet = iterator.next();
            Sheet shanghaiSheet = findSheetFromShangHai(sheet.getSheetName(), wbshanghait1, wbshanghait4);
            File file = findFileFromYunNan(sheet.getSheetName(), files);
            Sheet sheetYun = new XSSFWorkbook(file).getSheetAt(0);
            merge(sheet, shanghaiSheet, sheetYun);
        }


    }

    private static void merge(Sheet sheet, Sheet shanghaiSheet, Sheet sheetYun) {
        int lastRowNum = sheet.getLastRowNum();
        int x = 0;
        for (int i = 1; i <= lastRowNum; i++) {
            Row rowding = sheet.getRow(x + 1);
            Row row = shanghaiSheet.getRow(2 + x);
            if(row == null){
                System.out.println();
            }
            Cell cellshanghai = row.getCell(1);

            Cell cellYunnan = sheetYun.getRow(22 + x).getCell(1);
            Cell cell5 = rowding.getCell(5);
            double data1 = cell5.getNumericCellValue();
            Cell cell6 = rowding.getCell(6);
            if (cell6 == null) {
                cell6 = rowding.createCell(6);
            }
            int data2 = (int) cellYunnan.getNumericCellValue();
            cell6.setCellValue(data2);
            Cell cell7 = rowding.getCell(7);
            if (cell7 == null) {
                cell7 = rowding.createCell(7);
            }
            int data3 = (int) cellshanghai.getNumericCellValue();
            cell7.setCellValue(data3);

            x++;
        }

    }

    private static File findFileFromYunNan(String sheetName, File[] files) {
        for (File file : files) {
            String fileName = file.getName().replaceAll("[0-9]*","").replaceAll(" ","").replaceAll("-","").replaceAll(".xlsx","");
            if (sheetName.contains(fileName)) {
                return file;
            }
        }
        System.out.println("未找到文件!" + sheetName);
        System.exit(1);
        return null;
    }

    private static Sheet findSheetFromShangHai(String name, Workbook wbshanghait1, Workbook wbshanghait4) {
        Iterator<Sheet> sheetIterator = wbshanghait1.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
//            String sheetName = sheet.getSheetName();
            if (name.contains(sheet.getSheetName().replaceAll("[0-9]","").replaceAll(" ","").replaceAll("-","").replaceAll(".xlsx",""))) {
//            if (name.contains(sheetName)) {
                return sheet;
            }

        }

        sheetIterator = wbshanghait4.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
//            String sheetName = sheet.getSheetName();
            if (name.contains(sheet.getSheetName().replaceAll("[0-9]","").replaceAll(" ","").replaceAll("-","").replaceAll(".xlsx",""))) {
                return sheet;
            }
        }
        System.out.println("未找到文件" + name);
        System.exit(1);
        return null;
    }
}
