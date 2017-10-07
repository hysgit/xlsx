package com.woslx.xlsx.p2;

import com.sun.jndi.toolkit.ctx.StringHeadTail;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.java2d.opengl.GLXSurfaceData;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
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
            saveSheet(sheet);
        }
        FileOutputStream fos1 = new FileOutputStream("/home/hy/tmp/p2/fos1.xlsx");
        wb1.write(fos1);
        wb1.close();
        fos1.close();


        Workbook wb4 = new XSSFWorkbook("/home/hy/tmp/p2/ding/t4out.xlsx");
        iterator = wb4.sheetIterator();
        while (iterator.hasNext()) {
            Sheet sheet = iterator.next();
            Sheet shanghaiSheet = findSheetFromShangHai(sheet.getSheetName(), wbshanghait1, wbshanghait4);
            File file = findFileFromYunNan(sheet.getSheetName(), files);
            Sheet sheetYun = new XSSFWorkbook(file).getSheetAt(0);
            merge(sheet, shanghaiSheet, sheetYun);
            saveSheet(sheet);
        }
        FileOutputStream fos4 = new FileOutputStream("/home/hy/tmp/p2/fos4.xlsx");
        wb4.write(fos4);
        wb4.close();
        fos4.close();

        //拆分


    }

    private static void saveSheet(Sheet sheetYuan) throws IOException {
        Workbook wb = new XSSFWorkbook();
        FileOutputStream fos = new FileOutputStream("/home/hy/tmp/p2/result/" + sheetYuan.getSheetName() + ".xlsx");
        Sheet wbSheet = wb.createSheet();
        int lastRowNum = sheetYuan.getLastRowNum();
        Row row = wbSheet.createRow(0);
        row.createCell(0).setCellValue("epoch_number");
        row.createCell(1).setCellValue("unix_ts");
        row.createCell(2).setCellValue("date");
        row.createCell(3).setCellValue("start_time");
        row.createCell(4).setCellValue("final_score");
        row.createCell(5).setCellValue("score1(丁)");
        row.createCell(6).setCellValue("score2");
        row.createCell(7).setCellValue("score3");
        row.createCell(8).setCellValue("remark");
        for (int i = 1; i <= lastRowNum; i++) {
            Row rowYuan = sheetYuan.getRow(i);
            Row newsheetRow = wbSheet.createRow(i);
            Cell cell0 = newsheetRow.createCell(0);
            cell0.setCellValue(i);
            if (i == 1) {
                Cell cell2yuan = rowYuan.getCell(2);
                Date javaDate = HSSFDateUtil.getJavaDate(cell2yuan.getNumericCellValue());
                Cell newcell2 = newsheetRow.createCell(2);
                CellStyle cellStyle2 = wb.createCellStyle();
                CreationHelper createHelper2 = wb.getCreationHelper();
                cellStyle2.setDataFormat(createHelper2.createDataFormat().getFormat("m/d/yyyy"));
                newcell2.setCellStyle(cellStyle2);
                newcell2.setCellValue(javaDate);
            }
            Cell cell3yuan = rowYuan.getCell(3);
            Cell newcell3 = newsheetRow.createCell(3);
            Date javaDate = HSSFDateUtil.getJavaDate(cell3yuan.getNumericCellValue());
            CellStyle cellStyle = wb.createCellStyle();
            CreationHelper createHelper = wb.getCreationHelper();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
            newcell3.setCellStyle(cellStyle);
            newcell3.setCellValue(javaDate);

            Cell cell4 = newsheetRow.createCell(4);
            cell4.setCellValue(rowYuan.getCell(4).getNumericCellValue());

            Cell cell5 = newsheetRow.createCell(5);
            cell5.setCellValue(rowYuan.getCell(5).getNumericCellValue());

            Cell cell6 = newsheetRow.createCell(6);
            cell6.setCellValue(rowYuan.getCell(6).getNumericCellValue());

            Cell cell7 = newsheetRow.createCell(7);
            cell7.setCellValue(rowYuan.getCell(7).getNumericCellValue());
        }
        wb.write(fos);
        wb.close();
        fos.close();
    }


    private static void merge(Sheet sheet, Sheet shanghaiSheet, Sheet sheetYun) {
        int lastRowNum = sheet.getLastRowNum();
        int x = 0;
        int allSame = 0;//所有都相等
        int allNotSame = 0;//所有都不相等
        int not12 = 0;  //12不想等
        int not13 = 0;  //13不想等
        int not23 = 0;  //23不想等

        for (int i = 1; i <= lastRowNum; i++) {
            Row rowding = sheet.getRow(x + 1);
            Row row = shanghaiSheet.getRow(2 + x);
            if (row == null) {
                System.out.println();
            }
            Cell cellshanghai = row.getCell(1);

            Cell cellYunnan = sheetYun.getRow(22 + x).getCell(1);
            Cell cell5 = rowding.getCell(5);
            if (cell5.getCellType() == Cell.CELL_TYPE_STRING) {
                System.out.println();
            }
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

            int data3 = 0;
            int cellType = cellshanghai.getCellType();
            if (cellType == Cell.CELL_TYPE_STRING) {
                data3 = Integer.valueOf(cellshanghai.getStringCellValue());
            } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
                data3 = (int) cellshanghai.getNumericCellValue();
            } else {
                System.exit(1);
                System.out.println();
            }
            cell7.setCellValue(data3);
            Cell cell4 = rowding.createCell(4);
            if ((data1 == data2) && (data1 == data3)) {
                allSame++;
                cell4.setCellValue(data1);
            } else {
                if (data1 == data2) {
                    cell4.setCellValue(data1);
                } else if (data1 == data3) {
                    cell4.setCellValue(data1);
                } else if (data2 == data3) {
                    cell4.setCellValue(data2);
                } else {
                    allNotSame++;
                    cell4.setCellValue(0);
                }
            }

            if (data1 != data2) {
                not12++;
            }
            if (data1 != data3) {
                not13++;
            }
            if (data2 != data3) {
                not23++;
            }

            x++;
        }

        System.out.println("name:" + sheet.getSheetName() + " 总数量:" + x);
        System.out.println("全部一样:" + allSame);
        System.out.println("全部不一样:" + allNotSame + " 差异率:" + getpersent(allNotSame, x));
        System.out.println("12不一样:" + not12 + " 差异率:" + getpersent(not12, x));
        System.out.println("13不一样:" + not13 + " 差异率:" + getpersent(not13, x));
        System.out.println("23不一样:" + not23 + " 差异率:" + getpersent(not23, x));
        System.out.println();

    }

    private static String getpersent(int allNotSame, int j) {
        String s = allNotSame * 1.0 / j * 100 + "";
        return s.substring(0, s.indexOf(".") + 2) + "%";
    }

    private static File findFileFromYunNan(String sheetName, File[] files) {
        for (File file : files) {
            String fileName = file.getName().replaceAll("[0-9]*", "").replaceAll(" ", "").replaceAll("-", "").replaceAll(".xlsx", "");
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
            if (name.contains(sheet.getSheetName().replaceAll("[0-9]", "").replaceAll(" ", "").replaceAll("-", "").replaceAll(".xlsx", ""))) {
//            if (name.contains(sheetName)) {
                return sheet;
            }

        }

        sheetIterator = wbshanghait4.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
//            String sheetName = sheet.getSheetName();
            if (name.contains(sheet.getSheetName().replaceAll("[0-9]", "").replaceAll(" ", "").replaceAll("-", "").replaceAll(".xlsx", ""))) {
                return sheet;
            }
        }
        System.out.println("未找到文件" + name);
        System.exit(1);
        return null;
    }
}
