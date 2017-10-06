package com.woslx.xlsx.taskset2;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by hy on 9/4/17.
 */
public class Taskx {
    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream("/home/hy/tmp/newdata/mergeout.xlsx");
        FileInputStream yunnan = new FileInputStream("/home/hy/tmp/newdata/云南数据out.xlsx");

        Workbook wb = new XSSFWorkbook(fis);        //mergeout
        Workbook yunnanwb = new XSSFWorkbook(yunnan);        //mergeout

        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            String sheetName = sheet.getSheetName();
            Sheet sheetyunnan = getTxtByName(yunnanwb, sheetName);
            handlerSheetAndFile(sheet, sheetyunnan);
            System.out.println("上海:" + sheet.getSheetName() + "  云南:" + sheetyunnan.getSheetName());
            System.out.println();
        }
    }

    private static void handlerSheetAndFile(Sheet sheetshanghai, Sheet sheetYun) throws Exception {
        //获取日期和
        String datestr = sheetYun.getSheetName().substring(sheetYun.getSheetName().length() - 8, sheetYun.getSheetName().length());
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHH:mm:ss");
        //获取起始时间
        Cell cell = sheetshanghai.getRow(2).getCell(5);
        SimpleDateFormat sdfx = new SimpleDateFormat("HH:mm:ss");
        Date javaDate = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
        String format1 = sdfx.format(javaDate);
        String s = datestr + format1;
        Date parse = sdf.parse(s);
        Calendar instance = Calendar.getInstance();
        instance.setTime(parse);

        //对外存储的文件名
        String fileOutName = "/home/hy/tmp/newdata/s3/" + sheetYun.getSheetName() + ".xlsx";
        OutputStream ops = new FileOutputStream(fileOutName);

        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheetout = wb.createSheet();
        sheetout.setColumnWidth(2, 2560);
        Row rowx = sheetout.createRow(0);
        rowx.createCell(0).setCellValue("epoch_number");
        rowx.createCell(1).setCellValue("unix_ts");
        rowx.createCell(2).setCellValue("date");
        rowx.createCell(3).setCellValue("start_time");
        rowx.createCell(4).setCellValue("final_score");
        rowx.createCell(5).setCellValue("score1");
        rowx.createCell(6).setCellValue("score2");
        rowx.createCell(7).setCellValue("score3");
        rowx.createCell(8).setCellValue("remark");


        int i = 1;
        int allSame = 0;//所有都相等
        int allNotSame = 0;//所有都不相等
        int not12 = 0;  //12不想等
        int not13 = 0;  //13不想等
        int not23 = 0;  //23不想等
        int j = 0;
        int lastRowNum = sheetshanghai.getLastRowNum();
        for(int linenum = 2;linenum<=lastRowNum;linenum++){

            try {
                j++;
                Date time = instance.getTime();

                Row rowOut = sheetout.createRow(i++);

                Cell cell0 = rowOut.createCell(0);
                cell0.setCellValue(i - 1);

                Cell cell2 = rowOut.createCell(2);
                CellStyle cellStyle2 = wb.createCellStyle();
                CreationHelper createHelper2 = wb.getCreationHelper();
                cellStyle2.setDataFormat(createHelper2.createDataFormat().getFormat("m/d/yyyy"));
                cell2.setCellStyle(cellStyle2);
                cell2.setCellValue(time);


                Cell cell3 = rowOut.createCell(3);
                CellStyle cellStyle = wb.createCellStyle();
                CreationHelper createHelper = wb.getCreationHelper();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
                cell3.setCellStyle(cellStyle);
                cell3.setCellValue(time);

                int type = (int) sheetYun.getRow(linenum+20).getCell(1).getNumericCellValue();
                Row row1 = sheetshanghai.getRow(i);
                Cell cell5 = rowOut.createCell(5);     //  云南
                cell5.setCellValue(type);

                Cell cell6 = rowOut.createCell(6); //      刘医生
                double data2 = row1.getCell(2).getNumericCellValue();
                cell6.setCellValue(data2);
                Cell cell7 = rowOut.createCell(7);     //上海医生
                double data3 = row1.getCell(1).getNumericCellValue();
                cell7.setCellValue(data3);

                Cell cell4 = rowOut.createCell(4);
                if ((type == data2) && (type == data3)) {
                    allSame++;      //全部一样
                    cell4.setCellValue(type);
                } else {
                    if (type == data2) {
//                        not13++;
//                        not23++;
                        cell4.setCellValue(type);
                    } else if (type == data3) {
//                        not12++;
//                        not23++;
                        cell4.setCellValue(type);
                    } else if (data2 == data3) {
//                        not12++;
//                        not13++;
                        cell4.setCellValue(data2);
                    } else {
                        allNotSame++;       //全部不一样
                        cell4.setCellValue(0);
                    }
                }
                if(type!=data2){
                    not12++;
                }
                if(type != data3){
                    not13++;
                }
                if(data2!=data3){
                    not23++;
                }

            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("sheetshanghai.getSheetName:"+sheetshanghai.getSheetName());
                System.out.println("sheetYun.getSheetName:"+sheetYun.getSheetName());
                System.out.println("linenum:"+linenum);
                System.out.println("lastRowNum:"+lastRowNum);
                System.exit(1);
            }
            instance.add(Calendar.SECOND, 30);

        }
        System.out.println("name:" + sheetshanghai.getSheetName() + " 总数量:" + j);
        System.out.println("全部一样:" + allSame);
        System.out.println("全部不一样:" + allNotSame + " 差异率:" + getpersent(allNotSame, j));
        System.out.println("12不一样:" + not12 + " 差异率:" + getpersent(not12, j));
        System.out.println("13不一样:" + not13 + " 差异率:" + getpersent(not13, j));
        System.out.println("23不一样:" + not23 + " 差异率:" + getpersent(not23, j));
        wb.write(ops);
        ops.close();
        wb.close();
    }

    private static String getpersent(int allNotSame, int j) {
        String s = allNotSame * 1.0 / j * 100 + "";
        return s.substring(0, s.indexOf(".") + 2) + "%";
    }

    private static int getInt(Integer integer) {

        if (integer == 10) {
            return 5;
        } else if (integer == 1) {
            return 1;
        } else if (integer == 2) {
            return 2;
        } else if (integer == 3) {
            return 3;
        } else if (integer == 5) {
            return 4;
        } else {
            System.out.println("有未知的类型:" + integer);
            System.exit(1);
            return 0;
        }
    }

    private static Sheet getTxtByName(Workbook wb, String sheetName) {

        Iterator<Sheet> iter = wb.sheetIterator();
        while (iter.hasNext()) {
            Sheet sheet = iter.next();
            if (sheet.getSheetName().contains(sheetName)) {
                return sheet;
            }
        }
        System.out.println(sheetName + "未找到对应的sheet");
        System.exit(1);
        return null;
    }
}
