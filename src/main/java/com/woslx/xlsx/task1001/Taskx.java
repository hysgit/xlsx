package com.woslx.xlsx.task1001;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by hy on 9/4/17.
 */
public class Taskx {
    public static void main(String[] args) throws Exception {
        Workbook wbYunNan = new XSSFWorkbook(new FileInputStream("/home/hy/tmp/newdata/云南数据out.xlsx"));
        Workbook wbDing = new XSSFWorkbook(new FileInputStream("/home/hy/tmp/newdata/ding.xlsx"));
        Workbook wbShanghai = new XSSFWorkbook(new FileInputStream("/home/hy/tmp/newdata/杭州七院20份数据睡眠分期out.xlsx"));

        Iterator<Sheet> iter = wbShanghai.sheetIterator();
        while (iter.hasNext()) {
            Sheet sheetShangHai = iter.next();
            Sheet sheetDing = getDingSheet(wbDing, sheetShangHai.getSheetName());
            Sheet sheetYunNan = getDingSheet(wbYunNan, sheetShangHai.getSheetName());
            OutputStream ops =  new FileOutputStream("/home/hy/tmp/newdata/res/" + sheetDing.getSheetName() + ".xlsx");

            Workbook wbout = new XSSFWorkbook();

            Sheet sheetOut = wbout.createSheet();
            handlerSheetAndFile(wbout,sheetOut,sheetShangHai,sheetDing,sheetYunNan);
            wbout.write(ops);
            ops.close();
            wbout.close();

        }

    }

    private static void handlerSheetAndFile(Workbook wb,Sheet sheetOut, Sheet sheetShangHai, Sheet sheetDing, Sheet sheetYunNan) throws ParseException {
        String sheetDingName = sheetDing.getSheetName();
        String ymd = sheetDingName.substring(sheetDingName.length() - 8, sheetDingName.length());
        Cell celltime = sheetShangHai.getRow(2).getCell(2);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHH:mm:ss");
        SimpleDateFormat sdfx = new SimpleDateFormat("HH:mm:ss");
        Date javaDate = HSSFDateUtil.getJavaDate(celltime.getNumericCellValue());

        Date parse = sdf.parse(ymd+sdfx.format(javaDate));
        Calendar instance = Calendar.getInstance();
        instance.setTime(parse);

        sheetOut.setColumnWidth(2, 2560);
        Row rowx = sheetOut.createRow(0);
        rowx.createCell(0).setCellValue("epoch_number");
        rowx.createCell(1).setCellValue("unix_ts");
        rowx.createCell(2).setCellValue("date");
        rowx.createCell(3).setCellValue("start_time");
        rowx.createCell(4).setCellValue("final_score");
        rowx.createCell(5).setCellValue("score1");
        rowx.createCell(6).setCellValue("score2");
        rowx.createCell(7).setCellValue("score3");
        rowx.createCell(8).setCellValue("remark");

        int lastRowNum = sheetDing.getLastRowNum();
        int i = 1;
        int allSame = 0;//所有都相等
        int allNotSame = 0;//所有都不相等
        int not12 = 0;  //12不想等
        int not13 = 0;  //13不想等
        int not23 = 0;  //23不想等
        int j = 0;
        for (int cnt=0; cnt <= lastRowNum;cnt++) {
            try {
                j++;
                Date time = instance.getTime();

                Row row = sheetOut.createRow(i++);

                Cell cell0 = row.createCell(0);
                cell0.setCellValue(i - 1);

                Cell cell2 = row.createCell(2);
                CellStyle cellStyle2 = wb.createCellStyle();
                CreationHelper createHelper2 = wb.getCreationHelper();
                cellStyle2.setDataFormat(createHelper2.createDataFormat().getFormat("m/d/yyyy"));
                cell2.setCellStyle(cellStyle2);
                cell2.setCellValue(time);


                Cell cell3 = row.createCell(3);
                CellStyle cellStyle = wb.createCellStyle();
                CreationHelper createHelper = wb.getCreationHelper();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
                cell3.setCellStyle(cellStyle);
                cell3.setCellValue(time);

                int data1 = (int) sheetDing.getRow(cnt).getCell(1).getNumericCellValue();
                Cell cell5 = row.createCell(5);     // 丁
                cell5.setCellValue(data1);

                Cell cell6 = row.createCell(6);     //云南
                int data2 = (int) sheetYunNan.getRow(22+cnt).getCell(1).getNumericCellValue();
                cell6.setCellValue(data2);

                Cell cell7 = row.createCell(7);     //上海
                int  data3 = (int) sheetShangHai.getRow(2+cnt).getCell(1).getNumericCellValue();
                cell7.setCellValue(data3);

                Cell cell4 = row.createCell(4);
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

                if(data1!=data2){
                    not12++;
                }
                if(data1 != data3){
                    not13++;
                }
                if(data2!=data3){
                    not23++;
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
            instance.add(Calendar.SECOND, 30);

        }
        System.out.println("name:" + sheetDingName + " 总数量:" + j);
        System.out.println("全部一样:" + allSame);
        System.out.println("全部不一样:" + allNotSame + " 差异率:" + getpersent(allNotSame, j));
        System.out.println("12不一样:" + not12 + " 差异率:" + getpersent(not12, j));
        System.out.println("13不一样:" + not13 + " 差异率:" + getpersent(not13, j));
        System.out.println("23不一样:" + not23 + " 差异率:" + getpersent(not23, j));
        System.out.println();
    }

    private static Sheet getDingSheet(Workbook wb, String sheetName) {
        Iterator<Sheet> iter = wb.sheetIterator();
        while (iter.hasNext()) {
            Sheet next = iter.next();
            if (next.getSheetName().contains(sheetName)) {
                return next;
            }
        }
        System.out.println("sheetName:" + sheetName + "未找到对应的sheet");
        System.exit(1);
        return null;
    }

    private static void handlerSheetAndFile(Sheet sheet, File file) throws Exception {
        String nametxt = file.getName();
        //获取日期和
        String[] split1 = nametxt.split("-");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHH:mm:ss");
        //获取起始时间
        Cell cell = sheet.getRow(2).getCell(5);
        SimpleDateFormat sdfx = new SimpleDateFormat("HH:mm:ss");
        Date javaDate = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
        String format1 = sdfx.format(javaDate);
        String s = split1[0] + split1[1] + split1[2] + format1;
        Date parse = sdf.parse(s);
        Calendar instance = Calendar.getInstance();
        instance.setTime(parse);

        //对外存储的文件名
        String fileOutName = "/home/hy/tmp/newdata/s3/" + split1[0] + "-" + split1[1] + "-" + split1[2] + split1[3] + ".xlsx";
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

        InputStream is = new FileInputStream(file);
        BufferedReader br = new BufferedReader(new InputStreamReader(is, "utf-16"));
        String line = br.readLine();
        int i = 1;
        int allSame = 0;//所有都相等
        int allNotSame = 0;//所有都不相等
        int not12 = 0;  //12不想等
        int not13 = 0;  //13不想等
        int not23 = 0;  //23不想等
        int j = 0;
        while (null != (line = br.readLine())) {
            try {
                j++;
                Date time = instance.getTime();

                Row row = sheetout.createRow(i++);

                Cell cell0 = row.createCell(0);
                cell0.setCellValue(i - 1);

                Cell cell2 = row.createCell(2);
                CellStyle cellStyle2 = wb.createCellStyle();
                CreationHelper createHelper2 = wb.getCreationHelper();
                cellStyle2.setDataFormat(createHelper2.createDataFormat().getFormat("m/d/yyyy"));
                cell2.setCellStyle(cellStyle2);
                cell2.setCellValue(time);


                Cell cell3 = row.createCell(3);
                CellStyle cellStyle = wb.createCellStyle();
                CreationHelper createHelper = wb.getCreationHelper();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
                cell3.setCellStyle(cellStyle);
                cell3.setCellValue(time);

                Integer integer = Integer.valueOf(line.split(",")[1]);
                int type = getInt(integer);
                Row row1 = sheet.getRow(i);
                Cell cell5 = row.createCell(5);     //  顶秀梅
                cell5.setCellValue(type);

                Cell cell6 = row.createCell(6); //
                double data2 = row1.getCell(2).getNumericCellValue();
                cell6.setCellValue(data2);
                Cell cell7 = row.createCell(7);     //医生三
                double data3 = row1.getCell(1).getNumericCellValue();
                cell7.setCellValue(data3);

                Cell cell4 = row.createCell(4);
                if ((type == data2) && (type == data3)) {
                    allSame++;
                    cell4.setCellValue(type);
                } else {
                    if (type == data2) {
                        cell4.setCellValue(type);
                    } else if (type == data3) {
                        cell4.setCellValue(type);
                    } else if (data2 == data3) {
                        cell4.setCellValue(data2);
                    } else {
                        allNotSame++;
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
            }
            instance.add(Calendar.SECOND, 30);

        }
        System.out.println("name:" + sheet.getSheetName() + " 总数量:" + j);
        System.out.println("全部一样:" + allSame);
        System.out.println("全部不一样:" + allNotSame + " 差异率:" + getpersent(allNotSame, j));
        System.out.println("12不一样:" + not12 + " 差异率:" + getpersent(not12, j));
        System.out.println("13不一样:" + not13 + " 差异率:" + getpersent(not13, j));
        System.out.println("23不一样:" + not23 + " 差异率:" + getpersent(not23, j));
        System.out.println();
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

    private static File getTxtByName(File[] files, String sheetName) {

        for (File file : files) {
            if (file.getName().split("-")[3].equals(sheetName)) {
                return file;
            }
        }
        System.out.println(sheetName + "未找到对应的文件");
        System.exit(1);
        return null;
    }
}
