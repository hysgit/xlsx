package com.woslx.xlsx.task1127;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Calendar;
import java.util.Date;

/**
 * Created by hy on 8/21/17.
 * 把T1和T4进行转换
 */
public class Task1_txt {
    public static void main(String[] args) throws IOException {
        //加载txt文件
        String input = "/home/hy/Downloads/27/11.23-11.28T4/王伟玲-11.23.txt";
        String output = "/home/hy/Downloads/27/11.23-11.28T4/2017-11-23-wangweiling.xlsx";

        try {
            FileInputStream fis = new FileInputStream(input);
            BufferedReader br = new BufferedReader(new InputStreamReader(fis,"utf-16"));

            Workbook wbout = new XSSFWorkbook();
            Sheet sheetOut = wbout.createSheet();

            Calendar calendar = Calendar.getInstance();
            calendar.set(Calendar.MONTH, Calendar.DECEMBER);
            calendar.set(Calendar.DAY_OF_MONTH, 23);
            calendar.set(Calendar.HOUR_OF_DAY,20);
            calendar.set(Calendar.MINUTE,19);
            calendar.set(Calendar.SECOND,57);
            calendar.set(Calendar.MILLISECOND,0);
            Date startTime = calendar.getTime();
            startTime.setYear(117);
            startTime.setMonth(10);
            startTime.setDate(23);
            save(wbout, br, sheetOut, startTime);
            fis.close();
            FileOutputStream fos = new FileOutputStream(output);
            wbout.write(fos);
            fos.flush();
            fos.close();
            br.close();
            fis.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void save(Workbook wb, BufferedReader br, Sheet sheetOut, Date startTime) throws IOException {
        Row rowx = sheetOut.createRow(0);
        rowx.createCell(0).setCellValue("epoch_number");
        rowx.createCell(1).setCellValue("unix_ts");
        rowx.createCell(2).setCellValue("date");
        rowx.createCell(3).setCellValue("start_time");
        rowx.createCell(4).setCellValue("score1");
        rowx.createCell(5).setCellValue("remark");

        Calendar instance = Calendar.getInstance();
        instance.setTime(startTime);

        String line = br.readLine();//丢弃第一行
        int i = 1;
        while(null != (line = br.readLine()))
        {
            String[] split = line.split(",");
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

            Cell cell4 = row.createCell(4);     //  顶秀梅
            cell4.setCellValue(getInt(Integer.valueOf(split[1])));

            instance.add(Calendar.SECOND, 30);
        }
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

    private static int hcell(Cell cell) {
        if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
            String stringCellValue = cell.getStringCellValue().trim();
            if (("n拺").equals(stringCellValue)) {
                return 5;
            } else if (("N1").equals(stringCellValue)) {
                return 1;
            } else if (("N2").equals(stringCellValue)) {
                return 2;
            } else if (("N3").equals(stringCellValue)) {
                return 3;
            } else if (("REM").equals(stringCellValue)) {
                return 4;
            } else if (("清醒").equals(stringCellValue)) {
                return 5;
            } else {
                System.exit(1);
                System.out.println("发生未知类型:" + stringCellValue);
                return 1;
            }
        } else {
            System.out.println("发生错误");
            System.exit(1);
            return 1;
        }
    }
}
