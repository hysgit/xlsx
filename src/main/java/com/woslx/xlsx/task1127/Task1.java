package com.woslx.xlsx.task1127;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by hy on 8/21/17.
 * 把T1和T4进行转换
 */
public class Task1 {
    public static void main(String[] args) throws IOException {
        //加载txt文件
        String input = "/home/hy/Downloads/27/11.23-28/2017-11-24-王小芳.xlsx";
        String output = "/home/hy/Downloads/27/11.23-28/2017-11-24-wangxiaofang.xlsx";

        try {
            FileInputStream fis = new FileInputStream(input);
            Workbook wbInput = new XSSFWorkbook(fis);
            Sheet sheetInput = wbInput.getSheetAt(0);
            Workbook wbout = new XSSFWorkbook();
            Sheet sheetOut = wbout.createSheet();

            Date startTime = getDateTime(sheetInput);
            startTime.setYear(117);
            startTime.setMonth(10);
            startTime.setDate(24);
            save(wbout,sheetInput,sheetOut,startTime);
            fis.close();
            FileOutputStream fos = new FileOutputStream(output);
            wbout.write(fos);
            fos.flush();
            fos.close();
            wbInput.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void save(Workbook wb, Sheet sheetInput, Sheet sheetOut, Date startTime) {
        Row rowx = sheetOut.createRow(0);
        rowx.createCell(0).setCellValue("epoch_number");
        rowx.createCell(1).setCellValue("unix_ts");
        rowx.createCell(2).setCellValue("date");
        rowx.createCell(3).setCellValue("start_time");
        rowx.createCell(4).setCellValue("score1");
        rowx.createCell(5).setCellValue("remark");

        Calendar instance = Calendar.getInstance();
        instance.setTime(startTime);


        int i= 1;
        int lastRowNum = sheetInput.getLastRowNum();
        for(int x = 22; x <= lastRowNum;x++) {
            Date time = instance.getTime();
            Row row = sheetOut.createRow(i++);
            Row rowInput = sheetInput.getRow(x);
            Cell cell = rowInput.getCell(1);
            Integer value = hcell(cell);
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
            cell4.setCellValue(value);

            instance.add(Calendar.SECOND, 30);
        }
    }

    private static Date getDateTime(Sheet sheetInput) {
        Row row = sheetInput.getRow(22);
        Date date = HSSFDateUtil.getJavaDate(row.getCell(3).getNumericCellValue());
        return date;
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
            }
            else{
                System.exit(1);
                System.out.println("发生未知类型:"+stringCellValue);
                return 1;
            }
        }
        else{
            System.out.println("发生错误");
            System.exit(1);
            return 1;
        }
    }
}
