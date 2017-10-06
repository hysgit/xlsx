package com.woslx.xlsx.sub;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by hy on 8/21/17.
 */
public class Task4 {
    public static void main(String[] args) throws IOException {
        //加载txt文件
//        String filepath = "/home/hy/Downloads/第二晚转换数据整合-0815-再转换/7-11-黄跃.xlsx";
//        String fileout = "/home/hy/Downloads/第二晚转换数据整合-0815-再转换/7-11-黄跃out.xlsx";

        String filepath = "/home/hy/Downloads/第二晚转换数据整合-0815-再转换/7-13-张冰珂.xlsx";
        String fileout = "/home/hy/Downloads/第二晚转换数据整合-0815-再转换/7-13-张冰珂out.xlsx";

        try {
            FileInputStream fis = new FileInputStream(filepath);
            Workbook wb = new XSSFWorkbook(fis);


            Sheet sheet = wb.getSheetAt(0);
            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();
            int same = 0;
            int notsame = 0;
            int not = 0;
            int not1 = 0;
            int not2 = 0;
            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                Cell cell4 = row.getCell(4);
                Cell cell5 = row.getCell(5);
                Cell cell6 = row.getCell(6);
                Cell cell7 = row.getCell(7);

                if (cell4 == null || cell5 == null || cell6 == null || cell7 == null) {
                    System.out.println(i);
                } else {
                    if (cell4.getCellType() == Cell.CELL_TYPE_NUMERIC &&
                            cell5.getCellType() == Cell.CELL_TYPE_NUMERIC &&
                            cell6.getCellType() == Cell.CELL_TYPE_NUMERIC
                            ) {
                        double c4 = cell4.getNumericCellValue();
                        double c5 = cell5.getNumericCellValue();
                        double c6 = cell6.getNumericCellValue();
                        if ((c4 == c5) && (c5 == c6)) {
                            same++;
                            cell7.setCellValue(c4);
                        } else {
                            if(c4!=c5){
                                not++;
                            }
                            if (c4 != c6) {
                                not1++;
                            }
                            if (c5 != c6) {
                                not2++;
                            }
                            if (c4 == c5) {
                                cell7.setCellValue(c4);
                            } else if (c4 == c6) {
                                cell7.setCellValue(c4);
                            } else if (c5 == c6) {
                                cell7.setCellValue(c5);
                            } else {
                                notsame++;
                                cell7.setCellValue(0);
//                                System.out.println("line:"+i);
//                                System.exit(1);
                            }
                        }
                    } else {
                        System.out.println("有非数字:" + i);
                        System.exit(1);
                    }
                }
            }
            System.out.println("same:" + same);
            System.out.println("notsame:" + notsame+" "+notsame*1.0/1230);
            System.out.println("not1:" + not1+" "+not1*1.0/1230);
            System.out.println("not2:" + not2+" "+not2*1.0/1230);
            System.out.println("not:" + not+" "+not*1.0/1230);

            fis.close();
            FileOutputStream fos = new FileOutputStream(fileout);
            wb.write(fos);
            fos.flush();
            fos.close();
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
