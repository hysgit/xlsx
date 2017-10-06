package com.woslx.xlsx.taskset2;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;

/**
 * Created by hy on 9/3/17.
 */
public class Task4 {
    public static void main(String[] args) throws Exception {
        String s = "/home/hy/tmp/newdata/merge.xlsx";

        FileInputStream fis = new FileInputStream(s);
        Workbook wb = new XSSFWorkbook(fis);
        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            int lastRowNum = sheet.getLastRowNum();
            int same = 0;
            int allnotsame = 0;
            int notequals34 = 0;
            int notequals35 = 0;
            int notequals45 = 0;
            for (int i = 2; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                Cell cell1 = row.getCell(1);
                Cell cell2 = row.getCell(2);
                Cell cell3 = row.getCell(3);

                if (cell1 == null || cell2 == null | cell3 == null) {
                    System.out.println("空行:" + i);
                } else {
                    Cell cell = row.createCell(4);
                    double numericCellValue3 = cell1.getNumericCellValue();
                    double numericCellValue4 = cell2.getNumericCellValue();
                    double numericCellValue5 = cell3.getNumericCellValue();
                    if((numericCellValue3 == numericCellValue4)&&
                            (numericCellValue4==numericCellValue5)){
                        same++;
                        cell.setCellValue(numericCellValue3);
                    }
                    else{
                        if(numericCellValue3==numericCellValue4){
                            notequals35++;
                            notequals45++;
                            cell.setCellValue(numericCellValue3);
                        }
                        else if(numericCellValue4== numericCellValue5){
                            notequals34++;
                            notequals35++;
                            cell.setCellValue(numericCellValue4);
                        }
                        else if(numericCellValue5==numericCellValue3){
                            notequals34++;
                            notequals45++;
                            cell.setCellValue(numericCellValue5);
                        }
                        else{
                            allnotsame++;
                            cell.setCellValue(0);
                        }
                    }
                }
            }
            System.out.println("Name:"+sheet.getSheetName());
            System.out.println("same:"+same);
            System.out.println("allnotsame:"+allnotsame);
            System.out.println("notequals23:"+notequals34);
            System.out.println("notequals24:"+notequals35);
            System.out.println("notequals34:"+notequals45);
            System.out.println();

        }

        FileOutputStream fos = new FileOutputStream("/home/hy/tmp/newdata/mergeout.xlsx");
        wb.write(fos);
    }
}
