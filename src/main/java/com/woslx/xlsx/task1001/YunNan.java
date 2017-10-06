package com.woslx.xlsx.task1001;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;

/**
 * Created by hy on 10/1/17.
 */
public class YunNan {
    public static void main(String[] args) throws Exception {
        File file = new File("/home/hy/tmp/sing/res");
        File[] files = file.listFiles();
        FileInputStream fis = new FileInputStream("/home/hy/tmp/sing/云南数据out.xlsx");

        Workbook wb = new XSSFWorkbook(fis);
        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while(sheetIterator.hasNext()){
            Sheet next = sheetIterator.next();

            String sheetName = next.getSheetName();
            File fileName = findFileName(sheetName,files);

            FileInputStream fileInputStream = new FileInputStream(fileName);
            Workbook wbtemp = new XSSFWorkbook(fileInputStream);
            Sheet sheetfinal = wbtemp.getSheetAt(0);
            int lastRowNum = next.getLastRowNum();
            int x = 0;
            for (int i = 22; i <= lastRowNum; i ++) {
                Cell cellyunan = next.getRow(i).getCell(1);
                try {
                    Cell cell = sheetfinal.getRow(i - 21).getCell(6);
                    cell.setCellValue(cellyunan.getNumericCellValue());
                    x++;
                } catch (Exception e) {
                    System.out.println("fileName:"+fileName+" x:"+x+" lastRowNum"+lastRowNum);
                    System.exit(1);
                }
            }
            System.out.println("filename:"+fileName+" line:"+x);
            fileInputStream.close();
            FileOutputStream fileOutputStream = new FileOutputStream(fileName);
            wbtemp.write(fileOutputStream);
            fileOutputStream.close();
            wbtemp.close();
        }


    }

    private static File findFileName(String sheetName, File[] files) {
        for (File file : files) {
            String name = file.getName();
            String s = sheetName.replaceAll("[0-9]", "");
            if (name.contains(s)) {
                return file;
            }

        }
        return null;
    }
}
