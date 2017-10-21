package com.woslx.xlsx.p2;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Find3to5 {
    public static void main(String[] args) throws Exception {
        File dir = new File("/home/hy/tmp/p");
        File[] files = dir.listFiles();
        for (File file : files) {
            if(file.getName().startsWith("."))
            {
                continue;
            }
            InputStream is = null;
            Workbook wb = null;
            try {
                is = new FileInputStream(file);
                wb = new XSSFWorkbook(is);

                Sheet sheet = wb.getSheetAt(0);
                int rowNum = sheet.getLastRowNum();

                int c5 = 0;
                int c6 = 0;
                int c7 = 0;

                int cnt5 = 0;
                int cnt6 = 0;
                int cnt7 = 0;

                for (int i = 1; i <= rowNum; i++) {
                    Row row = sheet.getRow(i);
                    int v5 = (int) row.getCell(5).getNumericCellValue();
                    int v6 = (int) row.getCell(6).getNumericCellValue();
                    int v7 = (int) row.getCell(7).getNumericCellValue();

                    if (i > 1) {
                        if (c5 == 3 && v5 == 5) {
                            cnt5++;
                        }
                        if (c6 == 3 && v6 == 5) {
                            cnt6++;
                        }
                        if (c7 == 3 && v7 == 5) {
                            cnt7++;
                        }
                    }

                    c5 = v5;
                    c6 = v6;
                    c7 = v7;
                }
                System.out.println(file.getName());
                System.out.println("医生1:" + cnt5);
                System.out.println("医生2:" + cnt6);
                System.out.println("医生3:" + cnt7);
                System.out.println();
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("file:"+file.getName());
            }
            finally {
                if(wb!= null){
                    wb.close();
                }
                if(is!=null){
                    is.close();
                }
            }
        }
    }
}
