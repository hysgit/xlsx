package com.woslx.xlsx.task1001;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * 把txt文件,转成单个的excel
 */
public class Task0 {
    public static void main(String[] args) throws Exception {
        Path path = Paths.get("/home/hy/tmp/newdata/ding");
        File file = path.toFile();
        File[] files = file.listFiles();
        Workbook wb = new XSSFWorkbook();
        for (File filetemp : files) {
            BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(filetemp), "utf-16"));
            String[] strings = filetemp.getName().split("\\.")[0].split("-");
            String sheetName = strings[3]+strings[0]+strings[1]+strings[2];

            Sheet sheet = wb.createSheet(sheetName);
            br.readLine();
            String line = null;

            int i = 0;
            while(null != (line=br.readLine())){
                String[] split = line.split(",");
                Integer index = Integer.valueOf(split[0]);
                Integer state = getInt(Integer.valueOf(split[1]));
                Row row = sheet.createRow(i++);
                Cell cell0 = row.createCell(0);
                cell0.setCellValue(index);
                Cell cell1 = row.createCell(1);
                cell1.setCellValue(state);
            }

            FileOutputStream fos = new FileOutputStream("/home/hy/tmp/newdata/ding.xlsx");
            wb.write(fos);


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
}
