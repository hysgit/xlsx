package com.woslx.xlsx.NightOne;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.java2d.opengl.GLXSurfaceData;

import java.io.*;

/**
 * 转换丁医生的数据到一个execl中,sheet名是文件名
 */
public class Step1 {
    public static void main(String[] args) throws Exception {
        File file = new File("/home/hy/tmp/nightone/ding");
        File[] files = file.listFiles(new FileFilter() {
            @Override
            public boolean accept(File pathname) {
                String name = pathname.getName();
                if (!pathname.isFile()) {
                    return false;
                }
                if (name.startsWith(".")) {
                    return false;
                }
                if (name.endsWith(".txt")) {
                    return true;
                }
                return false;
            }
        });
        Workbook wb = new XSSFWorkbook();
        for (File ftemp : files) {
            file2sheet(ftemp, wb);
        }

        OutputStream os = new FileOutputStream("/home/hy/tmp/nightone/dingout.xlsx");
        wb.write(os);
        wb.close();
        os.close();


    }

    private static void file2sheet(File ftemp, Workbook wb) throws Exception {
        String name = ftemp.getName();
        System.out.println(name);
        Sheet sheet = wb.createSheet(getSheetNameFromFile(ftemp));
        if (!name.contains("陆莹")) {
            BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(ftemp), "utf-16"));
            br.readLine();      //去除第一行
            String line = null;
            int i = 0;
            while (null != (line = br.readLine())) {
                Row row = sheet.createRow(i++);
                String[] split = line.split(",");

                Cell cell0 = row.createCell(0);
                Cell cell1 = row.createCell(1);
                cell0.setCellValue(Integer.valueOf(split[0]));
                Integer value = getInt(Integer.valueOf(split[1]));
                cell1.setCellValue(value);
            }
        } else {
            Workbook wb2 = new XSSFWorkbook("/home/hy/tmp/nightone/ding/丁秀梅.xlsx");
            Sheet sheetx = wb2.getSheet("陆莹9-11");
            int lastRowNum = sheetx.getLastRowNum();
            int x = 0;
            for (int i = 22; i <= lastRowNum; i++) {
                Row row = sheet.createRow(x);
                Cell cell0 = row.createCell(0);
                Cell cell1 = row.createCell(1);
                Row rowx = sheetx.getRow(i);

                cell0.setCellValue(getIntFromCell(rowx.getCell(0)));
                cell1.setCellValue(getIntFromCell(rowx.getCell(1)));
                x++;
            }
        }
    }

    public static int getIntFromCell(Cell cell){
        int cellType = cell.getCellType();
        if(cellType == Cell.CELL_TYPE_STRING){
            String stringCellValue = cell.getStringCellValue();
            if(stringCellValue.contains("n拺")){
                return 5;
            }
            else if(stringCellValue.contains("清醒")){
                return 5;
            }
            else if(stringCellValue.contains("N1")){
                return 1;
            }
            else if(stringCellValue.contains("N2")){
                return 2;
            }
            else if(stringCellValue.contains("N3")){
                return 3;
            }
            else if(stringCellValue.contains("REM")){
                return 4;
            }
            else{
                System.out.println("未知字符串:"+stringCellValue);
                System.exit(1);
                return 0;
            }
        }
        else if(cellType == Cell.CELL_TYPE_NUMERIC){
            return (int) cell.getNumericCellValue();
        }
        else{
            System.out.println("类型:"+cellType);
            System.exit(1);
            return 0;
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

    private static String getSheetNameFromFile(File ftemp) {
        String name = ftemp.getName();
        String[] split = name.split("-");

        return split[0] + split[1] + split[2] + split[3];
    }
}
