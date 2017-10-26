package com.woslx.xlsx.NightOne;

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
                if(name.startsWith(".")){
                    return false;
                }
                if(name.endsWith(".txt"))
                {
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
        wb.createSheet(getSheetNameFromFile(ftemp));

        BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(ftemp), "utf-16"));
        br.readLine();      //去除第一行
        String line = null;
        while(null != (line = br.readLine())){
            String[] split = line.split(",");

        }
    }

    private static String getSheetNameFromFile(File ftemp) {
        String name = ftemp.getName();
        String[] split = name.split("-");

        return split[0]+split[1]+split[2]+split[3];
    }
}
