package com.woslx.xlsx.findwake;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sun.java2d.opengl.GLXSurfaceData;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class FindWake {
    public static void main(String[] args) throws Exception {
        File dir = new File("/home/hy/tmp/p");
        File[] files = dir.listFiles();
        for (File file : files) {
            if (file.getName().startsWith(".")) {
                continue;
            }
            InputStream is = null;
            Workbook wb = null;
            try {
                is = new FileInputStream(file);
                wb = new XSSFWorkbook(is);

                Sheet sheet = wb.getSheetAt(0);
                int rowNum = sheet.getLastRowNum();
                List<WakeStatus> list = new ArrayList<>();
                WakeStatus ws = null;
                for (int i = 1; i <= rowNum; i++) {
                    Cell cell = sheet.getRow(i).getCell(4);
                    int value = (int) cell.getNumericCellValue();
                    if (5 == value) {
                        if (ws == null) {
                            ws = new WakeStatus();
                            ws.setStart(i);
                            list.add(ws);
                        }

                    } else {
                        if (ws != null && i != 1) {
                            ws.setEnd(i - 1);
                        }
                        ws = null;
                    }
                    if (ws != null && i == rowNum) {
                        ws.setEnd(i);
                        list.remove(ws);
                    }
                }

                System.out.print(file.getName().replaceAll(".xlsx", "") + ": 清醒次数:" + list.size());
                list.removeIf(next -> next.getEnd() - next.getStart() > 3);
                System.out.println(" 短时清醒:" + list.size());

                for (WakeStatus wakeStatus : list) {
                    Integer start = wakeStatus.getStart();
                    Integer end = wakeStatus.getEnd();
                    int value = (int) sheet.getRow(start - 1).getCell(4).getNumericCellValue();
                    wakeStatus.setPreStatus(value);
                    wakeStatus.setPreEnd(start - 1);
                    int temp = start - 1;
                    while (value == (int) sheet.getRow(--temp).getCell(4).getNumericCellValue()) {

                    }
                    wakeStatus.setPreStart(temp + 1);

                    value = (int) sheet.getRow(end + 1).getCell(4).getNumericCellValue();
                    wakeStatus.setAfterStatus(value);
                    wakeStatus.setAfterStart(end + 1);
                    temp = end + 1;
                    try {
                        while (value == (int) sheet.getRow(++temp).getCell(4).getNumericCellValue()) {

                        }
                    } catch (Exception e) {

                    }
                    wakeStatus.setAfterEnd(temp - 1);
                }

                for(WakeStatus wakeStatus:list){
                    System.out.println("pre状态:"+wakeStatus.getPreStatus() +",时长:"+(wakeStatus.getPreEnd()-wakeStatus.getPreStart()+1)
                    +" 清醒时长:"+(wakeStatus.getEnd()-wakeStatus.getStart()+1)+" after状态:"+wakeStatus.getAfterStatus()+",时长:"+(wakeStatus.getAfterEnd()-wakeStatus.getAfterStart()+1));
                }
                System.out.println();
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("file:" + file.getName());
            } finally {
                if (wb != null) {
                    wb.close();
                }
                if (is != null) {
                    is.close();
                }
            }
        }
    }
}
