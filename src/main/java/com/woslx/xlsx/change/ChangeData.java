package com.woslx.xlsx.change;

import com.sun.xml.internal.ws.api.pipe.FiberContextSwitchInterceptor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ChangeData {
//    private static String[] dates = {"20171028", "20171029", "20171030", "20171031", "20171101", "20171102", "20171103", "20171104"};
    private static String[] dates = { "20171018", "20171019", "20171020", "20171021", "20171022", "20171023", "20171024", "20171025"};
    private static Integer[] userids = {188, 189, 190, 191, 192};


    public static void main(String[] args) throws Exception {
        File path = new File("/home/hy/tmp/changedata");
        File[] files = path.listFiles(pathname -> {
            String name = pathname.getName();
            if (!pathname.isFile()) {
                return false;
            }
            if (name.startsWith(".")) {
                return false;
            }
            if (name.endsWith(".xlsx")) {
                return true;
            }
            return false;
        });

        FileWriter fw = new FileWriter("/home/hy/tmp/changedata/out.txt");
        String[] dataset = new String[files.length];
        String[] startTime = new String[files.length];
        int num = 0;
        int[] cnt0ar = new int[files.length];
        int[] cnt1ar = new int[files.length];
        int[] cnt2ar = new int[files.length];
        int[] cnt3ar = new int[files.length];
        for (File file : files) {
            Workbook wb = new XSSFWorkbook(new FileInputStream(file));
            Sheet sheet = wb.getSheetAt(0);

            int lastRowNum = sheet.getLastRowNum();
            StringBuilder sb = new StringBuilder();
            Integer pre = null;
            int cnt0 = 0;
            int cnt1 = 0;
            int cnt2 = 0;
            int cnt3 = 0;
            for (int i = 1; i <= lastRowNum; i++) {
                if (i == 1) {
                    Date dateCellValue = sheet.getRow(i).getCell(3).getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
                    startTime[num] = sdf.format(dateCellValue);
                }
                int value = (int) sheet.getRow(i).getCell(4).getNumericCellValue();
                if (value == 0) {
                    value = pre;
                } else {
                    if (value == 5) {
                        value = 3;
                    } else if (value == 4) {
                        value = 2;
                    } else if (value == 2) {
                        value = 1;
                    }
                    else if(value == 3){
                        value = 0;
                    }
                }
                if (value == 0) {               //深睡
                    cnt0++;
                } else if (value == 1) {        //浅睡
                    cnt1++;
                } else if (value == 2) {        //rem
                    cnt2++;
                } else if (value == 3) {        //清醒
                    cnt3++;
                }

                if (i != 1) {
                    sb.append(",");
                }

                sb.append(value);

                pre = value;
            }
            cnt0ar[num] = cnt0*30;
            cnt1ar[num] = cnt1*30;
            cnt2ar[num] = cnt2*30;
            cnt3ar[num] = cnt3*30;
            dataset[num++] = "[" + sb.toString() + "]";
        }

        num = 0;
        SimpleDateFormat sdf2 = new SimpleDateFormat("yyyyMMddHH:mm:ss");
        for (String date : dates) {
            for (Integer userid : userids) {
                String sql = "insert into nurssz.sleep_data (" +
                        "user_id, date_str, start_time, data, " +
                        "deep_sleep, sleep, wake, rem, " +
                        "create_time, update_time" +
                        ") " +
                        "values(" +
                        "%s, '%s', %s, '%s', %s, %s, %s, %s,now(),now()" +
                        ");\n";
                System.out.println(date + startTime[num]);
                long time = sdf2.parse(date + startTime[num]).getTime();
                System.out.println(time);
                sql = String.format(sql, userid, date, time, dataset[num],
                        cnt0ar[num], cnt1ar[num], cnt3ar[num], cnt2ar[num]);
                fw.write(sql);
                num++;
            }
        }
        fw.flush();
        fw.close();

    }
}
