package com.woslx.xlsx.NightOne;

import com.sun.xml.internal.ws.api.pipe.FiberContextSwitchInterceptor;
import com.woslx.xlsx.p2.HanyuPinyinHelper;
import net.sourceforge.pinyin4j.PinyinHelper;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

import static com.sun.jmx.snmp.ThreadContext.contains;

public class Step4 {
    public static void main(String[] args) throws IOException, InvalidFormatException, ParseException, InterruptedException {
        File ding = new File("/home/hy/tmp/nightone/dingout.xlsx");
        File yunnan = new File("/home/hy/tmp/nightone/yunnanout.xlsx");
        File shanghai = new File("/home/hy/tmp/nightone/shanghaiout.xlsx");
        String fileOutPath = "/home/hy/tmp/nightone/out/";

        Workbook wbyunnan = new XSSFWorkbook(yunnan);
        System.out.println(wbyunnan.getNumberOfSheets());
        Workbook wbShanghai = new XSSFWorkbook(shanghai);
        System.out.println(wbShanghai.getNumberOfSheets());
        Workbook wbding = new XSSFWorkbook(ding);
        System.out.println(wbding.getNumberOfSheets());

        Iterator<Sheet> sheetIterator = wbding.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet sheetding = sheetIterator.next();
            String sheetName = sheetding.getSheetName();
            Sheet sheetYunNan = findFromWorkbookByName(wbyunnan, sheetName.replaceAll(" ", "").replaceAll("-", "").replaceAll("[0123456789]", ""));
            Sheet sheetShanghai = findFromWorkbookByName(wbShanghai, sheetName.replaceAll(" ", "").replaceAll("-", "").replaceAll("[0123456789]", ""));
            Workbook wb = mergeToSheetFile(sheetding, sheetYunNan, sheetShanghai);
            FileOutputStream fos = new FileOutputStream(fileOutPath + HanyuPinyinHelper.getPinyinString(sheetName) + ".xlsx");
            wb.write(fos);
            wb.close();
            fos.close();
        }
    }

    private static Workbook mergeToSheetFile(Sheet sheetding, Sheet sheetYunNan, Sheet sheetShanghai) throws ParseException, InterruptedException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheetNew = wb.createSheet();
        sheetNew.setColumnWidth(2, 2560);
        //获取到时间,从ding的sheetName中获取日期和sheetYunNan中获取时间
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHH:mm:ss");
        SimpleDateFormat sdfx = new SimpleDateFormat("HH:mm:ss");
        String sheetNameDing = sheetding.getSheetName();
        String yyyyMMdd = sheetNameDing.substring(0, 8);
        Date date = sheetYunNan.getRow(0).getCell(2).getDateCellValue();
        Date parse = sdf.parse(yyyyMMdd + sdfx.format(date));
        Calendar instance = Calendar.getInstance();
        instance.setTime(parse);

        Row rowx = sheetNew.createRow(0);
        rowx.createCell(0).setCellValue("epoch_number");
        rowx.createCell(1).setCellValue("unix_ts");
        rowx.createCell(2).setCellValue("date");
        rowx.createCell(3).setCellValue("start_time");
        rowx.createCell(4).setCellValue("final_score");
        rowx.createCell(5).setCellValue("score1");
        rowx.createCell(6).setCellValue("score2");
        rowx.createCell(7).setCellValue("score3");
        rowx.createCell(8).setCellValue("remark");

        int lastRowNum = sheetding.getLastRowNum();
        int i = 1;
        int allSame = 0;//所有都相等
        int allNotSame = 0;//所有都不相等
        int not12 = 0;  //12不相等
        int not13 = 0;  //13不相等
        int not23 = 0;  //23不相等
        int j = 0;
        for (int cnt = 0; cnt <= lastRowNum; cnt++) {
            try {
                j++;
                Date time = instance.getTime();

                Row row = sheetNew.createRow(i++);

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

                int data1 = (int) sheetding.getRow(cnt).getCell(1).getNumericCellValue();
                Cell cell5 = row.createCell(5);     // 丁
                cell5.setCellValue(data1);

                Cell cell6 = row.createCell(6);     //云南
                int data2 = (int) sheetYunNan.getRow(cnt).getCell(1).getNumericCellValue();
                cell6.setCellValue(data2);

                Cell cell7 = row.createCell(7);     //上海
                int data3 = (int) sheetShanghai.getRow(cnt).getCell(1).getNumericCellValue();
                cell7.setCellValue(data3);

                Cell cell4 = row.createCell(4);
                if ((data1 == data2) && (data1 == data3)) {
                    allSame++;
                    cell4.setCellValue(data1);
                } else {
                    if (data1 == data2) {
                        cell4.setCellValue(data1);
                    } else if (data1 == data3) {
                        cell4.setCellValue(data1);
                    } else if (data2 == data3) {
                        cell4.setCellValue(data2);
                    } else {
                        allNotSame++;
                        cell4.setCellValue(0);
                    }
                }

                if (data1 != data2) {
                    not12++;
                }
                if (data1 != data3) {
                    not13++;
                }
                if (data2 != data3) {
                    not23++;
                }

            } catch (Exception e) {
                e.printStackTrace();
                Thread.sleep(500);
                System.exit(1);
            }
            instance.add(Calendar.SECOND, 30);

        }
        System.out.println("name:" + sheetNameDing + " 总数量:" + j);
        System.out.println("全部一样:" + allSame);
        System.out.println("全部不一样:" + allNotSame + " 差异率:" + getpersent(allNotSame, j));
        System.out.println("12不一样:" + not12 + " 差异率:" + getpersent(not12, j));
        System.out.println("13不一样:" + not13 + " 差异率:" + getpersent(not13, j));
        System.out.println("23不一样:" + not23 + " 差异率:" + getpersent(not23, j));
        System.out.println();

        return wb;
    }

    private static String getpersent(int allNotSame, int j) {
        String s = allNotSame * 1.0 / j * 100 + "";
        return s.substring(0, s.indexOf(".") + 2) + "%";
    }

    private static Sheet findFromWorkbookByName(Workbook wb, String name) {

        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet next = sheetIterator.next();
            String sheetName = next.getSheetName();
//            System.out.println(sheetName);
            if (sheetName.contains(name) ||
                    sheetName.equals(name)) {
                return next;
            }
        }

        System.out.println(getLineInfo() + " - 未找到sheet,name:" + name);
        System.exit(1);
        return null;
    }


    public static String getLineInfo() {
        StackTraceElement ste = new Throwable().getStackTrace()[1];
        return ste.getFileName() + ": Line " + ste.getLineNumber();
    }
}
