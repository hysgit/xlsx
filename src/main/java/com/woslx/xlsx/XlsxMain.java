package com.woslx.xlsx;

import com.woslx.xlsx.entity.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Hello world!
 */
public class XlsxMain {

    public static String today;
    public static String nextday;

    private String inputFilePath;
    private String outputFilePath;

    public String getInputFilePath() {
        return inputFilePath;
    }

    public void setInputFilePath(String inputFilePath) {
        this.inputFilePath = inputFilePath;
    }

    public String getOutputFilePath() {
        return outputFilePath;
    }

    public void setOutputFilePath(String outputFilePath) {
        this.outputFilePath = outputFilePath;
    }

    public void start() throws Exception {
        //加载txt文件
        InputStream inputStream = new FileInputStream(inputFilePath);
        BufferedReader br = new BufferedReader(new InputStreamReader(inputStream, Charset.forName("UTF16")));

        List<Data> list = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
        String line = null;
        while(null != (line = br.readLine())){
            try {
                String[] split = line.split(",");
                String timeStr = split[0];
                String status = split[2];

                list.add(new Data(timeStr, status));
            }
            catch (Exception e){
                System.out.println(line);
                e.printStackTrace();
            }
        }
        List<Data> outList = new ArrayList<>();

        delSame(list);
        to0and30(list);
        long timestamp = list.get(0).getTimestamp();
        long timenow = to0or30(timestamp);
        for (int i = 0; i < list.size() - 1; i++) {
            Data datanow = list.get(i);
            Data next = list.get(i + 1);

            do {

                Date date = new Date();
                date.setTime(timenow);
                Data save = new Data(new SimpleDateFormat("HH:mm:ss").format(date), datanow.getStatus());
                outList.add(save);
                timenow += 30000;
            }
            while (timenow < next.getTimestamp());
        }

        save(outList);
        System.out.println();

    }

    private void to0and30(List<Data> list) {
        for (int i = 0; i < list.size() - 1; i++) {
            Data pre = list.get(i);
            Data next = list.get(i + 1);
            long changed = to0or30(pre.getTimestamp());
            long nextChanged = to0or30(next.getTimestamp());
            if (changed == nextChanged) {
                nextChanged += 30;
            }
            pre.setTimestamp(changed);
            next.setTimestamp(nextChanged);
        }

    }

    private void save(List<Data> outList) throws Exception {

//        FileOutputStream fileOut = new FileOutputStream("/home/hy/tmp/2017-07-11-黄跃-评分事件-丁秀梅时间修正.xlsx");
        FileOutputStream fileOut = new FileOutputStream(outputFilePath);
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        sheet.setColumnWidth(2,2560);
        //epoch_number	unix_ts	date	start_time 	丁秀梅	final_score	score1	score2	remark
        Row rowx = sheet.createRow(0);
        rowx.createCell(0).setCellValue("epoch_number");
        rowx.createCell(1).setCellValue("unix_ts");
        rowx.createCell(2).setCellValue("date");
        rowx.createCell(3).setCellValue("start_time");
        rowx.createCell(4).setCellValue("丁秀梅");
        rowx.createCell(5).setCellValue("final_score");
        rowx.createCell(6).setCellValue("score1");
        rowx.createCell(7).setCellValue("score2");
        rowx.createCell(8).setCellValue("remark");

        int i = 1;
        for (Data data : outList) {
            Row row = sheet.createRow(i++);

            Cell cell2 = row.createCell(2);
            CellStyle cellStyle2 = wb.createCellStyle();
            CreationHelper createHelper2 = wb.getCreationHelper();
            cellStyle2.setDataFormat(createHelper2.createDataFormat().getFormat("m/d/yyyy"));
            cell2.setCellStyle(cellStyle2);
            cell2.setCellValue(data.getDate());


            Cell cell3 = row.createCell(3);
            CellStyle cellStyle = wb.createCellStyle();
            CreationHelper createHelper = wb.getCreationHelper();
            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm:ss"));
            cell3.setCellStyle(cellStyle);
            cell3.setCellValue(data.getDate());

            Cell cell4 = row.createCell(4);
            cell4.setCellValue(data.getStatus());

        }
        wb.write(fileOut);
        fileOut.close();
        wb.close();
    }

    private long to0or30(long timestamp) {
        Date date = new Date();
        date.setTime(timestamp);
        int seconds = date.getSeconds();
        if (seconds >= 16 && seconds <= 45) {
            date.setSeconds(30);
        } else {
            date.setSeconds(0);
            if(seconds>45){
                date.setMinutes(date.getMinutes()+1);
            }
        }
        return date.getTime();
    }


    private void delSame(List<Data> outList) {
        Data first = null;
        Data second = null;

        boolean ac = true;
        while (ac) {
            ac = false;
            for (int i = 0; i < outList.size(); i++) {
                Data data = outList.get(i);
                if (first == null) {
                    first = data;
                } else {
                    second = data;
                    if (!second.getStatus().equals(first.getStatus())) {
                        first = second;
                        second = null;
                        second = null;
                    } else {
                        if (i != outList.size() - 1) {
                            outList.remove(i);
                            first = null;
                            second = null;
                            ac = true;
                        }
                        break;
                    }
                }
            }
        }
    }

    public static void main(String[] args) throws ParseException {
        if(args.length ==0){
            System.out.println("缺少文件,请输入完整文件名,包括扩展名");
            System.exit(1);
        }
        System.out.println(Arrays.deepToString(args));
        Date today = new SimpleDateFormat("yyyy-MM-dd").parse(args[0].substring(0, 10));
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(today);
        calendar.add(Calendar.DATE,1);
        Date nextDay = calendar.getTime();
        SimpleDateFormat yyyyMMdd = new SimpleDateFormat("yyyyMMdd");
        XlsxMain.today = yyyyMMdd.format(today);
        XlsxMain.nextday = yyyyMMdd.format(nextDay)+"0";

        String path = System.getProperty("user.dir");
        path = path+"/"+args[0];
        System.out.println(path);
        File file = new File(path);
        if(!file.exists()){
            System.out.println("文件不存在,请输入完整文件名,包括扩展名");
            System.exit(1);
        }


        try {
            XlsxMain xlsxMain = new XlsxMain();
            xlsxMain.setInputFilePath(path);
            xlsxMain.setOutputFilePath(path.substring(0,path.indexOf('.'))+"时间修正.xlsx");
            xlsxMain.start();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
