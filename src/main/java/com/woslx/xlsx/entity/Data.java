package com.woslx.xlsx.entity;

import com.woslx.xlsx.XlsxMain;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by hy on 7/13/17.
 */
public class Data {
    private long timestamp;
    private String dateStr;
    private Date   date;
    private String timeStr;
    private String status;

    public Data(String timeStr, String status) {
        this.timeStr = timeStr;
        this.status = status;

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHH:mm:ss");
        if(Integer.valueOf(timeStr.split(":")[0]) > 12){
            //2017.07.11
            String day11 = XlsxMain.today + timeStr;
            try {
                date = sdf.parse(day11);
            } catch (ParseException e) {
                e.printStackTrace();
            }

        }
        else{
            //07.12
            String day12 = XlsxMain.nextday + timeStr;
            try {
                date = sdf.parse(day12);
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
        timestamp = date.getTime();
        
        setDateStr(date);
    }

    private void setDateStr(Date date) {
        dateStr = new SimpleDateFormat("yyyy.MM.dd HH:mm:ss").format(date);
    }

    public long getTimestamp() {
        return timestamp;
    }

    public void setTimestamp(long timestamp) {
        this.timestamp = timestamp;
        Date date = new Date();
        date.setTime(timestamp);
        this.date = date;
        setDateStr(date);
        this.setTimeStr(new SimpleDateFormat("HH:mm:ss").format(date));
    }

    public String getDateStr() {
        return dateStr;
    }

    public void setDateStr(String dateStr) {
        this.dateStr = dateStr;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getTimeStr() {
        return timeStr;
    }

    public void setTimeStr(String timeStr) {
        this.timeStr = timeStr;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }
}
