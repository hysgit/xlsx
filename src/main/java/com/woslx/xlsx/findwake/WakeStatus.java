package com.woslx.xlsx.findwake;

import java.io.InputStream;

public class WakeStatus {
    private Integer preStatus;
    private Integer preStart;
    private Integer preEnd;

    private Integer start;  //wake起始序号
    private Integer end;    //wake结束序号

    private Integer afterStatus;
    private Integer afterStart;
    private Integer afterEnd;

    public Integer getPreStatus() {
        return preStatus;
    }

    public void setPreStatus(Integer preStatus) {
        this.preStatus = preStatus;
    }

    public Integer getPreStart() {
        return preStart;
    }

    public void setPreStart(Integer preStart) {
        this.preStart = preStart;
    }

    public Integer getPreEnd() {
        return preEnd;
    }

    public void setPreEnd(Integer preEnd) {
        this.preEnd = preEnd;
    }

    public Integer getStart() {
        return start;
    }

    public void setStart(Integer start) {
        this.start = start;
    }

    public Integer getEnd() {
        return end;
    }

    public void setEnd(Integer end) {
        this.end = end;
    }

    public Integer getAfterStatus() {
        return afterStatus;
    }

    public void setAfterStatus(Integer afterStatus) {
        this.afterStatus = afterStatus;
    }

    public Integer getAfterStart() {
        return afterStart;
    }

    public void setAfterStart(Integer afterStart) {
        this.afterStart = afterStart;
    }

    public Integer getAfterEnd() {
        return afterEnd;
    }

    public void setAfterEnd(Integer afterEnd) {
        this.afterEnd = afterEnd;
    }
}
