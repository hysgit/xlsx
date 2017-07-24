package com.woslx.xlsx;

import org.apache.poi.ddf.EscherColorRef;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by hy on 7/19/17.
 */
public class Solve {
    public static void main(String[] args) throws Exception {
        String path = "/home/hy/tmp/a.txt";
        BufferedReader br = new BufferedReader( new InputStreamReader(new FileInputStream(path)));

        Pattern pattern = Pattern.compile("index=(\\d{1,3})");

        List<Integer> list = new ArrayList<>();
        String line;
        while(null != (line = br.readLine())){
            Matcher matcher = pattern.matcher(line);
            if (matcher.find()) {
                String group = matcher.group();

                Integer index = Integer.valueOf(group.split("=")[1]);
                list.add(index);
            }
        }

        Integer pre = null;
        Integer cnt = 0;
        Integer seq = 0;
        for (Integer indexnow : list) {
            seq++;
            if(pre != null){
                if(indexnow != 1){
                    if(indexnow - pre != 1){
                        cnt++;
                        System.out.println("seq: "+seq+" now: "+indexnow+" pre:"+ pre);
                    }
                }
                else if(pre!=255){
                    System.out.println("seq: "+seq+" now: "+indexnow+" pre:"+ pre);
                    cnt++;
                }
            }
            pre = indexnow;
        }
        System.out.println(cnt);
    }
}
