package com.woslx.xlsx.NightOne;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Step2 {
    public static void main(String[] args) throws Exception {
        File path = new File("/home/hy/tmp/nightone/yunnan36人数据");
        path.listFiles(pathname -> {
            String name = pathname.getName();
            if (!pathname.isFile()) {
                return false;
            }
            if(name.startsWith(".")){
                return false;
            }
            if(name.endsWith(".xlsx"))
            {
                return true;
            }
            return false;
        });


    }
}
