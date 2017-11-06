package com.woslx.xlsx.NightOne;

import com.sun.xml.internal.ws.api.pipe.FiberContextSwitchInterceptor;
import com.woslx.xlsx.p2.HanyuPinyinHelper;
import net.sourceforge.pinyin4j.PinyinHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import static com.sun.jmx.snmp.ThreadContext.contains;

public class Step4 {
    public static void main(String[] args) throws IOException, InvalidFormatException {
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
            Workbook wb = mergeToSheetFile(sheetding,sheetYunNan,sheetShanghai);
            FileOutputStream fos = new FileOutputStream(fileOutPath+HanyuPinyinHelper.getPinyinString(sheetName)+".xlsx");
            wb.write(fos);
            wb.close();
            fos.close();
        }
    }

    private static Workbook mergeToSheetFile(Sheet sheetding, Sheet sheetYunNan, Sheet sheetShanghai) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheetNew = wb.createSheet();
        return wb;
    }

    private static Sheet findFromWorkbookByName(Workbook wb, String name) {

        Iterator<Sheet> sheetIterator = wb.sheetIterator();
        while(sheetIterator.hasNext()){
            Sheet next = sheetIterator.next();
            String sheetName = next.getSheetName();
            System.out.println(sheetName);
            if(sheetName.contains(name) ||
                    sheetName.equals(name)){
                return next;
            }
        }

        System.out.println(getLineInfo()+" - 未找到sheet,name:"+name);
        System.exit(1);
        return null;
    }


    public static String getLineInfo()
    {
        StackTraceElement ste = new Throwable().getStackTrace()[1];
        return ste.getFileName() + ": Line " + ste.getLineNumber();
    }
}
