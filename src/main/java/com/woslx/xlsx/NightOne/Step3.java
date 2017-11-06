package com.woslx.xlsx.NightOne;

import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * 上海数据
 */


public class Step3 {
    public static void main(String[] args) {

    }

    public static int getIntFromCell(Cell cell){
        int cellType = cell.getCellType();
        if(cellType == Cell.CELL_TYPE_STRING){
            String stringCellValue = cell.getStringCellValue();
            if(stringCellValue.contains("n拺")){
                return 5;
            }
            else if(stringCellValue.contains("清醒")){
                return 5;
            }
            else if(stringCellValue.contains("N1")){
                return 1;
            }
            else if(stringCellValue.contains("N2")){
                return 2;
            }
            else if(stringCellValue.contains("N3")){
                return 3;
            }
            else if(stringCellValue.contains("REM")){
                return 4;
            }
            else{
                System.out.println("未知字符串:"+stringCellValue);
                System.exit(1);
                return 0;
            }
        }
        else if(cellType == Cell.CELL_TYPE_NUMERIC){
            return (int) cell.getNumericCellValue();
        }
        else{
            System.out.println("类型:"+cellType);
            System.exit(1);
            return 0;
        }
    }
}
