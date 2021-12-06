package org.imc.service.nancy.excel;

import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Component
public class ExcelData {
    private XSSFSheet sheet;

    ExcelData(){

    }
    /**
     * 构造函数，初始化excel数据
     * @param filePath  excel路径
     * @param sheetName sheet表名
     */
    public ExcelData(String filePath, String sheetName){
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(filePath);
            XSSFWorkbook sheets = new XSSFWorkbook(fileInputStream);
            //获取sheet
            sheet = sheets.getSheet(sheetName);
        } catch (Exception e) {
            System.out.println(filePath+"失败");
        }
    }

    /**
     * 根据行和列的索引获取单元格的数据
     * @param row
     * @param column
     * @return
     */
    public String getExcelDateByIndex(int row,int column){
        XSSFRow row1 = sheet.getRow(row);
        String cell = row1.getCell(column).toString();
        return cell;
    }

    public String getExcelNumberByIndex(int row,int column){
        XSSFRow row1 = sheet.getRow(row);
        String cell = NumberToTextConverter.toText(row1.getCell(column).getNumericCellValue());
        return cell;
    }



    /**
     * 根据某一列值为“******”的这一行，来获取该行第x列的值
     * @param caseName
     * @param currentColumn 当前单元格列的索引
     * @param targetColumn 目标单元格列的索引
     * @return
     */
    public String getCellByCaseName(String caseName,int currentColumn,int targetColumn){
        String operateSteps="";
        //获取行数
        int rows = sheet.getPhysicalNumberOfRows();
        for(int i=0;i<rows;i++){
            XSSFRow row = sheet.getRow(i);
            String cell = row.getCell(currentColumn).toString();
            if(cell.equals(caseName)){
                operateSteps = row.getCell(targetColumn).toString();
                break;
            }
        }
        return operateSteps;
    }


    public int getNumberOfRows(){
        return sheet.getPhysicalNumberOfRows();
    }

    //测试方法
    public static void main(String[] args){
        ExcelData sheet1 = new ExcelData("resource/FirstTests.xlsx", "username");
        //获取第二行第4列
        String cell2 = sheet1.getExcelDateByIndex(1, 3);
        //根据第3列值为“customer23”的这一行，来获取该行第2列的值
        String cell3 = sheet1.getCellByCaseName("customer23", 2,1);
        System.out.println(cell2);
        System.out.println(cell3);
    }
}