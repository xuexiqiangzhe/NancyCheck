package org.imc.service.nancy.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;


/**
 * @author bpan
 *
 * created 2018年2月27日
 */
@Component
public class CreateExcelFile {

    private static XSSFWorkbook xWorkbook = null;


    /**
     * 创建Excel(xlsx)
     * @param fileDir  文件名称及地址
     * @param sheetName sheet的名称
     * @param titleRow  表头
     */
    public static void createExcelXlsx(String fileDir, String sheetName, String titleRow[]){
        //创建workbook
        xWorkbook = new XSSFWorkbook();
        //新建文件
        FileOutputStream fileOutputStream = null;
        XSSFRow row = null;
        try {

//            CellStyle cellStyle = xWorkbook.createCellStyle();
//            cellStyle.setAlignment(HorizontalAlignment.LEFT);
//            cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);

            //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)

            xWorkbook.createSheet(sheetName);
            xWorkbook.getSheet(sheetName).createRow(0);
            //添加表头, 创建第一行
            row = xWorkbook.getSheet(sheetName).createRow(0);
            row.setHeight((short)(20*20));
            for (short j = 0; j < titleRow.length; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(titleRow[j]);
//                cell.setCellStyle(cellStyle);
            }
            fileOutputStream = new FileOutputStream(fileDir);
            xWorkbook.write(fileOutputStream);
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }finally {

            if (fileOutputStream != null) {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }
    }

    public static void writeXlsx(String fileDir,String sheetName,List<Map<String,String>> mapList){
        //创建workbook
        File file = new File(fileDir);

        try {
            xWorkbook = new XSSFWorkbook(new FileInputStream(file));
        }catch(FileNotFoundException e){
            e.printStackTrace();
        }catch (IOException e) {
            e.printStackTrace();
        }

        //文件流
        FileOutputStream fileOutputStream = null;
        XSSFSheet sheet = xWorkbook.getSheet(sheetName);
        // 获取表格的总行数
         int rowCount = sheet.getLastRowNum();
        //获取表头的列数
        int columnCount = sheet.getRow(0).getLastCellNum();

        try {
            // 获得表头行对象
            XSSFRow titleRow = sheet.getRow(0);
            //创建单元格显示样式

            if(titleRow!=null){
                for(int rowId = rowCount; rowId < mapList.size(); rowId++){
                    Map<String,String> map = mapList.get(rowId);
                    XSSFRow newRow=sheet.createRow(rowId+1);
                    newRow.setHeight((short)(20*20));//设置行高  基数为20

                    for (short columnIndex = 0; columnIndex < columnCount; columnIndex++) {  //遍历表头
                        //trim()的方法是删除字符串中首尾的空格
                        String mapKey = titleRow.getCell(columnIndex).toString().trim();
                        XSSFCell cell = newRow.createCell(columnIndex);
                        cell.setCellValue(map.get(mapKey)==null ? null : map.get(mapKey).toString());
                    }
                }
            }

            fileOutputStream = new FileOutputStream(fileDir);
            xWorkbook.write(fileOutputStream);
        } catch (Exception e) {
           e.printStackTrace();
        } finally {
            try {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
