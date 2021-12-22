package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;
import org.imc.service.nancy.excel.CreateExcelFile;
import org.imc.service.nancy.excel.ExcelData;
import org.imc.service.nancy.model.ChapterModel;
import org.imc.service.nancy.model.EmployeeModel;
import org.imc.service.nancy.model.NovalModel;
import org.imc.tools.CommonTool;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

@Component
@Slf4j
public class TranslatorNumGenerateService {

    private List<String> files = new LinkedList<>();
    private List<String> errorOpenFiles = new LinkedList<>();
    private Map<String,Map<String,List<String>>> translatorModelMap = new HashMap<>();
    private Map<String,Map<String,List<String>>> editorModelMap = new HashMap<>();
    private Map<String,Map<String,List<String>>> qualityModelMap = new HashMap<>();
    private Map<String,Map<String,String>> novelTranslatorNameLevelMap = new HashMap<>();

    public void generate(String path) {
        // 1.记录和移除隐藏文件
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".xlsx");
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);

        // 2.业务解析
        File outputDirectory = new File("译员编号信息");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        for(String file:files){
            ExcelData sheet = new ExcelData(file, "译员资料");
            Map<String,String> translatorNameLevelMap = new HashMap<>();
            try{
                buildTranslatorNameLevelMap(sheet,translatorNameLevelMap);
                buildInternalDataStruct(file,translatorNameLevelMap,novelTranslatorNameLevelMap);
            }catch (Exception e){
                errorOpenFiles.add(file);
            }
        }
        // 文件挂统计
        for(String errorOpenFile:errorOpenFiles){
            log.error("文件挂了,文件名"+errorOpenFile);
        }
        if(errorOpenFiles.size()>0){
            log.error("有文件挂了,请检查");
            CommonTool.enterKeyContinue("按回车结束程序");
            return;
        }
        buildExcel();
        CommonTool.enterKeyContinue("解析结束，按回车键继续");
    }

    private void buildExcel() {
        for(Map.Entry<String,Map<String,List<String>>> entry:translatorModelMap.entrySet()){
            String novelName = entry.getKey();
            Map<String,List<String>> translatorChapterMap = translatorModelMap.get(novelName);
            Map<String,List<String>> editorChapterMap = editorModelMap.get(novelName);
            Map<String,List<String>> qualityChapterMap = qualityModelMap.get(novelName);
            Map<String,String> translatorNameLevelMap = novelTranslatorNameLevelMap.get(novelName);
            String fileDir = "译员编号信息\\"+novelName+"-"+"译员编号信息"+".xlsx";
            CreateExcelFile.createExcelXlsx(fileDir,"译员编号信息", new String[]{"书名",novelName});
            File file = new File(fileDir);
            //创建workbook
            FileOutputStream fileOutputStream = null;

            try {
                XSSFWorkbook xWorkbook = new XSSFWorkbook(new FileInputStream(file));
                //文件流
                XSSFSheet sheet = xWorkbook.getSheet("译员编号信息");
                // 获取表格的总行数
                int rowCount = sheet.getLastRowNum();
                //获取表头的列数 int columnCount = sheet.getRow(0).getLastCellNum();
                // 黄色底色
                XSSFCellStyle cellStyle = (XSSFCellStyle) xWorkbook.createCellStyle();
                cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 255, 0)));
                cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                rowCount = buildRoleExcel(translatorChapterMap, translatorNameLevelMap, sheet, rowCount, cellStyle);
                rowCount = buildRoleExcel(editorChapterMap, translatorNameLevelMap, sheet, rowCount, cellStyle);
                rowCount = buildRoleExcel(qualityChapterMap, translatorNameLevelMap, sheet, rowCount, cellStyle);
                fileOutputStream = new FileOutputStream(fileDir);
                xWorkbook.write(fileOutputStream);
            }catch (IOException e) {
                e.printStackTrace();
            }finally {
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

    private int buildRoleExcel(Map<String, List<String>> roleChapterMap, Map<String, String> translatorNameLevelMap, XSSFSheet sheet, int rowCount, XSSFCellStyle cellStyle) {
        for(Map.Entry<String,List<String>> entry : roleChapterMap.entrySet()) {
            String roleName = entry.getKey();
            String roleLevel = translatorNameLevelMap.get(roleName);
            XSSFRow titleRow = sheet.createRow(++rowCount);
            XSSFCell titleNameCell = titleRow.createCell(0);
            XSSFCell titleLevelCell = titleRow.createCell(1);
            titleNameCell.setCellValue("译员姓名");
            titleLevelCell.setCellValue("译员等级");
            XSSFRow roleRow = sheet.createRow(++rowCount);
            XSSFCell roleNameCell = roleRow.createCell(0);
            XSSFCell roleLevelCell = roleRow.createCell(1);
            roleNameCell.setCellValue(roleName);
            roleNameCell.setCellStyle(cellStyle);
            roleLevelCell.setCellValue(roleLevel);
            roleLevelCell.setCellStyle(cellStyle);
            List<String> chapterNames = entry.getValue();
            for(String chapterName:chapterNames){
                XSSFRow chapterNameRow = sheet.createRow(++rowCount);
                XSSFCell nullCell1 = chapterNameRow.createCell(0);
                XSSFCell nullCell2 = chapterNameRow.createCell(1);
                XSSFCell chapterNameCell = chapterNameRow.createCell(2);
                nullCell1.setCellValue("");
                nullCell2.setCellValue("");
                chapterNameCell.setCellValue(chapterName);
            }
        }
        return rowCount;
    }

    private void buildTranslatorNameLevelMap(ExcelData sheet,Map<String,String> outMap) throws Exception{
        if(!"译员姓名".equals(sheet.getExcelDateByIndex(0,0))||!"译员等级".equals(sheet.getExcelDateByIndex(0,3))){
            throw new Exception();
        }
        int rows = sheet.getNumberOfRows();
        for(int i = 1;i<rows;i++){
            String name = sheet.getExcelDateByIndex(i,0);
            String leve = sheet.getExcelDateByIndex(i,3);
            if(!isValidName(name)){
                continue;
            }
            if(outMap.containsKey(name)){
                log.error("sheet译员资料，译员姓名重复，姓名: "+name," 行："+i+1);
                throw new Exception();
            }
            outMap.put(name,leve);
        }
    }

    private void buildInternalDataStruct(String file,Map<String,String> translatorNameLevelMap,Map<String,Map<String,String>> novelTranslatorNameLevelMap) throws Exception{
        ExcelData sheet = new ExcelData(file, "W2进度");
        if(!"序号".equals(sheet.getExcelDateByIndex(0,0))||!"小说名称全称".equals(sheet.getExcelDateByIndex(0,1))){
            throw new Exception();
        }
        int rows = sheet.getNumberOfRows();
        String novelName = sheet.getExcelDateByIndex(1,1);
        novelTranslatorNameLevelMap.put(novelName,translatorNameLevelMap);
        if(translatorModelMap.containsKey(novelName)){
            log.error("小说重复，重复的小说:"+file);
            throw new Exception();
        }
        translatorModelMap.put(novelName,new HashMap<>());
        editorModelMap.put(novelName,new HashMap<>());
        qualityModelMap.put(novelName,new HashMap<>());
        Map<String,List<String>> translatorChapterMap = translatorModelMap.get(novelName);
        Map<String,List<String>> editorChapterMap = editorModelMap.get(novelName);
        Map<String,List<String>> qualityChapterMap = qualityModelMap.get(novelName);
        for(int i = 1;i<rows;i++){
            String chapterName = sheet.getExcelDateByIndex(i,2);
            String translatorName = sheet.getExcelDateByIndex(i,11);
            String editorName = sheet.getExcelDateByIndex(i,20);
            String qualityName = sheet.getExcelDateByIndex(i,25);
            String state = sheet.getExcelDateByIndex(i,9);
            // 后续没有内容了
            if(!isValidName(translatorName)&&!isValidName(editorName)&&!isValidName(qualityName)&&"".equals(chapterName)){
                break;
            }
            if(!"待提交".equals(state)){
                continue;
            }
            buildRole(translatorChapterMap, chapterName, translatorName);
            buildRole(editorChapterMap, chapterName, editorName);
            buildRole(qualityChapterMap, chapterName, qualityName);
        }
    }

    private void buildRole(Map<String, List<String>> roleChapterMap, String chapterName, String roleName) {
        if(isValidName(roleName)){
            if(!roleChapterMap.containsKey(roleName)){
                roleChapterMap.put(roleName,new LinkedList<>());
            }
            roleChapterMap.get(roleName).add(chapterName);
        }
    }

    private boolean isValidName(String roleName) {
        return !"0.0".equals(roleName)&&!"".equals(roleName)&&!"无".equals(roleName);
    }
}
