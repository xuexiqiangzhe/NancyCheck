package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.maven.surefire.shade.org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.*;
import org.imc.service.nancy.excel.CreateExcelFile;
import org.imc.service.nancy.excel.ExcelData;
import org.imc.service.nancy.model.ChapterModel;
import org.imc.service.nancy.model.EmployeeModel;
import org.imc.service.nancy.model.NovalModel;
import org.imc.service.nancy.model.TranslatorModel;
import org.imc.tools.CommonTool;
import org.imc.tools.FileExportUtil;
import org.imc.tools.NumberTool;
import org.springframework.stereotype.Component;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;

@Component
@Slf4j
public class PaymentGenerateService {
    //1 小说名称   2 章节名称  7.字数  8.章节数  9.当前状态  19.翻译人员  20翻译人员单价  22 翻译成本
    // 26 母语翻译 27 母语编辑单价  29母语编辑成本
    int[] col = {1,2,7,8,9,19,20,22,26,27,29};
    private List<String> files = new LinkedList<>();
    private List<String> errorOpenFiles = new LinkedList<>();
    private TranslatorModel translatorModel = new TranslatorModel();
    private Map<String,List<Pair<String,String>>> filePeopleAccountZeroMap = new HashMap<>();

    public void generate(String path) {
        Map<String, EmployeeModel> translatorModelMap =  translatorModel.getTranslatorModelMap();
        Map<String, EmployeeModel> editorModelMap =   translatorModel.getEditorModelMap();
        Map<String, EmployeeModel> qualityModelMap =                 translatorModel.getQualityModelMap();

        // 1.记录和移除隐藏文件
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".xlsx");
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);

        // 2.业务解析
        File outputDirectory = new File("译费结算结果");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        for(String file:files){
            ExcelData sheet = new ExcelData(file, "W2成本副本");
            try{
                int rows = sheet.getNumberOfRows();
                // 2.1.翻译
                buildTranslator(file,translatorModelMap, sheet, rows);
                // 2.2.编辑
                buildEditor(file,editorModelMap,sheet,rows);
                // 2.3.审校
                buildQuality(file,qualityModelMap,sheet,rows);
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

        // 算钱
        calculatorSum(translatorModelMap);
        calculatorSum(editorModelMap);
        calculatorSum(qualityModelMap);

        // 输出excels
        buildExcels(translatorModelMap,editorModelMap,qualityModelMap);
        buildExcels2(editorModelMap,qualityModelMap);
        buildExcels3(qualityModelMap);

        // 输出成本太低txt
        String errorAccount = "";
        for (Map.Entry<String,List<Pair<String,String>>> entry:filePeopleAccountZeroMap.entrySet()){
            String file = entry.getKey();
            for(Pair<String,String> pair:entry.getValue()){
                String content ="成本疑似异常，文件："+file+", "+"人："+pair.getLeft()+", "+"章节："+pair.getRight()+"\n";
                errorAccount+=content;
            }
        }
        File errorAccountOutFile = new File("成本疑似位置.txt");
        FileExportUtil.buildNormalOutPutFile(errorAccountOutFile, errorAccount);
        CommonTool.enterKeyContinue("解析结束，按回车键继续");
    }


    private void calculatorSum(Map<String, EmployeeModel> modelMap){
        for(Map.Entry<String, EmployeeModel> entry: modelMap.entrySet()) {
            String name = entry.getKey();
            EmployeeModel employeeModel = entry.getValue();
            for (Map.Entry<String, NovalModel> subEntry:employeeModel.getNovalModelMap().entrySet()) {
                NovalModel novalModel = subEntry.getValue();

                employeeModel.setTotalAccount(addTwoDouble(employeeModel.getTotalAccount(),novalModel.getAccount()));
            }
        }
    }
    private void buildExcels(Map<String, EmployeeModel> translatorModelMap,Map<String, EmployeeModel> editorMap,Map<String, EmployeeModel> qualityMap) {
        for(Map.Entry<String, EmployeeModel> entry: translatorModelMap.entrySet()){
            String name = entry.getKey();
            EmployeeModel employeeModel = entry.getValue();
            buildExcelForEmployee(name, employeeModel);
            if(editorMap.containsKey(name)){
                buildExcelForEmployee(name,editorMap.get(name));
                editorMap.remove(name);
            }
            if(qualityMap.containsKey(name)){
                buildExcelForEmployee(name,qualityMap.get(name));
                qualityMap.remove(name);
            }
        }
    }

    private void buildExcels2(Map<String, EmployeeModel> editorMap,Map<String, EmployeeModel> qualityMap) {
        for(Map.Entry<String, EmployeeModel> entry: editorMap.entrySet()){
            String name = entry.getKey();
            EmployeeModel employeeModel = entry.getValue();
            buildExcelForEmployee(name, employeeModel);
            if(qualityMap.containsKey(name)){
                buildExcelForEmployee(name,qualityMap.get(name));
                qualityMap.remove(name);
            }
        }
    }

    private void buildExcels3(Map<String, EmployeeModel> qualityMap) {
        for(Map.Entry<String, EmployeeModel> entry: qualityMap.entrySet()){
            String name = entry.getKey();
            EmployeeModel employeeModel = entry.getValue();
            buildExcelForEmployee(name, employeeModel);
        }
    }

    private void buildExcelForEmployee(String name, EmployeeModel employeeModel) {
        String fileDir = "译费结算结果\\"+name +".xlsx";
        File file = new File(fileDir);
        Boolean isNewFile = false;
        if (!file.exists()){
            CreateExcelFile.createExcelXlsx(fileDir,"结算", new String[]{"姓名", "总译费"});
            file = new File(fileDir);
            isNewFile=true;
        }

        //创建workbook
        FileOutputStream fileOutputStream = null;

        try {
            XSSFWorkbook xWorkbook = new XSSFWorkbook(new FileInputStream(file));
            //文件流
            XSSFSheet sheet = xWorkbook.getSheet("结算");
            // 获取表格的总行数
            int rowCount = sheet.getLastRowNum();
            //获取表头的列数
            int columnCount = sheet.getRow(0).getLastCellNum();
            // 黄色底色
            XSSFCellStyle cellStyle = (XSSFCellStyle) xWorkbook.createCellStyle();
            cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 255, 0)));
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

            // 翻译人
            if(isNewFile){
                XSSFRow newRow=sheet.createRow(++rowCount);
                XSSFCell nameCell = newRow.createCell(0);
                XSSFCell totalAccountCell = newRow.createCell(1);
                nameCell.setCellValue(name);
                nameCell.setCellStyle(cellStyle);
                totalAccountCell.setCellValue(employeeModel.getTotalAccount());
                totalAccountCell.setCellStyle(cellStyle);
            }

            for (Map.Entry<String, NovalModel> subEntry: employeeModel.getNovalModelMap().entrySet()) {
                String novalName = subEntry.getKey();
                NovalModel novalModel = subEntry.getValue();
                XSSFRow novalRow =sheet.createRow(++rowCount);
                XSSFCell novalNameCell = novalRow.createCell(0);
                XSSFCell novalChapterCountCell = novalRow.createCell(1);
                XSSFCell novalWordCountCell = novalRow.createCell(2);
                XSSFCell novalAccountCell = novalRow.createCell(3);
                novalNameCell.setCellValue("小说");
                novalChapterCountCell.setCellValue("章节数");
                novalWordCountCell.setCellValue("小说字数");
                novalAccountCell.setCellValue("小说总译费");
                XSSFRow novalRow1 =sheet.createRow(++rowCount);
                XSSFCell novalNameCell1 = novalRow1.createCell(0);
                XSSFCell novalChapterCountCell1 = novalRow1.createCell(1);
                XSSFCell novalWordCountCell1 = novalRow1.createCell(2);
                XSSFCell novalAccountCell1 = novalRow1.createCell(3);
                novalNameCell1.setCellValue(novalName);
                novalNameCell1.setCellStyle(cellStyle);
                novalChapterCountCell1.setCellValue(novalModel.getChapterCount());
                novalChapterCountCell1.setCellStyle(cellStyle);
                novalWordCountCell1.setCellValue(novalModel.getWordCount());
                novalWordCountCell1.setCellStyle(cellStyle);
                novalAccountCell1.setCellValue(novalModel.getAccount());
                novalAccountCell1.setCellStyle(cellStyle);
                XSSFRow novalRow2 =sheet.createRow(++rowCount);
                XSSFCell chapterNameCell = novalRow2.createCell(0);
                XSSFCell chapterCountCell = novalRow2.createCell(1);
                XSSFCell wordCountCell = novalRow2.createCell(2);
                XSSFCell accountCell = novalRow2.createCell(3);
                XSSFCell singlePriceCell = novalRow2.createCell(4);
                chapterNameCell.setCellValue("章节名称");
                chapterCountCell.setCellValue("章节数");
                wordCountCell.setCellValue("章节字数");
                accountCell.setCellValue("译费");
                singlePriceCell.setCellValue("单价");
                for(Map.Entry<String,ChapterModel> chapterModelEntry:novalModel.getChapterModelMap().entrySet()){
                    String chapterName = chapterModelEntry.getKey();
                    ChapterModel chapterModel = chapterModelEntry.getValue();
                    XSSFRow novalRow3 =sheet.createRow(++rowCount);
                    XSSFCell chapterNameCell1 = novalRow3.createCell(0);
                    XSSFCell chapterCountCell1 = novalRow3.createCell(1);
                    XSSFCell wordCountCell1 = novalRow3.createCell(2);
                    XSSFCell accountCell1 = novalRow3.createCell(3);
                    XSSFCell singlePriceCell1 = novalRow3.createCell(4);
                    chapterNameCell1.setCellValue(chapterName);
                    chapterCountCell1.setCellValue(chapterModel.getChapterDetailMap().get("chapterCount"));
                    wordCountCell1.setCellValue(chapterModel.getChapterDetailMap().get("wordCount"));
                    accountCell1.setCellValue(chapterModel.getChapterDetailMap().get("account"));
                    singlePriceCell1.setCellValue(chapterModel.getChapterDetailMap().get("singlePrice"));
                }
            }

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

    private Double addTwoDouble(Double d1,Double d2){
        BigDecimal b1=new BigDecimal(Double.toString(d1));
        BigDecimal b2=new BigDecimal(Double.toString(d2));
        return b1.add(b2).doubleValue();
    }

    private void buildTranslator(String file,Map<String, EmployeeModel> translatorModelMap, ExcelData sheet, int rows) throws Exception {
        if(!"序号".equals(sheet.getExcelDateByIndex(0, 0))){
            throw new Exception();
        }
        for(int x = 1; x< rows; x++){
            String currentState = sheet.getExcelDateByIndex(x, 9);
            if(!"待结算".equals(currentState)){
                continue;
            }

            // 当前人员
            int y = 19;
            String name = sheet.getExcelDateByIndex(x, y);
            if("无".equals(name)||"".equals(name)||"0.0".equals(name)){
                continue;
            }

            if(!translatorModelMap.containsKey(name)){
                translatorModelMap.put(name,new EmployeeModel());
            }
            Map<String, NovalModel> novalModelMap = translatorModelMap.get(name).getNovalModelMap();

            // 当前小说
            y = 1;
            String novalName = sheet.getExcelDateByIndex(x, y);
            if(!novalModelMap.containsKey(novalName)){
                novalModelMap.put(novalName,new NovalModel());
            }
            Map<String, ChapterModel> novalDetail = novalModelMap.get(novalName).getChapterModelMap();
            // 当前章
            String chapterName = sheet.getExcelDateByIndex(x, 2);
            novalDetail.put(chapterName,new ChapterModel());
            Map<String,String> chapterDetail  = novalDetail.get(chapterName).getChapterDetailMap();

            // 一章四个属性
            String wordCount = sheet.getExcelNumberByIndex(x, 7);
            chapterDetail.put("wordCount",wordCount);
            novalModelMap.get(novalName).setWordCount(addTwoDouble(novalModelMap.get(novalName).getWordCount(),Double.parseDouble(wordCount)));
            String chapterCount = sheet.getExcelNumberByIndex(x, 8);
            chapterDetail.put("chapterCount",chapterCount);
            novalModelMap.get(novalName).setChapterCount(addTwoDouble(novalModelMap.get(novalName).getChapterCount(),Double.parseDouble(chapterCount)));
            String singlePrice = sheet.getExcelNumberByIndex(x, 20);
            chapterDetail.put("singlePrice",singlePrice);
            String account = sheet.getExcelNumberByIndex(x, 22);
            chapterDetail.put("account",account);
            recordAccountTooLow(file, name, chapterName, account);
            novalModelMap.get(novalName).setAccount(addTwoDouble(novalModelMap.get(novalName).getAccount(),Double.parseDouble(account)));
        }
    }

    // 26 母语翻译 27 母语编辑单价  29母语编辑成本
    private void buildEditor(String file,Map<String, EmployeeModel> editorMap, ExcelData sheet, int rows) throws Exception {
        if(!"序号".equals(sheet.getExcelDateByIndex(0, 0))){
            throw new Exception();
        }
        for(int x = 1; x< rows; x++){
            String currentState = sheet.getExcelDateByIndex(x, 9);
            if(!"待结算".equals(currentState)){
                continue;
            }

            // 当前人员
            int y = 26;
            String name = sheet.getExcelDateByIndex(x, y);
            if("无".equals(name)||"".equals(name)||"0.0".equals(name)){
                continue;
            }
            if(!editorMap.containsKey(name)){
                editorMap.put(name,new EmployeeModel());
            }
            Map<String, NovalModel> novalModelMap = editorMap.get(name).getNovalModelMap();

            // 当前小说
            y = 1;
            String novalName = sheet.getExcelDateByIndex(x, y);
            if(!novalModelMap.containsKey(novalName)){
                novalModelMap.put(novalName,new NovalModel());
            }
            Map<String, ChapterModel> novalDetail = novalModelMap.get(novalName).getChapterModelMap();
            // 当前章
            String chapterName = sheet.getExcelDateByIndex(x, 2);
            novalDetail.put(chapterName,new ChapterModel());
            Map<String,String> chapterDetail  = novalDetail.get(chapterName).getChapterDetailMap();

            // 一章四个属性
            String wordCount = sheet.getExcelNumberByIndex(x, 7);
            chapterDetail.put("wordCount",wordCount);
            novalModelMap.get(novalName).setWordCount(addTwoDouble(novalModelMap.get(novalName).getWordCount(),Double.parseDouble(wordCount)));
            String chapterCount = sheet.getExcelNumberByIndex(x, 8);
            chapterDetail.put("chapterCount",chapterCount);
            novalModelMap.get(novalName).setChapterCount(addTwoDouble(novalModelMap.get(novalName).getChapterCount(),Double.parseDouble(chapterCount)));
            String singlePrice = sheet.getExcelNumberByIndex(x, 27);
            chapterDetail.put("singlePrice",singlePrice);
            String account = sheet.getExcelNumberByIndex(x, 29);
            chapterDetail.put("account",account);
            recordAccountTooLow(file, name, chapterName, account);
            novalModelMap.get(novalName).setAccount(addTwoDouble(novalModelMap.get(novalName).getAccount(),Double.parseDouble(account)));
        }
    }

    // 33 质检姓名 34 单价  36 成本
    private void buildQuality(String file,Map<String, EmployeeModel> qualityMap, ExcelData sheet, int rows) throws Exception {
        if(!"序号".equals(sheet.getExcelDateByIndex(0, 0))){
            throw new Exception();
        }
        for(int x = 1; x< rows; x++){
            String currentState = sheet.getExcelDateByIndex(x, 9);
            if(!"待结算".equals(currentState)){
                continue;
            }

            // 当前人员
            int y = 33;
            String name = sheet.getExcelDateByIndex(x, y);
            if("无".equals(name)||"".equals(name)||"0.0".equals(name)){
                continue;
            }

            if(!qualityMap.containsKey(name)){
                qualityMap.put(name,new EmployeeModel());
            }
            Map<String, NovalModel> novalModelMap = qualityMap.get(name).getNovalModelMap();

            // 当前小说
            y = 1;
            String novalName = sheet.getExcelDateByIndex(x, y);
            if(!novalModelMap.containsKey(novalName)){
                novalModelMap.put(novalName,new NovalModel());
            }
            Map<String, ChapterModel> novalDetail = novalModelMap.get(novalName).getChapterModelMap();
            // 当前章
            String chapterName = sheet.getExcelDateByIndex(x, 2);
            novalDetail.put(chapterName,new ChapterModel());
            Map<String,String> chapterDetail  = novalDetail.get(chapterName).getChapterDetailMap();

            // 一章四个属性
            String wordCount = sheet.getExcelNumberByIndex(x, 7);
            chapterDetail.put("wordCount",wordCount);
            novalModelMap.get(novalName).setWordCount(addTwoDouble(novalModelMap.get(novalName).getWordCount(),Double.parseDouble(wordCount)));
            String chapterCount = sheet.getExcelNumberByIndex(x, 8);
            chapterDetail.put("chapterCount",chapterCount);
            novalModelMap.get(novalName).setChapterCount(addTwoDouble(novalModelMap.get(novalName).getChapterCount(),Double.parseDouble(chapterCount)));
            String singlePrice = sheet.getExcelNumberByIndex(x, 34);
            chapterDetail.put("singlePrice",singlePrice);
            String account = sheet.getExcelNumberByIndex(x, 36);
            chapterDetail.put("account",account);
            recordAccountTooLow(file, name, chapterName, account);
            novalModelMap.get(novalName).setAccount(addTwoDouble(novalModelMap.get(novalName).getAccount(),Double.parseDouble(account)));
        }
    }

    private void recordAccountTooLow(String file, String name, String chapterName, String account) {
        if(Double.parseDouble(account)<0.55){
            Pair<String,String> peopleChapter = Pair.of(name, chapterName);
            if(!filePeopleAccountZeroMap.containsKey(file)){
                filePeopleAccountZeroMap.put(file,new LinkedList<>());
            }
            filePeopleAccountZeroMap.get(file).add(peopleChapter);
        }
    }
}
