package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.imc.service.nancy.excel.CreateExcelFile;
import org.imc.service.nancy.excel.ExcelData;
import org.imc.service.nancy.model.ChapterModel;
import org.imc.service.nancy.model.EmployeeModel;
import org.imc.service.nancy.model.NovalModel;
import org.imc.service.nancy.model.TranslatorModel;
import org.imc.tools.NumberTool;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

@Component
@Slf4j
public class PaymentGenerateService {
    //1 小说名称   2 章节名称  7.字数  8.章节数  9.当前状态  19.翻译人员  20翻译人员单价  22 翻译成本
    // 26 母语翻译 27 母语编辑单价  29母语编辑成本
    int[] col = {1,2,7,8,9,19,20,22,26,27,29};
    private List<String> files = new LinkedList<>();
    private TranslatorModel translatorModel = new TranslatorModel();


    public void generate(String path) {
        Map<String, EmployeeModel> translatorModelMap =  translatorModel.getTranslatorModelMap();
        Map<String, EmployeeModel> editorModelMap =   translatorModel.getEditorModelMap();
        Map<String, EmployeeModel> qualityModelMap =                 translatorModel.getQualityModelMap();
                log.info("开始记录文件");
        recordFile(path);
        for(String file:files){
            log.info("已记录文件:"+file);
        }

        for(String file:files){
            ExcelData sheet = new ExcelData(file, "W2成本副本");
            // 1.翻译
            int rows = sheet.getNumberOfRows();
            buildTranslator(translatorModelMap, sheet, rows);
            // 2.编辑
            buildEditor(editorModelMap,sheet,rows);
            // 3.审校
            buildQuality(qualityModelMap,sheet,rows);

            calculatorSum(translatorModelMap);
            calculatorSum(editorModelMap);
            calculatorSum(qualityModelMap);

            buildExcels(translatorModelMap,editorModelMap,qualityModelMap);
            buildExcels2(editorModelMap,qualityModelMap);
            buildExcels3(qualityModelMap);
            String cell2 = sheet.getExcelDateByIndex(1, 3);
//        //根据第3列值为“customer23”的这一行，来获取该行第2列的值
//        String cell3 = sheet1.getCellByCaseName("customer23", 2,1);
//        System.out.println(cell2);
//        System.out.println(cell3);
        }

    }

    private void calculatorSum(Map<String, EmployeeModel> modelMap){
        for(Map.Entry<String, EmployeeModel> entry: modelMap.entrySet()) {
            String name = entry.getKey();
            EmployeeModel employeeModel = entry.getValue();
            for (Map.Entry<String, NovalModel> subEntry:employeeModel.getNovalModelMap().entrySet()) {
                NovalModel novalModel = subEntry.getValue();
                employeeModel.setTotalAccount(employeeModel.getTotalAccount()+novalModel.getAccount());
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
        String fileDir = name +".xlsx";
        File file = new File(fileDir);

        if (!file.exists()){
            CreateExcelFile.createExcelXlsx(fileDir,"结算", new String[]{"姓名", "总译费"});
            file = new File(fileDir);
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

            // 翻译人
            XSSFRow newRow=sheet.createRow(++rowCount);
            XSSFCell nameCell = newRow.createCell(0);
            XSSFCell totalAccountCell = newRow.createCell(1);
            nameCell.setCellValue(name);
            totalAccountCell.setCellValue(employeeModel.getTotalAccount());

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
                novalChapterCountCell1.setCellValue(novalModel.getChapterCount());
                novalWordCountCell1.setCellValue(novalModel.getWordCount());
                novalAccountCell1.setCellValue(novalModel.getAccount());
                XSSFRow novalRow2 =sheet.createRow(++rowCount);
                XSSFCell chapterNameCell = novalRow2.createCell(0);
                XSSFCell chapterCountCell = novalRow2.createCell(1);
                XSSFCell wordCountCell = novalRow2.createCell(2);
                XSSFCell accountCell = novalRow2.createCell(3);
                XSSFCell singlePriceCell = novalRow2.createCell(4);
                chapterNameCell.setCellValue("章节名称");
                chapterCountCell.setCellValue("章节数");
                wordCountCell.setCellValue("章节字数");
                accountCell.setCellValue("成本");
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

    private void buildTranslator(Map<String, EmployeeModel> translatorModelMap, ExcelData sheet, int rows) {
        for(int x = 1; x< rows; x++){
            String currentState = sheet.getExcelDateByIndex(x, 9);
            if(!"待结算".equals(currentState)){
                continue;
            }

            // 当前人员
            int y = 19;
            String name = sheet.getExcelDateByIndex(x, y);
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
            String wordCount = sheet.getExcelDateByIndex(x, 7);
            chapterDetail.put("wordCount",wordCount);
            novalModelMap.get(novalName).setWordCount(novalModelMap.get(novalName).getWordCount()+Double.parseDouble(wordCount));
            String chapterCount = sheet.getExcelDateByIndex(x, 8);
            chapterDetail.put("chapterCount",chapterCount);
            novalModelMap.get(novalName).setChapterCount(novalModelMap.get(novalName).getChapterCount()+Double.parseDouble(chapterCount));
            String singlePrice = sheet.getExcelDateByIndex(x, 20);
            chapterDetail.put("singlePrice",singlePrice);
            String account = sheet.getExcelDateByIndex(x, 22);
            chapterDetail.put("account",account);
            novalModelMap.get(novalName).setAccount(novalModelMap.get(novalName).getAccount()+Double.parseDouble(account));
        }
    }

    // 26 母语翻译 27 母语编辑单价  29母语编辑成本
    private void buildEditor(Map<String, EmployeeModel> editorMap, ExcelData sheet, int rows) {
        for(int x = 1; x< rows; x++){
            String currentState = sheet.getExcelDateByIndex(x, 9);
            if(!"待结算".equals(currentState)){
                continue;
            }

            // 当前人员
            int y = 26;
            String name = sheet.getExcelDateByIndex(x, y);
            if("无".equals(name)){
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
            String wordCount = sheet.getExcelDateByIndex(x, 7);
            chapterDetail.put("wordCount",wordCount);
            novalModelMap.get(novalName).setWordCount(novalModelMap.get(novalName).getWordCount()+Double.parseDouble(wordCount));
            String chapterCount = sheet.getExcelDateByIndex(x, 8);
            chapterDetail.put("chapterCount",chapterCount);
            novalModelMap.get(novalName).setChapterCount(novalModelMap.get(novalName).getChapterCount()+Double.parseDouble(chapterCount));
            String singlePrice = sheet.getExcelDateByIndex(x, 27);
            chapterDetail.put("singlePrice",singlePrice);
            String account = sheet.getExcelDateByIndex(x, 29);
            chapterDetail.put("account",account);
            novalModelMap.get(novalName).setAccount(novalModelMap.get(novalName).getAccount()+Double.parseDouble(account));
        }
    }

    // 33 质检姓名 34 单价  36 成本
    private void buildQuality(Map<String, EmployeeModel> qualityMap, ExcelData sheet, int rows) {
        for(int x = 1; x< rows; x++){
            String currentState = sheet.getExcelDateByIndex(x, 9);
            if(!"待结算".equals(currentState)){
                continue;
            }

            // 当前人员
            int y = 33;
            String name = sheet.getExcelDateByIndex(x, y);
            if("无".equals(name)){
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
            String wordCount = sheet.getExcelDateByIndex(x, 7);
            chapterDetail.put("wordCount",wordCount);
            novalModelMap.get(novalName).setWordCount(novalModelMap.get(novalName).getWordCount()+Double.parseDouble(wordCount));
            String chapterCount = sheet.getExcelDateByIndex(x, 8);
            chapterDetail.put("chapterCount",chapterCount);
            novalModelMap.get(novalName).setChapterCount(novalModelMap.get(novalName).getChapterCount()+Double.parseDouble(chapterCount));
            String singlePrice = sheet.getExcelDateByIndex(x, 34);
            chapterDetail.put("singlePrice",singlePrice);
            String account = sheet.getExcelDateByIndex(x, 36);
            chapterDetail.put("account",account);
            novalModelMap.get(novalName).setAccount(novalModelMap.get(novalName).getAccount()+Double.parseDouble(account));
        }
    }

    private void buildOutPut(File file, String content) throws IOException {
        // write
        FileWriter fw = new FileWriter(file, true);
        BufferedWriter bw = new BufferedWriter(fw);
        bw.write(content);
        bw.flush();
        bw.close();
        fw.close();
    }



    private void recordFile(String path){
        File file = new File(path);
        File[] tempList = file.listFiles();
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile()) {
                String fileName = tempList[i].toString();
                Integer end = fileName.length();
                Integer begin  = fileName.length()-5;
                String suffix = fileName.substring(begin,end);
                if(".xlsx".equals(suffix)){
                    files.add(tempList[i].toString());
                }
            }
            if (tempList[i].isDirectory()) {
                String directory = tempList[i].toString();
                recordFile(directory);
            }
        }
    }
    private String getChapter(String fileName){
        String res = null;
        try{
            res= fileName.split(" ")[0];
            if(NumberTool.isInteger(res)){
                return res;
            }
        }catch (Exception e){

        }
        try{
            res= fileName.split("-")[0];
            if(NumberTool.isInteger(res)){
                return res;
            }
        }catch (Exception e){

        }
        return null;
    }
}
