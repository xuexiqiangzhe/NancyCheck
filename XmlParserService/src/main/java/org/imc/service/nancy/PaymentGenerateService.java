package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
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


            String cell2 = sheet.getExcelDateByIndex(1, 3);
//        //根据第3列值为“customer23”的这一行，来获取该行第2列的值
//        String cell3 = sheet1.getCellByCaseName("customer23", 2,1);
//        System.out.println(cell2);
//        System.out.println(cell3);
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
