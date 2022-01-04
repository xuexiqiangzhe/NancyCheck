package org.imc.service.nancy;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.TextSelection;
import lombok.extern.slf4j.Slf4j;
import org.imc.service.nancy.excel.ExcelData;
import org.imc.tools.CommonTool;
import org.springframework.stereotype.Component;

import java.awt.*;
import java.io.File;
import java.util.*;
import java.util.List;

@Component
@Slf4j
public class WordHighLightService {


    private List<String> enFiles = new LinkedList<>();
    private List<String> cnFiles = new LinkedList<>();
    private Set<String> replaceEnSet = new HashSet<>();
    private Set<String> replaceCnSet = new HashSet<>();
    public void highLight(String path) {
        // 1.构建替换表
        buildWordsSet(path);
        log.info("开始记录文件");
        CommonTool.recordFile(enFiles,path+"\\英文",".docx");
        CommonTool.recordFile(cnFiles,path+"\\中文",".docx");
        // 移除隐藏文件
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(enFiles);
        CommonTool.removeHideFiles(cnFiles);
        File outputDirectory = new File("输出");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        File outputDirectory2 = new File("输出\\中文");
        if(!outputDirectory2.isDirectory()){
            outputDirectory2.mkdir();
        }
        File outputDirectory1 = new File("输出\\英文");
        if(!outputDirectory1.isDirectory()){
            outputDirectory1.mkdir();
        }
        // 高亮英文
        highLightALanguage(enFiles,replaceEnSet, "英文");
        // 高亮中文
        highLightALanguage(cnFiles, replaceCnSet,"中文");
        CommonTool.enterKeyContinue("高亮结束,按回车退出");
    }

    private void highLightALanguage(List<String> files,Set<String> replaceSet, String language) {
        for (String file : files) {
            try {
                String[] filePath = file.split("\\\\");
                String fileName = filePath[filePath.length - 1];
                Document doc = new Document();
                doc.loadFromFile(file);
                replaceBySpire(doc, fileName, language,replaceSet);
            } catch (Exception e) {
                log.info("替换发生异常，中断程序了，请联系开发，目标文件：" + file);
                e.printStackTrace();
            }
        }
    }

    private void buildWordsSet(String path) {
        String replaceDictionary = path +"\\术语表.xlsx";
        ExcelData sheet = new ExcelData(replaceDictionary, "术语表");
        int rows = sheet.getNumberOfRows();
        //记录英文术语
        for(int x = 0;x<rows;x++){
            try {
                String src = sheet.getExcelDateByIndex(x, 1);
                if ("".equals(src)) {
                    continue;
                }
                if (!replaceEnSet.contains(src)) {
                    replaceEnSet.add(src);
                }
            }catch (Exception e){
                log.error("读取术语表异常，位置:{}行，请先检查，无法解决则联系开发",x+1);
            }
        }
        //记录中文术语
        for(int x = 0;x<rows;x++){
            try {
                String src = sheet.getExcelDateByIndex(x, 0);
                if ("".equals(src)) {
                    continue;
                }
                if (!replaceCnSet.contains(src)) {
                    replaceCnSet.add(src);
                }
            }catch (Exception e){
                log.error("读取术语表异常，位置:{}行，请先检查，无法解决则联系开发",x+1);
            }
        }
    }

    private void replaceBySpire(Document doc,String fileName,String language,Set<String> replaceSet){
        for(String src: replaceSet){
            try {
                //调用方法用新文本替换原文本内容
                //查找所有需要高亮的文本
                TextSelection[] textSelections = doc.findAllString(src, false, false);
                //设置高亮颜色
                if(textSelections==null){
                    continue;
                }
                for (TextSelection selection : textSelections) {
                    selection.getAsOneRange().getCharacterFormat().setHighlightColor(Color.YELLOW);
                }
            }catch (Exception e){
                log.error("术语:"+src+" 替换失败");
                return;
            }
        }
        //保存文档
        doc.saveToFile("输出\\"+language+"\\"+fileName,FileFormat.Docx_2013);
        doc.dispose();
    }

}
