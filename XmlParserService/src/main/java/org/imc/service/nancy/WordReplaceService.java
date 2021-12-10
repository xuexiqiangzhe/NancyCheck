package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.imc.service.nancy.excel.ExcelData;
import org.imc.tools.CommonTool;
import org.imc.tools.FileExportUtil;
import org.springframework.stereotype.Component;

import java.io.File;
import java.util.*;
import com.spire.doc.*;
@Component
@Slf4j
public class WordReplaceService {


    private List<String> files = new LinkedList<>();
    private Map<String,String> replaceMap = new HashMap<>();
    public void replaceWord(String path) {
        // 1.构建替换表
        List<String> duplicateKeys = buildReplaceMap(path);
        if(duplicateKeys.size()>1){
            System.out.print("重复术语:"+"\t");
            for(String duplicateKey:duplicateKeys){
                System.out.print(duplicateKey+",\t");
            }
            System.out.print("\n");
            CommonTool.enterKeyContinue("重复术语,按回车退出");
            return;
        }
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".docx");
        // 移除隐藏文件
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);
        File outputDirectory = new File("输出");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        for(String file:files){
          try{
//              FileInputStream fis = new FileInputStream(file);
//              XWPFDocument xdoc = new XWPFDocument(fis);
//              XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
//              String doc = extractor.getText();
//              replace(fileName,doc,replaceMap);
//              fis.close();
              String[] filePath = file.split("\\\\");
              String fileName = filePath[filePath.length-1];
              Document doc = new Document();
              doc.loadFromFile(file);
              replaceBySpire(doc,fileName);
          } catch (Exception e){
              log.info("替换发生异常，中断程序了，请联系开发，目标文件："+file);
              e.printStackTrace();
          }
        }
        CommonTool.enterKeyContinue("替换结束,按回车退出");
    }

    private List<String> buildReplaceMap(String path) {
        List<String> duplicateKeys =  new LinkedList<>();
        String replaceDictionary = path +"\\术语替换表.xlsx";
        ExcelData sheet = new ExcelData(replaceDictionary, "术语替换表");
        int rows = sheet.getNumberOfRows();
        for(int x = 1;x<rows;x++){
            try {
                String src = sheet.getExcelDateByIndex(x, 0);
                String tar = sheet.getExcelDateByIndex(x, 1);
                if ("".equals(src)) {
                    continue;
                }
                if (!replaceMap.containsKey(src)) {
                    replaceMap.put(src,tar);
                } else {
                    if(!replaceMap.get(src).equals(tar)){
                        duplicateKeys.add(src);
                    }
                }
            }catch (Exception e){
                log.error("读取术语替换表异常，位置:{}行，请先检查，无法解决则联系开发",x+1);
            }
        }
        return duplicateKeys;
    }

    private void replaceBySpire(Document doc,String fileName){
        for(Map.Entry<String,String> entry:replaceMap.entrySet()){
            try {
                String src = entry.getKey();
                String tar = entry.getValue();
                //调用方法用新文本替换原文本内容
                doc.replace(src, tar, false, true);
            }catch (Exception e){
                log.error("术语:"+entry.getKey()+" 替换失败");
                return;
            }
        }
        //保存文档
        doc.saveToFile("输出\\"+fileName,FileFormat.Docx_2013);
        doc.dispose();
    }

}
