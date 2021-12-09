package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.imc.service.nancy.excel.ExcelData;
import org.imc.tools.CommonTool;
import org.imc.tools.FileExportUtil;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.*;

@Component
@Slf4j
public class WordReplaceService {


    private List<String> files = new LinkedList<>();
    private Map<String,String> replaceMap = new HashMap<>();
    public void replaceWord(String path) {
        // 1.构建替换表
        buildReplaceMap(path);
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
              FileInputStream fis = new FileInputStream(file);
              XWPFDocument xdoc = new XWPFDocument(fis);
              XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
              String doc = extractor.getText();
              String[] filePath = file.split("\\\\");
              String fileName = filePath[filePath.length-1];
              replace(fileName,doc,replaceMap);
              fis.close();
          } catch (Exception e){
              log.info("替换发生异常，中断程序了，请联系开发，目标文件："+file);
              e.printStackTrace();
          }
        }

    }

    private void buildReplaceMap(String path) {
        String replaceDictionary = path +"\\术语替换表.xlsx";
        ExcelData sheet = new ExcelData(replaceDictionary, "术语替换表");
        int rows = sheet.getNumberOfRows();
        for(int x = 1;x<rows;x++){
            String src = sheet.getExcelDateByIndex(x, 0);
            String tar = sheet.getExcelDateByIndex(x, 1);
            if("".equals(src)){
                continue;
            }
            if(!replaceMap.containsKey(src)){
                replaceMap.put(src,tar);
            }else{
                log.error("术语:"+src+" 重复,无法进行下一步解析");
                return;
            }
        }
    }

    private void replace(String fileName,String doc,Map<String,String> replaceMap){
        for(Map.Entry<String,String> entry:replaceMap.entrySet()){
            String src = entry.getKey();
            String tar = entry.getValue();
            doc = doc.replaceAll(src,tar);
//            FileExportUtil.exportDocx("输出\\"+fileName,res);
        }
        String name = fileName.substring(0,fileName.length()-5);
        String suffix = ".txt";
        File file = new File("输出\\"+name+suffix);
        FileExportUtil.buildNormalOutPutFile(file,doc);
    }
}
