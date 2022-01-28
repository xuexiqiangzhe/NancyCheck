package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.map.LinkedMap;
import org.apache.maven.surefire.shade.org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.imc.tools.CommonTool;
import org.imc.tools.DigitUtils;
import org.imc.tools.FileExportUtil;
import org.imc.tools.NumberTool;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.*;

@Component
@Slf4j
public class SplitFileService {
    private List<String> files = new LinkedList<>();
    private Map<String,List<FileGenerated>> fileGeneratedMap = new LinkedHashMap<>();
    private List<Pair<String,String>> errorPositions = new LinkedList<>();
    private class Chapter{
        Integer chapterStart;
        String chapterTitle;
        Chapter(Integer start,String title){
            chapterStart = start;
            chapterTitle = title;
        }
    }

    private class FileGenerated{
        String rowTitle;
        String fileName;
        String content;
        String outFilePath;
        FileGenerated(String rowT,String fileN,String cont,String path){
            rowTitle=rowT;
            fileName=fileN;
            content=cont;
            outFilePath = path;
        }
    }

    public void split(String path) {
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".docx");
        // 移除隐藏文件
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);
        if(files.size()==0){
            log.error("没有找到要拆分的文件");
            return;
        }

        for(String file:files){
            try {
                FileInputStream fis = new FileInputStream(file);
                XWPFDocument xdoc = new XWPFDocument(fis);
                XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
                String doc = extractor.getText();
                fis.close();
                String[] filePath = file.split("\\\\");
                String fileName= filePath[filePath.length-1];
                splitDoc(doc,fileName);
            } catch (Exception e) {
                CommonTool.enterKeyContinue("中途，读取Docx失败，请联系程序员");
            }
        }

        // 检查所有输出
        for(Map.Entry<String,List<FileGenerated>> entry:fileGeneratedMap.entrySet()){
           String fileName = entry.getKey();
           List<FileGenerated> fileGenerateds = entry.getValue();
            for(int i =0;i<fileGenerateds.size();i++){
                FileGenerated fileGenerated = fileGenerateds.get(i);
                String outFileName = fileGenerated.fileName;
                if(outFileName==null){
                    errorPositions.add(Pair.of(fileName,fileGenerated.rowTitle));
                }
            }
        }

        // 错的生成一个Txt提示错误
        if(errorPositions.size()>0){
            String error="";
            for(Pair<String,String> pair:errorPositions){
                error+=pair.getLeft()+" "+pair.getRight()+"\n";
            }
            String outFilePath = ".\\"+"不合格文件的位置"+".txt";
            FileExportUtil.buildNormalOutPutFile(outFilePath, error);
            CommonTool.enterKeyContinue("拆分失败，请看错误文件，按回车键退出");
            return;
        }

        //没错则正常生成文档

        File outputDirectory = new File("输出");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        for(Map.Entry<String,List<FileGenerated>> entry:fileGeneratedMap.entrySet()){
            String fileName = entry.getKey();
            List<FileGenerated> fileGenerateds = entry.getValue();
            File outputDirectory1 = new File("输出\\"+fileName);
            if(!outputDirectory1.isDirectory()){
                outputDirectory1.mkdir();
            }
            //每个原始文档对应多个文件。
            for(int i =0;i<fileGenerateds.size();i++){
                FileGenerated fileGenerated = fileGenerateds.get(i);
                String content = fileGenerated.content;
                String outFilePath = fileGenerated.outFilePath;
                FileExportUtil.exportDocx(outFilePath,content);
            }
        }

        CommonTool.enterKeyContinue("拆分完成，按回车键继续");

    }

    private void splitDoc(String doc,String fileName) throws Exception {

        List<Integer> jinBegin = new LinkedList<>();
        for(int i=0;i<doc.length()-1;i++){
            if(doc.charAt(i)=='#'&&doc.charAt(i+1)=='#'){
                jinBegin.add(i);
            }
        }
        log.info("文件名：{},按既定规则识别出{}章",fileName,jinBegin.size());

        List<Chapter> chapters = new LinkedList<>();
        for(int i = 0;i<jinBegin.size();i++){
            Integer jin = jinBegin.get(i);
            if(doc.charAt(jin-1)=='\n'&&doc.charAt(jin-2)=='\n'){
               Integer titleIndex = buildTitleIndex(doc, jin);
               String title = doc.substring(titleIndex,jin-2);
               chapters.add(new Chapter(titleIndex,title));
            }else{
                throw new Exception();
            }
        }
        chapters.add(new Chapter(doc.length(),""));

        List<FileGenerated> fileGenerateds = new LinkedList<>();
        for(int i=0;i<chapters.size()-1;i++){
            String title = chapters.get(i).chapterTitle;
            Integer begin = chapters.get(i).chapterStart;
            Integer end = chapters.get(i+1).chapterStart-1;
            String content = doc.substring(begin,end+1);
            String outFileName = buildOutFileName(title);
            fileGenerateds.add(new FileGenerated(title,outFileName,content,"输出\\"+fileName+"\\"+outFileName+".docx"));
        }
        fileGeneratedMap.put(fileName,fileGenerateds);
    }

    private Integer buildTitleIndex(String doc, Integer jin) {
        Integer titleIndex = jin -3;
        for(; titleIndex >=0; titleIndex--){
            if(doc.charAt(titleIndex)!='\n'&& titleIndex !=0){
                continue;
            }
            break;
        }
        if(titleIndex !=0){
            titleIndex++;
        }
        return titleIndex;
    }

    private String buildOutFileName(String title) {
        if ("Volume".equals(title.substring(0, 6))) {
            int i = 7;
            for (; i < title.length(); i++) {
                if (title.charAt(i) == ' ') {
                    break;
                }
            }
            Integer volume = Integer.parseInt(title.substring(7, i));
            if ("Chapter".equals(title.substring(i + 1, i + 8))) {
                int j = i + 9;
                for (; j < title.length(); j++) {
                    if (title.charAt(j) == ' ') {
                        break;
                    }
                }
                Integer chapter = Integer.parseInt(title.substring(i + 9, j));
                return "第"+volume+"卷 第"+chapter+"章";
            }else{
                return null;
            }
        } else if ("Chapter".equals(title.substring(0, 7))) {
            int j = 8;
            for (; j < title.length(); j++) {
                if (title.charAt(j) == ' ') {
                    break;
                }
            }
            Integer chapter = Integer.parseInt(title.substring(8, j));
            String chapterString = transTo4Bit(chapter);
            return chapterString+" 第"+chapter+"章";
        }
        return null;
    }

    private String transTo4Bit(Integer num){
        Integer numCopy = num;
        Integer bit = 4;
        while(numCopy/10>0){
            bit--;
            numCopy = numCopy/10;
        }
        String zero ="";
        while(--bit>0){
            zero+="0";
        }
        zero+=num.toString();
        return zero;
    }
}
