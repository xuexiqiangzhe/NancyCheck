package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.imc.tools.NumberTool;
import org.springframework.stereotype.Component;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

@Component
@Slf4j
public class CombineFileService {


    private List<String> files = new LinkedList<>();
    private String res = "";
    private Map<Integer,String> fileMap = new LinkedHashMap<>();
    public void combine(String path) {
        log.info("开始记录文件");
        recordFile(path);
        for(String file:files){
            log.info("已记录文件"+file);
        }

        for(String file:files){
            String[] filePath = file.split("\\\\");
            String fileName = filePath[filePath.length-1];
            String chapter = getChapter(fileName);
            if(chapter==null){
                System.out.println("-------------"+fileName+"获取章节号失败，按回车键继续-----------------");
                try {
                    new BufferedReader(new InputStreamReader(System.in)).readLine();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                return;
            }
            Integer chapterNum = Integer.valueOf(chapter);
            fileMap.put(chapterNum,file);
        }

        for(Map.Entry<Integer,String> entry:fileMap.entrySet()){
            String file =  entry.getValue();
            try {
                FileInputStream fis = new FileInputStream(file);
                XWPFDocument xdoc = new XWPFDocument(fis);
                XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
                String doc = extractor.getText();
                fis.close();
                res+=doc;
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        SimpleDateFormat df = new SimpleDateFormat("yyMMdd-HH-mm-ss");//设置日期格式
        String time = df.format(new Date());
        log.info("当前时间:"+time);
        String outFilePath = ".\\"+"合并文件"+time+".txt";
        String outDocFilePath = ".\\"+"合并文件"+time+".docx";
        File outFile = new File(outFilePath);
        File outDocFile = new File(outDocFilePath);
        try {
            buildOutPut(outFile, res);
            buildOutPut(outDocFile, res);
        } catch (IOException e) {
            e.printStackTrace();
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
                if(".docx".equals(suffix)){
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
