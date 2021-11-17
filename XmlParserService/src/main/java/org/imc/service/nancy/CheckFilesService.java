package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.imc.tools.NumberTool;
import org.springframework.stereotype.Component;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

@Component
@Slf4j
public class CheckFilesService {


    private List<String> files = new LinkedList<>();
    private List<String> notOKFiles = new LinkedList<>();
    private List<String> titleNotOKFiles = new LinkedList<>();

    public void check(String path) {
        log.info("开始记录文件");
        recordFile(path);
        for(String file:files){
            log.info("已记录文件"+file);
        }
        for(String file:files){
            judge(file);
        }
        SimpleDateFormat df = new SimpleDateFormat("yyMMdd-HH-mm-ss");//设置日期格式
        String time = df.format(new Date());
        log.info("当前时间:"+time);
        String tileErrorFilePath = ".\\"+"标题不合格文件"+time+".txt";
        File tileErrorFile = new File(tileErrorFilePath);

        String contentErrorFilePath = ".\\"+"内容格式不合格文件"+time+".txt";
        File contentErrorFile = new File(contentErrorFilePath);
        try {
            contentErrorFile.createNewFile();
            for (String name : notOKFiles) {
                buildOutPut(contentErrorFile, name);
            }
            tileErrorFile.createNewFile();
            for (String name : titleNotOKFiles) {
                buildOutPut(tileErrorFile, name);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void buildOutPut(File file, String name) throws IOException {
        // write
        FileWriter fw = new FileWriter(file, true);
        BufferedWriter bw = new BufferedWriter(fw);
        bw.write(name + "\n");
        bw.flush();
        bw.close();
        fw.close();
    }


    private void judge(String path){
        File file = new File(path);
        try {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument xdoc = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            String doc = extractor.getText();
            fis.close();
            String[] filePath = path.split("\\\\");
            String fileName = filePath[filePath.length-1];
            String chapter = getChapter(fileName);
            if(chapter==null){
                titleNotOKFiles.add(fileName);
                System.out.println(path+"File title is not OK");
                return ;
            }
            Integer chapterNum = Integer.valueOf(chapter);
            if(!isOk(doc,String.valueOf(chapterNum))){
                notOKFiles.add(fileName);
                System.out.println(path+"File content format is not OK");
            }else{
                System.out.println(path+"File title and content format is OK");
            }
        } catch (Exception e) {
            log.error("读取文件失败或解析Bug");
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
    private Boolean isOk(String doc,String chapterNum) {
        String standard = "##Chapter "+chapterNum+" ";
        if(doc==null||doc.length()<standard.length()){
            return false;
        }
        if(standard.equals(doc.substring(0,standard.length()))){
            return true;
        }
        return false;
    }

}
