package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

@Component
@Slf4j
public class FilterFilesService {


    private List<String> files = new LinkedList<>();
    private List<String> notOKFiles = new LinkedList<>();

    public void filter(String path) {
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
        String outPutFilePath = ".\\"+time+"不合格文件.txt";
        File file = new File(outPutFilePath);
        try {
            file.createNewFile();
            for (String name : notOKFiles) {
                buildOutPut(file, name);
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
          //  System.out.println(doc);
            fis.close();
            if(!isOk(doc)){
                String[] filePath = path.split("\\\\");
                notOKFiles.add(filePath[filePath.length-1]);
                System.out.println(path+"File format is not OK");
//                if(file.delete()){
//                    System.out.println(path+" File deleted");
//                }else
//                    System.out.println("File "+path+" doesn't exist");
            }else{
                System.out.println(path+"File format is OK");
            }
        } catch (Exception e) {
            log.error("读取文件失败或解析Bug");
        }
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
    private Boolean isOk(String doc) {
        if(doc==null||doc.length()<15){
            return false;
        }
        String head = doc.substring(0,15);
        if(!"***ChapterName:".equals(head)){
            return false;
        }
        int i = 14;
        for(;i<doc.length();i++){
            if('\n'==doc.charAt(i)){
                i++;
                break;
            }
        }
        if(i>=doc.length()){
            return  false;
        }
        if(doc.length()<i+12){
            return false;
        }
        String tail = doc.substring(i,i+12);
        if(!"***Explicit:".equals(tail)){
            return false;
        }
        return true;
    }
}
