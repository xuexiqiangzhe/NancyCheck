package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.maven.surefire.shade.org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.imc.tools.CommonTool;
import org.imc.tools.DigitUtils;
import org.imc.tools.FileExportUtil;
import org.imc.tools.NumberTool;
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
    private List<ErrorEnterKey> enterKeyNotOkFiles = new LinkedList<>();

    class ErrorEnterKey{
        String fileName;
        List<Pair<Integer,String>> enterKeyPosition;
    }
    public void filter(String path) {
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".docx");
        // 移除隐藏文件
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);
        // 判断
        for(String file:files){
            judge(file);
        }

        // 输出
        SimpleDateFormat df = new SimpleDateFormat("yyMMdd-HH-mm-ss");//设置日期格式
        String time = df.format(new Date());
        log.info("当前时间:"+time);
        String outPutFilePath = ".\\"+time+"不合格文件(除换行符之外).txt";
        String content = "";
        for (String name : notOKFiles) {
            content+=name+"\n";
        }
        FileExportUtil.buildNormalOutPutFile(outPutFilePath, content);

        // 回车错误输出
        String outPutFilePath1 = ".\\"+time+"换行符不合格文件.txt";
        String content1 = "";
        for (ErrorEnterKey errorEnterKey : enterKeyNotOkFiles) {
            content1+=errorEnterKey.fileName+"\n";
            for(Pair<Integer,String> enterKeyPosition:errorEnterKey.enterKeyPosition){
                content1+="位置角标："+enterKeyPosition.getLeft()+"  亦此内容后的回车："+enterKeyPosition.getRight()+"\n";
            }
        }
        FileExportUtil.buildNormalOutPutFile(outPutFilePath1, content1);
        CommonTool.enterKeyContinue("不合格文件检查完毕，按回车退出");
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
            String[] filePath = path.split("\\\\");
            String fileName = filePath[filePath.length-1];
            if(!isOk(doc,fileName)){
                notOKFiles.add(fileName);
                System.out.println(path+"File format is not OK");
//                if(file.delete()){
//                    System.out.println(path+" File deleted");
//                }else
//                    System.out.println("File "+path+" doesn't exist");
            }else{
                System.out.println(path+"File format is OK");
            }
            Pair<Boolean,List<Integer>> enterKeyJudgeResult = enterKeyIsOk(doc);

            // 新增判断回车
            if(!enterKeyJudgeResult.getLeft()){
                ErrorEnterKey errorEnterKey = new ErrorEnterKey();
                errorEnterKey.fileName=fileName;
                errorEnterKey.enterKeyPosition = new LinkedList<>();
                List<Integer> enterKeyErrorPos = enterKeyJudgeResult.getRight();
                for(Integer pos:enterKeyErrorPos){
                    errorEnterKey.enterKeyPosition.add(Pair.of(pos+1,doc.substring(pos-25,pos)));
                }
                enterKeyNotOkFiles.add(errorEnterKey);
                System.out.println(path+"File enterKey is not OK");
            }else{
                System.out.println(path+"File enterKey is OK");
            }
        } catch (Exception e) {
            log.error("读取文件失败或解析Bug");
            CommonTool.enterKeyContinue("某文件读取失败，请联系程序员");
        }
    }

    private Boolean isOk(String doc,String fileName) {
        // 1.检查文件名提取章节号
        String[] fileNamePart = fileName.split(" ");
        if(fileNamePart.length!=3){
            return false;
        }
        String chapter = fileNamePart[1].substring(1,fileNamePart[1].length()-1);
        Integer chapterNum;
        if(!NumberTool.isInteger(chapter)){
            try {
                chapterNum = DigitUtils.chineseNumber2Int(chapter);
                chapter = String.valueOf(chapterNum);
            }catch (Exception e){
                return false;
            }
        }

        // 2.检查内容
        if(doc==null||doc.length()<15){
            return false;
        }
        // 2.1 条件1
        String head = doc.substring(0,15);
        if(!"***ChapterName:".equals(head)){
            return false;
        }
        // 2.2 条件2 和标题的章节号不等则不合格
        try{
            String chapterDoc = getChapterDoc(doc);
            if(!NumberTool.isInteger(chapterDoc)||!chapter.equals(chapterDoc)){
                return false;
            }
        }catch (Exception e){
            return false;
        }

        // 2.3 条件3
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
        if(doc.length()<=i+12){
            return false;
        }
        String tail = doc.substring(i,i+12);
        if(!"***Explicit:".equals(tail)){
            return false;
        }

        // 2.4 空格+false/true+若干个空格+换行符
        int j = i+12;
        for(;j<doc.length();j++){
            if('\n'==doc.charAt(j)){
                break;
            }
        }
        if('\n'!=doc.charAt(j)){
            return false;
        }
        tail = doc.substring(i+12,j);
        if(' '!=tail.charAt(0)||tail.length()<5){
            return false;
        }
        if('t'==tail.charAt(1)){
            if(tail.length()<5||!"true".equals(tail.substring(1,5))){
                return false;
            }
            if (!isConsistEmpty(tail.substring(5))) {
                return false;
            }
        }else if('f'==tail.charAt(1)){
            if(tail.length()<6||!"false".equals(tail.substring(1,6))){
                return false;
            }
            if (!isConsistEmpty(tail.substring(6))) {
                return false;
            }
        }else{
            return false;
        }
        return true;
    }


    private Pair<Boolean,List<Integer>> enterKeyIsOk(String doc) {
        List<Integer> enterKeyPosition = new LinkedList<>();
        for(int i=0;i<doc.length();i++){
            if('\n'==doc.charAt(i)){
                enterKeyPosition.add(i);
            }
        }

        List<Integer> errorPosition = new LinkedList<>();
        for(int i=2;i<enterKeyPosition.size()-1;i++){
            //是段落结尾的换行符且，接着不为空行
            Integer pos = enterKeyPosition.get(i);
            if(doc.charAt(pos-1)!='\n'){
                if(doc.charAt(pos+1)!='\n'){
                    errorPosition.add(pos);
                }
            }
        }
        if(errorPosition.size()>0){
            return Pair.of(false,errorPosition);
        }
        return Pair.of(true,errorPosition);
    }

    private boolean isConsistEmpty(String tail) {
        for (int k = 0; k < tail.length(); k++) {
            if (' ' != tail.charAt(k)) {
                return false;
            }
        }
        return true;
    }

    private String getChapterDoc(String doc) throws Exception {
        int count = 0;
        int begin =0;
        int end =0;
        for(int i=14;i<doc.length();i++){
            if(' '==doc.charAt(i)){
                count++;
                if(count==2){
                    begin = i+1;
                }else if(count==3){
                    end = i;
                    break;
                }
            }else if('\n'==doc.charAt(i)){
                throw new Exception();
            }
        }
        if(begin==0||end==0){
            throw new Exception();
        }
        return doc.substring(begin,end);
    }
}
