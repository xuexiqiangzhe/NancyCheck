package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.imc.tools.CommonTool;
import org.imc.tools.NumberTool;
import org.imc.tools.FileExportUtil;
import org.springframework.stereotype.Component;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

@Component
@Slf4j
public class CombineFileService {


    private List<String> files = new LinkedList<>();
    private String res = "";
    private Map<Integer,String> fileMap = new TreeMap<>();
    public void combine(String path) {
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".docx");
        // 移除隐藏文件
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);
        if(files.size()==0){
            log.error("没有找到要合并的文件");
            return;
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
                CommonTool.enterKeyContinue("中途，读取Txt失败，请联系程序员");
            }
        }

        SimpleDateFormat df = new SimpleDateFormat("yyMMdd-HH-mm-ss");//设置日期格式
        String time = df.format(new Date());
        String outFilePath = ".\\"+"合并文件"+time+".txt";
        String outDocFilePath = ".\\"+"合并文件"+time+".docx";
        FileExportUtil.buildNormalOutPutFile(outFilePath, res);
        FileExportUtil.exportDocx(outDocFilePath,res);
        // 检查连续性
        checkChapterContinuity();
        CommonTool.enterKeyContinue("合并完成，按回车键继续");

    }

    private void checkChapterContinuity() {
        List<Integer> chapterList = new LinkedList<>();
        for(Map.Entry<Integer,String> entry:fileMap.entrySet()) {
            chapterList.add(entry.getKey());
        }

        System.out.println("最小章节："+chapterList.get(0)+", "+"最大章节："+chapterList.get(chapterList.size()-1));
        boolean flag = false;
        for(int i=1;i<chapterList.size();i++){
            int cur = chapterList.get(i);
            int pre = chapterList.get(i-1);
            if(cur-pre>1){
                flag =true;
                System.out.print("漏掉的章节：");
                for(int j=pre+1;j<cur;j++){
                    System.out.print(j+"\t");
                }
                System.out.print("\n");
            }
        }
        if(!flag){
            System.out.println("章节连续，不存在遗漏");
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
