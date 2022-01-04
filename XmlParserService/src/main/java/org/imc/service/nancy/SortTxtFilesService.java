package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.maven.surefire.shade.org.apache.commons.lang3.tuple.Pair;
import org.imc.tools.*;
import org.springframework.stereotype.Component;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.*;

@Component
@Slf4j
public class SortTxtFilesService {

    private List<String> files = new LinkedList<>();
    private String res = "";
    private Map<Integer,Map<Integer,String>> fileMap = new TreeMap<>();
    public void sort(String path) {
        log.info("开始记录文件");
        CommonTool.recordFile(files,path,".txt");
        // 移除隐藏文件
        log.info("开始移除隐藏文件");
        CommonTool.removeHideFiles(files);
        if(files.size()==0){
            log.error("没有找到要排序的文件");
            return;
        }

        //构造fileMap
        for(String file:files){
            String[] filePath = file.split("\\\\");
            String fileName = filePath[filePath.length-1];
            Pair<Integer,Integer> bookChapterPair = getTitle(fileName);
            Integer book = bookChapterPair.getLeft();
            Integer chapter = bookChapterPair.getRight();
            if(!fileMap.containsKey(book)){
                fileMap.put(book,new TreeMap<>());
            }
            fileMap.get(book).put(chapter,file);
        }

        Integer num = 1;
        File outputDirectory = new File("输出");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        for(Map.Entry<Integer,Map<Integer,String>> entry:fileMap.entrySet()){
            Map<Integer,String> files =  entry.getValue();
            for(Map.Entry<Integer,String> entry1:files.entrySet()){
                String file = entry1.getValue();
                String[] filePath = file.split("\\\\");
                String fileName = filePath[filePath.length-1];
                String[] fileNameSplit =  fileName.split(" ");
                String bookName = fileNameSplit[fileNameSplit.length-1];
                try {
                    String doc = FileImportUtil.readFile(file);
                    String outFilePath = "输出\\"+transTo4Bit(num)+" "+bookName;
                    FileExportUtil.buildNormalOutPutFile(outFilePath, doc);
                    num++;
                }catch (Exception e){
                    log.error("读取文件失败，文件："+file);
                }
            }
        }
        CommonTool.enterKeyContinue("重新编号完成，按回车键继续");

    }

    private String transTo4Bit(Integer num){
        Integer numCopy = num;
        Integer bit = 4;
        while(numCopy/10>0){
            bit--;
            numCopy = numCopy/10;
        }
        String zero ="";
        while(bit-->0){
            zero+="0";
        }
        zero+=num.toString();
        return zero;
    }

    private Pair<Integer,Integer> getTitle(String fileName){
        try{
            String[] emptyWordSplit = fileName.split(" ");
            if(emptyWordSplit.length>3){
                throw new Exception();
            }else if(emptyWordSplit.length == 3){
                String book = emptyWordSplit[0];
                String chapter = emptyWordSplit[1];
                book = book.substring(1,book.length()-1);
                chapter = chapter.substring(1,chapter.length()-1);
                int bookNum = DigitUtils.chineseNumber2Int(book);
                int chapterNum = DigitUtils.chineseNumber2Int(chapter);
                return Pair.of(bookNum, chapterNum);
            }else if(emptyWordSplit.length == 2){
                String bookChapter = emptyWordSplit[0];
                String[] bC = bookChapter.split("第");
                String book = bC[1].substring(0,bC[1].length()-1);
                String chapter = bC[2].substring(0,bC[2].length()-1);
                int bookNum = DigitUtils.chineseNumber2Int(book);
                int chapterNum = DigitUtils.chineseNumber2Int(chapter);
                return Pair.of(bookNum, chapterNum);
            }
        }catch (Exception e){
            log.error("文件名错误 或 处理失败，文件名："+fileName);
        }
        return null;
    }
}
