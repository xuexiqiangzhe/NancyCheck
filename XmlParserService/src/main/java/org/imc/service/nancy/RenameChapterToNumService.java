package org.imc.service.nancy;

import lombok.extern.slf4j.Slf4j;
import org.apache.maven.surefire.shade.org.apache.commons.lang3.tuple.Pair;
import org.imc.tools.CommonTool;
import org.imc.tools.DigitUtils;
import org.imc.tools.FileExportUtil;
import org.imc.tools.FileImportUtil;
import org.springframework.stereotype.Component;

import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

@Component
@Slf4j
public class RenameChapterToNumService {

    private List<String> files = new LinkedList<>();
    private String res = "";
    private Map<Integer,String> fileMap = new TreeMap<>();
    private List<Integer> noConsistChapters = new LinkedList<>();

    public void rename(String path) {
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
            Integer chapter = getChapter(fileName);
            if(fileMap.containsKey(chapter)){
                log.error(file+"的章节号与"+fileMap.get(chapter)+"重复");
                CommonTool.enterKeyContinue("重命名中途错误，按回车键结束");
                return;
            }
            fileMap.put(chapter,file);
        }


        List<Integer> chapterList = new LinkedList<>(fileMap.keySet());
        for(int i = 1;i<chapterList.size();i++){
            int j = i-1;
            if(chapterList.get(i)-chapterList.get(j)!=1){
                for(int k = chapterList.get(j)+1;k<chapterList.get(i);k++){
                    noConsistChapters.add(k);
                }
            }
        }

        System.out.print("请注意检查，不连续的章节包括:");
        for (int i = 0;i<noConsistChapters.size();i++){
            System.out.print(noConsistChapters.get(i)+", ");
        }
        System.out.print("\n");

        File outputDirectory = new File("输出");
        if(!outputDirectory.isDirectory()){
            outputDirectory.mkdir();
        }
        for(Map.Entry<Integer,String> entry:fileMap.entrySet()){
                String file = entry.getValue();
                Integer chapter = entry.getKey();
                String[] filePath = file.split("\\\\");
                String fileName = filePath[filePath.length-1];
                String[] fileNameSplit =  fileName.split(" ");
                String bookName = fileNameSplit[fileNameSplit.length-1];
                try {
                    String doc = FileImportUtil.readFile(file);
                    String outFilePath = "输出\\"+transTo5Bit(chapter)+" "+bookName;
                    FileExportUtil.buildNormalOutPutFile(outFilePath, doc);
                }catch (Exception e){
                    log.error("读取文件失败，文件："+file);
                }
            }

        CommonTool.enterKeyContinue("重新编号完成，按回车键继续");

    }

    private String transTo5Bit(Integer num){
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

    private Integer getChapter(String fileName){
        try{
            String[] emptyWordSplit = fileName.split(" ");
            if(emptyWordSplit.length>3) {
                 throw new Exception();
            }
            String bookChapter = emptyWordSplit[0];
            String chapter = bookChapter.substring(1,bookChapter.length()-1);
            if(bookChapter.charAt(0)!='第'||bookChapter.charAt(bookChapter.length()-1)!='章'){
                throw new Exception();
            }
            int chapterNum = DigitUtils.chineseNumber2Int(chapter);
           return chapterNum;
        }catch (Exception e){
            log.error("文件名错误 或 处理失败，文件名："+fileName);
        }
        return null;
    }
}
