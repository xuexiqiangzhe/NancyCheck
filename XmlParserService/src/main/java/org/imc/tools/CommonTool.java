package org.imc.tools;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.List;

@Component
@Slf4j
public class CommonTool {
    public static void removeHideFiles(List<String> files) {
        for(int i = 0;i<files.size();i++){
            String file = files.get(i);
            String[] filePathList = file.split("\\\\");
            String fileName = filePathList[filePathList.length-1];
            log.info("记录文件:"+file);
            if("~$".equals(fileName.substring(0,2))){
                log.info("隐藏文件,已经剔除。因为开头是："+fileName.substring(0,2));
                files.remove(i);
                i--;
            }
        }
    }

    public static void recordFile(List<String> files,String path,String targetSuffix){
        File file = new File(path);
        File[] tempList = file.listFiles();
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile()) {
                String fileName = tempList[i].toString();
                Integer end = fileName.length();
                Integer begin  = fileName.length()-5;
                String suffix = fileName.substring(begin,end);
                if(targetSuffix.equals(suffix)){
                    files.add(tempList[i].toString());
                }
            }
            if (tempList[i].isDirectory()) {
                String directory = tempList[i].toString();
                recordFile(files,directory,targetSuffix);
            }
        }
    }

    public static void enterKeyContinue(String infomation) {
        System.out.println("-------------"+infomation+"-----------------");
        try {
            new BufferedReader(new InputStreamReader(System.in)).readLine();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
