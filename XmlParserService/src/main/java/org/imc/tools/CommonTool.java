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

    public static String returnRegularExpressionString(String str){
        String reg = "\\\\d+";
      //  String reg = "\\d+";
        String reg1 = "\\\\:";
        String reg2 = "\\\\+";
        String reg3 = "\\\\(";
        String reg4 = "\\\\)";
        String reg5 = "\\\\[";
        String reg6 = "\\\\]";
        String reg7 = "\\\\$";
        String reg8 = "\\\\*";
        String tmp = str.replaceAll("\\+",reg2);
        String tmp1 = tmp.replaceAll(reg,"\\\\d+");
        String tmp2 = tmp1.replaceAll("\\:",reg1);
        String tmp3 = tmp2.replaceAll("\\(",reg3);
        String tmp4 = tmp3.replaceAll("\\)",reg4);
        String tmp5 = tmp4.replaceAll("\\[",reg5);
        String tmp6 = tmp5.replaceAll("\\]",reg6);
        String tmp7 = tmp6.replaceAll("\\$",reg7);
        String tmp8 = tmp7.replaceAll("\\*",reg8);
        return tmp8;
    }
}
