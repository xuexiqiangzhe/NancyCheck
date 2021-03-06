package org.imc.tools;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Component;

import javax.security.auth.login.Configuration;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
@Slf4j
public class FileExportUtil {

    public static void exportDocx(String filePath,String content) {
        try {
            XWPFDocument doc = new XWPFDocument(); //创建word文件
            String[] paragraphs = content.split("\n");
            for(String paragraph:paragraphs){
                XWPFParagraph p1 = doc.createParagraph(); //创建段落
                XWPFRun r1 = p1.createRun(); //创建段落文本
                r1.setText(paragraph); //设置文本

            }
            FileOutputStream out = new FileOutputStream(filePath); //创建输出流
            doc.write(out);  //输出
            out.close();  //关闭输出流
        } catch (IOException e) {
            log.error("生成"+filePath+"docx失败");
        }
    }

    /**
     * 输出：字体小四，字体Arial，第一行的标题居中
     * @param filePath
     * @param content
     */
    public static void exportSpecificDocx(String filePath,String content) {
        try {
            XWPFDocument doc = new XWPFDocument(); //创建word文件
            String[] paragraphs = content.split("\n");
            for(String paragraph:paragraphs){
                XWPFParagraph p1 = doc.createParagraph(); //创建段落
                XWPFRun r1 = p1.createRun(); //创建段落文本
                r1.setText(paragraph); //设置文本
//
//                CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
//                CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
//                fonts.setAscii("仿宋");
//                fonts.setEastAsia("仿宋");
//                fonts.setHAnsi("仿宋");
                r1.setFontSize(12);//设置字体大小
                r1.setFontFamily("Arial");//设置字体
            }
            XWPFParagraph firstParagraph = doc.getParagraphs().get(0);
            firstParagraph.setAlignment(ParagraphAlignment.CENTER);
            FileOutputStream out = new FileOutputStream(filePath); //创建输出流
            doc.write(out);  //输出
            out.close();  //关闭输出流
        } catch (IOException e) {
            log.error("生成"+filePath+"docx失败");
        }
    }

    public static void buildNormalOutPutFile(File file, String content){
        // write
        try {
            FileWriter fw = new FileWriter(file, true);
        BufferedWriter bw = new BufferedWriter(fw);
        bw.write(content);
        bw.flush();
        bw.close();
        fw.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void buildNormalOutPutFile(String filePath,String content){
        try {
            BufferedWriter writer = new BufferedWriter (new OutputStreamWriter (new FileOutputStream (filePath,true),"UTF-8"));
            writer.write(content);
            writer.flush();
            writer.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
