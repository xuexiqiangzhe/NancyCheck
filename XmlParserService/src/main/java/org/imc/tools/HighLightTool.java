package org.imc.tools;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.TextSegment;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTEmptyImpl;
import org.w3c.dom.Node;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

import javax.xml.namespace.QName;
import java.io.*;
import java.math.BigInteger;
import java.util.*;

/**
 * @program: ruoyi
 * @create: 2021-10-08 11:33
 * @author: sxl
 * @description:
 **/
public class HighLightTool {
    public static void highLightAndOutPut(String filePath, Set<String> keys,String language) throws Exception {
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);
        for(String key:keys){
            for (XWPFParagraph p : doc.getParagraphs()) {
                String paraDoc = p.getText();
                if (paraDoc.contains(key)) {
                    List<TextSegment> segments = searchText(p, key, new PositionInParagraph());
                    List<XWPFRun> runs = p.getRuns();
                    //一段里可能有多个出现关键字的字块
                    for(int j=0;j<segments.size();j++){
                        TextSegment segment = segments.get(j);
                        int beginRunIndex = segment.getBeginRun();
                        int endRunIndex = segment.getEndRun();
                        int beginPosChar = segment.getBeginChar();
                        int endPosChar = segment.getEndChar();
                        if (beginRunIndex == endRunIndex) {
                            XWPFRun run = runs.get(beginRunIndex);
                            String runDoc = run.getText(0);
                            String beforeDoc = runDoc.substring(0,beginPosChar);
                            String endDoc = runDoc.substring(endPosChar+1);
                            String targetDoc = runDoc.substring(beginPosChar,endPosChar);
                            if(beforeDoc.length()>0){

                            }
                            if(endDoc.length()>0){

                            }
                            highLight(p, run);
                        } else {
                            // 高亮中间的文本
                            for (int i = beginRunIndex + 1; i < endRunIndex; i++) {
                                XWPFRun run = runs.get(i);
                                highLight(p, run);
                            }
                        }

                    }
                }
            }
        }
        String[] filePathArray = filePath.split("\\\\");
        String fileName = filePathArray[filePathArray.length - 1];
        File file = new File("输出\\"+language+"\\"+fileName);
        FileOutputStream out = new FileOutputStream(file);
        doc.write(out);
        out.close();
        doc.close();
    }


    private static void highLight(XWPFParagraph p, XWPFRun run) {
        CTRPr pRpr = getRunCTRPr(p, run);
        CTHighlight highlight = pRpr.getHighlightList().size()>0 ? pRpr.getHighlightList().get(0) : pRpr.addNewHighlight();

        highlight.setVal(STHighlightColor.YELLOW);
    }


    /**
     * 得到XWPFRun的CTRPr
     */
    public static CTRPr getRunCTRPr(XWPFParagraph p, XWPFRun pRun) {
        CTRPr pRpr;
        if (pRun.getCTR() != null) {
            pRpr = pRun.getCTR().getRPr();
            if (pRpr == null) {
                pRpr = pRun.getCTR().addNewRPr();
            }
        } else {
            pRpr = p.getCTP().addNewR().addNewRPr();
        }
        return pRpr;
    }

    /**
     * POI本身的searchText不排除CTEmptyImpl的情况导致查不到文本
     */
    public static List<TextSegment> searchText(XWPFParagraph paragraph, String searched, PositionInParagraph startPos) {
        int startRun = startPos.getRun(),
                startText = startPos.getText(),
                startChar = startPos.getChar();
        int beginRunPos = 0, candCharPos = 0;
        boolean newList = false;
        List<TextSegment> segList = new ArrayList<>();
        CTR[] rArray = paragraph.getCTP().getRArray();
        for (int runPos = startRun; runPos < rArray.length; runPos++) {
            int beginTextPos = 0, beginCharPos = 0, textPos = 0, charPos = 0;
            CTR ctRun = rArray[runPos];
            XmlCursor c = ctRun.newCursor();
            c.selectPath("./*");
            try {
                while (c.toNextSelection()) {
                    XmlObject o = c.getObject();
                    if (o instanceof CTText) {
                        if (textPos >= startText) {
                            String candidate = ((CTText) o).getStringValue();
                            if (runPos == startRun) {
                                charPos = startChar;
                            } else {
                                charPos = 0;
                            }

                            for (; charPos < candidate.length(); charPos++) {
                                if ((candidate.charAt(charPos) == searched.charAt(0)) && (candCharPos == 0)) {
                                    beginTextPos = textPos;
                                    beginCharPos = charPos;
                                    beginRunPos = runPos;
                                    newList = true;
                                }
                                if (candidate.charAt(charPos) == searched.charAt(candCharPos)) {
                                    if (candCharPos + 1 < searched.length()) {
                                        candCharPos++;
                                    } else if (newList) {
                                        TextSegment segment = new TextSegment();
                                        segment.setBeginRun(beginRunPos);
                                        segment.setBeginText(beginTextPos);
                                        segment.setBeginChar(beginCharPos);
                                        segment.setEndRun(runPos);
                                        segment.setEndText(textPos);
                                        segment.setEndChar(charPos);
                                        segList.add(segment);
                                    }
                                } else {
                                    candCharPos = 0;
                                }
                            }
                        }
                        textPos++;
                    } else if (o instanceof CTProofErr) {
                        c.removeXml();
                    } else if (o instanceof CTRPr || o instanceof CTEmptyImpl) {
                        //do nothing
                    } else {
                        candCharPos = 0;
                    }
                }
            } finally {
                c.dispose();
            }
        }
        return  segList;
    }

}