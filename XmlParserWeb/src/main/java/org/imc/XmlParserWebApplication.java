package org.imc;

import org.imc.service.nancy.*;
import org.springframework.boot.ExitCodeGenerator;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.ComponentScan;

import java.util.Scanner;

@SpringBootApplication
@ComponentScan(basePackages = {"org.imc.*"})
public class XmlParserWebApplication {
    public static void main(String[] args) {

        ConfigurableApplicationContext ctx = SpringApplication.run(XmlParserWebApplication.class, args);
        System.out.println("请输入整数以调用的预定的程序");
        System.out.println("1 —————判断文档格式 ***ChapterName那个");
        System.out.println("2 —————检查文档格式 ##Chapter 313");
        System.out.println("3 —————合并文档");
        System.out.println("4 —————生成译费结算表");
        System.out.println("5 —————术语替换  需要,\"输入\\术语替换表.xlsx\"");
        System.out.println("6 —————合并txt文档");
        System.out.println("7 —————生成译员编号信息表");
        System.out.println("8 —————给txt卷章编号");
        System.out.println("9 —————高亮术语   需要,\"输入\\术语表.xlsx\" 和 \"输入\\中文\"………… 和 \"输入\\英文\"…………");
        System.out.println("10 ————根据章节号重命名文件 如：第一千章 盘古开天地.txt ->01000 盘古开天地.txt");
        System.out.println("11 ————文档拆分工具， 根据##与它之前的第2个和第3个回车之间的title");
        Scanner lll = new Scanner(System.in);
        int x = lll.nextInt();
        switch (x) {
            case 1:
                //  1. 判断文档格式***ChapterName:
                FilterFilesService filterFilesService = new FilterFilesService();
                String path = "输入";
                filterFilesService.filter(path);
                break;
            case 2:
                //2. 判断文档格式 ##Chapter 313
                CheckFilesService checkFilesService = new CheckFilesService();
                checkFilesService.check("输入");
                break;
            case 3:
                // 3.合并文档
                CombineFileService combineFileService = new CombineFileService();
                combineFileService.combine("输入");
                break;

            case 4:
                //4.生成结算表
                PaymentGenerateService paymentGenerateService = new PaymentGenerateService();
                paymentGenerateService.generate("输入");
                break;
            case 5:
                //5.术语替换
                WordReplaceService wordReplaceService = new WordReplaceService();
                wordReplaceService.replaceWord("输入");
                break;
            case 6:
                // 6.合并txt文档
                CombineTxtFileService combineTxtFileService = new CombineTxtFileService();
                combineTxtFileService.combine("输入");
                break;
            case 7:
                // 7.生成译员编号信息表
                TranslatorNumGenerateService translatorNumGenerateService = new TranslatorNumGenerateService();
                translatorNumGenerateService.generate("输入");
                break;
            case 8:
                // 7.生成译员编号信息表
                SortTxtFilesService sortTxtFilesService = new SortTxtFilesService();
                sortTxtFilesService.sort("输入");
                break;
            case 9:
                // 9.高亮术语
                WordHighLightService wordHighLightService = new WordHighLightService();
                wordHighLightService.highLight("输入");
                break;
            case 10:
                // 10.重命名文件（根据章节号）
                RenameChapterToNumService renameChapterToNumService = new RenameChapterToNumService();
                renameChapterToNumService.rename("输入");
                break;
            case 11:
                // 11.文档拆分
                SplitFileService splitFileService = new SplitFileService();
                splitFileService.split("输入");
                break;
                default:
                break;
        }

        exitApplication(ctx);
    }

    public static void exitApplication(ConfigurableApplicationContext context) {
        int exitCode = SpringApplication.exit(context, (ExitCodeGenerator) () -> 0);

        System.exit(exitCode);
    }
}
