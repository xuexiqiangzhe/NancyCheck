package org.imc;

import org.imc.service.nancy.CombineFileService;
import org.imc.service.nancy.PaymentGenerateService;
import org.springframework.boot.ExitCodeGenerator;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
@ComponentScan(basePackages = {"org.imc.*"})
public class XmlParserWebApplication {
    public static void main(String[] args) {

        ConfigurableApplicationContext ctx = SpringApplication.run(XmlParserWebApplication.class, args);

        //1. 判断文档格式***ChapterName:
//        FilterFilesService filterFilesService = new FilterFilesService();
//        String path = "过滤";
//        filterFilesService.filter(path);

        //2. 判断文档格式 ##Chapter 313
//        CheckFilesService checkFilesService = new CheckFilesService();
//        checkFilesService.check("检查");

        // 3.合并文档
//        CombineFileService combineFileService  = new CombineFileService();
//        combineFileService.combine("合并");
//

        //4.生成结算表
        PaymentGenerateService paymentGenerateService = new PaymentGenerateService();
        paymentGenerateService.generate("结算");
        exitApplication(ctx);
    }
    public static void exitApplication(ConfigurableApplicationContext context) {
        int exitCode = SpringApplication.exit(context, (ExitCodeGenerator) () -> 0);

        System.exit(exitCode);
    }
}
