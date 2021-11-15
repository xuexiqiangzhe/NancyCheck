package org.imc;

import com.alibaba.dubbo.spring.boot.annotation.EnableDubboConfiguration;
import org.imc.service.nancy.FilterFilesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ExitCodeGenerator;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.ComponentScan;

@EnableDubboConfiguration
@SpringBootApplication
@ComponentScan(basePackages = {"org.imc.*"})
public class XmlParserWebApplication {
    public static void main(String[] args) {

        ConfigurableApplicationContext ctx = SpringApplication.run(XmlParserWebApplication.class, args);
        FilterFilesService filterFilesService = new FilterFilesService();
        String path = "过滤";
        filterFilesService.filter(path);
        exitApplication(ctx);
    }
    public static void exitApplication(ConfigurableApplicationContext context) {
        int exitCode = SpringApplication.exit(context, (ExitCodeGenerator) () -> 0);

        System.exit(exitCode);
    }
}
