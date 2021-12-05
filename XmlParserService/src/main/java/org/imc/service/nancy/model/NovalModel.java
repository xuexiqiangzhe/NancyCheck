package org.imc.service.nancy.model;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;
@Component
@Slf4j
@Data
public class NovalModel {
   private Map<String,ChapterModel> chapterModelMap = new HashMap<>();
   private Double chapterCount = 0.0;
   private Double wordCount = 0.0;
   private Double account = 0.00;
}