package org.imc.service.nancy.model;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;

@Component
@Slf4j
@Data
public class TranslatorModel {
   private Map<String,EmployeeModel> editorModelMap = new HashMap<>();
   private Map<String,EmployeeModel> translatorModelMap = new HashMap<>();
   private Map<String,EmployeeModel> qualityModelMap = new HashMap<>();

}