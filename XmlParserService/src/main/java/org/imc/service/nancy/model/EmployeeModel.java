package org.imc.service.nancy.model;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
@Component
@Slf4j
@Data
public class EmployeeModel {
   private Map<String,NovalModel> novalModelMap = new LinkedHashMap<>();
   private Double totalAccount = 0.0;
}