package org.imc.dubbo.impl;

import org.imc.service.nancy.FilterFilesService;
import org.imc.dubbo.FilterFilesRemoteService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;


@Controller
@RequestMapping("/")
public class FilterFilesRemoteServiceImpl implements FilterFilesRemoteService {

    @Autowired
    private FilterFilesService filterFilesService;

    @Override
    @GetMapping("/filter")
    public void filter() {
        try {
            String path = "过滤";
            filterFilesService.filter(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}