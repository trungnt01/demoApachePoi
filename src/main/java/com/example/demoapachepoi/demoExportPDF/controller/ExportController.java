package com.example.demoapachepoi.demoExportPDF.controller;

import com.example.demoapachepoi.demoExportPDF.service.DecisionService;
import com.example.demoapachepoi.demoExportPDF.service.ExportEntityService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping(value = "/export")
public class ExportController {


    @Autowired
    private DecisionService decisionService;

    @Autowired
    private ExportEntityService exportEntityService;

    @GetMapping()
    public ResponseEntity exportData(){
        decisionService.export();
        return ResponseEntity.ok(null);
    }

}
