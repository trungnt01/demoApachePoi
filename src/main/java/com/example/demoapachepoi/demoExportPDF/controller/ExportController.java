package com.example.demoapachepoi.demoExportPDF.controller;

import com.example.demoapachepoi.demoExportPDF.DTO.GenFileDTO;
import com.example.demoapachepoi.demoExportPDF.service.DecisionService;
import com.example.demoapachepoi.demoExportPDF.service.ExportDecisionService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping(value = "/export")
public class ExportController {


    @Autowired
    private DecisionService decisionService;

    @Autowired
    private ExportDecisionService exportEntityService;

    @GetMapping()
    public ResponseEntity exportData(){
        return ResponseEntity.ok(decisionService.exportText());
    }

    @GetMapping("/export-style")
    public ResponseEntity exportStyle(){
        return ResponseEntity.ok(decisionService.exportStyle());
    }

    @GetMapping("/export-decision")
    public ResponseEntity exportDecision(){
        return ResponseEntity.ok(exportEntityService.exportStyle());
    }

    @GetMapping("/export-decision-btv")
    public ResponseEntity exportDecisionBtv(){
        return ResponseEntity.ok(exportEntityService.exportStyleBTV());
    }

    @PostMapping("/gen-file")
    public ResponseEntity genFile(@RequestBody Map<String, List<GenFileDTO>> stringListMap){
        exportEntityService.genFileAfterStep2(stringListMap);
        return ResponseEntity.ok(null);
    }

}
