package com.example.demoapachepoi.demoExportPDF.entity;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Data;

@Entity
@Data
public class ExportEntity {

    @Id
    private Integer id;
    private String decisionName;
    private String orgDecisionName;

}
