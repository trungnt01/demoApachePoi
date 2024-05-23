package com.example.demoapachepoi.demoExportPDF.DTO;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class DecisionParamDTO {
    private String key;
    private String value;
}
