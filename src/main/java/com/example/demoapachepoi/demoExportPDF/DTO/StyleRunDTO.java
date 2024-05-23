package com.example.demoapachepoi.demoExportPDF.DTO;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class StyleRunDTO {
    private String title;
    private String styles;
    private List<String> readOnly;
    private boolean bulletPoint = false;
}
