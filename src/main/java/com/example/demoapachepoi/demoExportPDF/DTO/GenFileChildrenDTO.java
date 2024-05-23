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
public class GenFileChildrenDTO {

    private String title;
    private String styleParagraph;
    private List<String> readOnlyList;
    private boolean bulletPoint;
}
