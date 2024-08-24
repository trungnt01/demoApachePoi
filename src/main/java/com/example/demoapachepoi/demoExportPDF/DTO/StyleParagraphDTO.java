package com.example.demoapachepoi.demoExportPDF.DTO;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.LinkedHashMap;
import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class StyleParagraphDTO {
    private String title;
    private String styleParagraph;
    private String stylesXML;
    private List<StyleRunDTO> styles;
}
