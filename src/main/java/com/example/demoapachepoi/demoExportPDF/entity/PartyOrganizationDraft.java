package com.example.demoapachepoi.demoExportPDF.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class PartyOrganizationDraft {
    private Long partyOrganizationDraftId;
    private String code;
    private String name;
    private Long parentId;
    private Long type;
    private Date effectiveDate;
    private String aliasName;
}