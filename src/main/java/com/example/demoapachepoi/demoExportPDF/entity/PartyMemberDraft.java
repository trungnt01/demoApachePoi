package com.example.demoapachepoi.demoExportPDF.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class PartyMemberDraft {

    private Long id;
    private Long partyMemberDraftId;
    private Long organizationId;
    private Long partyPositionId;

}
