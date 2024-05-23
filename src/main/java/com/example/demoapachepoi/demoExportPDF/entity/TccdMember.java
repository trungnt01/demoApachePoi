package com.example.demoapachepoi.demoExportPDF.entity;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

@Data
@AllArgsConstructor
public class TccdMember {
    private String TenTCCapDuoi;
    List<PartyMember> partyMembers;
}
