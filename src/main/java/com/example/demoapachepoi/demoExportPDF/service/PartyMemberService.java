package com.example.demoapachepoi.demoExportPDF.service;

import com.example.demoapachepoi.demoExportPDF.entity.PartyMember;
import com.example.demoapachepoi.demoExportPDF.repository.PartyMemberRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class PartyMemberService {

    @Autowired
    private PartyMemberRepository partyMemberRepository;

    public List<PartyMember> getUserList(){
        return partyMemberRepository.findAll();
    }

}
