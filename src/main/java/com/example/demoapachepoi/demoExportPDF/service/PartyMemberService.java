package com.example.demoapachepoi.demoExportPDF.service;

import com.example.demoapachepoi.demoExportPDF.entity.PartyMember;
import com.example.demoapachepoi.demoExportPDF.repository.PartyMemberRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class PartyMemberService {

    @Autowired
    private PartyMemberRepository partyMemberRepository;

    public List<PartyMember> getUserList(){
        return partyMemberRepository.findAll();
    }

//    @Scheduled(cron = "${cron.job}")
    public void cron(){
        System.out.println("job");
    }

}
