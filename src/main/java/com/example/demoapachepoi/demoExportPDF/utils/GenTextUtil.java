package com.example.demoapachepoi.demoExportPDF.utils;

import com.example.demoapachepoi.demoExportPDF.entity.PartyMemberDraft;
import com.example.demoapachepoi.demoExportPDF.entity.PartyOrganizationDraft;
import org.springframework.stereotype.Service;

import java.util.HashSet;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class GenTextUtil {


    /**
     * comment: gen bien ${sapNhap}
     * @param createType
     * @param list
     * @return: tra ve string bien ${sapNhap}
     */
    public String genCreateType(int createType, List<String> list){
        StringBuilder sapNhap = new StringBuilder();
        StringBuilder text = new StringBuilder("trên cơ sở sáp nhập ");
        if(createType == 1){
            sapNhap.append(text);
            if(list.size() == 1){
                sapNhap.append(list.get(0));
            } else {
                for (int i = 0; i < list.size(); i++){
                    String string = list.get(i);
                    if(i == list.size() - 1){
                        sapNhap.append(" và ").append(string);
                    } else {
                        sapNhap.append(string).append(", ");
                    }
                }
            }
        }
        return sapNhap.toString();
    }


    public String genTextCacDieuTren(){
        return null;
    }


    public String genSoLuongBanPhatHanh(List<PartyOrganizationDraft> partyOrganizationDraftList, List<PartyMemberDraft> partyMemberDrafts){

        int soLuongBanPhatHanh = 7; // bao gom 3 theo mac dinh(co quan soan thao, VT, DU), 4(tcd RQD, DB cap tren, tcd duoc thanh lap, tcd cap tren truc tiep)
        // todo: dem so luong dang con, so luong dang bi sap nhap, so luong thanh vien co ten trong quyet dinh(trung nhau khong duoc tinh)

        if(!partyOrganizationDraftList.isEmpty()){
            // convert list to set -> remove object duplicate
            HashSet<PartyOrganizationDraft> partyOrganizationDrafts = new HashSet<>(partyOrganizationDraftList);
            soLuongBanPhatHanh += partyOrganizationDrafts.size();
        }

        if(!partyMemberDrafts.isEmpty()){
            // convert list to set -> remove object duplicate
            HashSet<PartyMemberDraft> partyMemberDraftHashSet = new HashSet<>(partyMemberDrafts);
            soLuongBanPhatHanh += partyMemberDraftHashSet.size();
        }
        return String.valueOf(soLuongBanPhatHanh);
    }

}
