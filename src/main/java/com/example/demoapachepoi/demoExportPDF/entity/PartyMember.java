package com.example.demoapachepoi.demoExportPDF.entity;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Entity
@Table(name = "user")
public class PartyMember {
    @Id
    private Integer id;
    private String HoTen;
    private String QuanHam;
    private String ChucVuChinhQuyen;
    private String ChucVuCapUy;
    private String DiaChi;
}
