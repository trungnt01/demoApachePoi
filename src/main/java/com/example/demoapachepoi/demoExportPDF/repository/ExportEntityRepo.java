package com.example.demoapachepoi.demoExportPDF.repository;


import com.example.demoapachepoi.demoExportPDF.entity.ExportEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ExportEntityRepo extends JpaRepository<ExportEntity, Integer> {
}
