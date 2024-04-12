package com.example.demoapachepoi.demoExportPDF.repository;

import com.example.demoapachepoi.demoExportPDF.entity.Decision;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface DecisionRepo extends JpaRepository<Decision, Integer> {
}
