package com.example.demoapachepoi.demoExportPDF.repository;

import com.example.demoapachepoi.demoExportPDF.entity.DecisionParams;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface DecisionParamRepository extends JpaRepository<DecisionParams, Integer> {
}
