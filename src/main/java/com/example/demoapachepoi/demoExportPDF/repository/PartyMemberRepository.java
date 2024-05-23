package com.example.demoapachepoi.demoExportPDF.repository;

import com.example.demoapachepoi.demoExportPDF.entity.PartyMember;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

@Repository
public interface PartyMemberRepository extends JpaRepository<PartyMember, Integer> {

    @Query(value = ":sql", nativeQuery = true)
    String findDataParamBySqlQuery(@Param("sql") String sql);

}
