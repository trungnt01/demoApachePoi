package com.example.demoapachepoi.demoExportPDF.utils;

import jakarta.persistence.EntityManager;
import jakarta.persistence.Query;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

@Service
public class SqlQueryUtil {

    @Autowired
    EntityManager entityManager;

    // return string
    public String getDataBySqlString(String sql, Map<String, String> paramList){
        Query query = entityManager.createNativeQuery(sql);
        if(paramList != null && !paramList.isEmpty()){
            paramList.forEach((key, value) ->
                query.setParameter(key, value)
            );
        }
        return (String) query.getSingleResult();
    }

    // return list
    public List<Object[]> getListDataBySqlString(String sql, Map<String, String> paramList){
        Query query = entityManager.createNativeQuery(sql);
        if(paramList != null && !paramList.isEmpty()){
            paramList.forEach((key, value) ->
                query.setParameter(key, value)
            );
        }
        return query.getResultList();
    }

}
