package com.example.demoapachepoi.demoExportPDF.utils;

import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

public class CustomLinkedHashMap<K, V> extends LinkedHashMap<K, List<V>> {

    @Override
    public List<V> put(K key, List<V> value) {
        List<V> existingValue = get(key);
        if (existingValue == null) {
            // Key chưa tồn tại, thêm mới
            return super.put(key, value);
        } else {
            // Key đã tồn tại, thêm vào danh sách hiện tại
            List<V> newValueList = new LinkedList<>(existingValue);
            newValueList.addAll(value);
            return super.put(key, newValueList);
        }
    }

}
