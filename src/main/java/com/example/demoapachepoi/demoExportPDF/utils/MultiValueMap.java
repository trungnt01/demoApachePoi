package com.example.demoapachepoi.demoExportPDF.utils;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class MultiValueMap<K, V> {
    private final Map<K, List<V>> map = new LinkedHashMap<>();

    public void put(K key, V value) {
        map.computeIfAbsent(key, k -> new ArrayList<>()).add(value);
    }

    public List<V> get(K key) {
        return map.getOrDefault(key, new ArrayList<>());
    }

    public boolean containsKey(K key) {
        return map.containsKey(key);
    }

    public boolean containsValue(K key, V value) {
        List<V> values = map.get(key);
        return values != null && values.contains(value);
    }

    // Các phương thức khác tùy theo nhu cầu của bạn
}

