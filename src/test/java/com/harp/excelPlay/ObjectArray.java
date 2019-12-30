package com.harp.excelPlay;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class ObjectArray {

    public static void main(String[] args) {
        // Key is ID and Value is in form of array. The map is sorted according to the natural ordering of its keys
        Map<String , Object[]> data = new TreeMap <String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[] {1, "Amit", "Shukla"});
        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
        data.put("4", new Object[] {3, "John", "Adwards"});
        data.put("5", new Object[] {4, "Brian", "Schultz"});



    Set<String> keySet = data.keySet();
    for(String key:keySet){
        Object[] objArr = data.get(key);

    }
    }
}
