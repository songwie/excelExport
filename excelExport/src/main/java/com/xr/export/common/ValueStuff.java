package com.xr.export.common;

import java.util.ArrayList;
import java.util.concurrent.ConcurrentHashMap;

import com.xr.export.columnEntity.ColumnInfo;

public class ValueStuff {

	private static ConcurrentHashMap<String,ArrayList<ColumnInfo>>  map = new ConcurrentHashMap<String,ArrayList<ColumnInfo>>();
	
	public static ArrayList<ColumnInfo>  getValue(String key){
		return map.get(key);
	}
	public static  void setValue(String key,ArrayList<ColumnInfo> value){
		map.put(key, value);
	}
}
