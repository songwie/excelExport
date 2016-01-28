package com.xr.export.typeconversion;

import java.util.ArrayList;

import org.springframework.core.convert.converter.Converter;
import org.springframework.stereotype.Component;

import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.common.MyConverter;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
/**
 *
 * @author sw
 *
 */

@SuppressWarnings("rawtypes")
@MyConverter
@Component
public class ColumnInfoTypeConversion implements Converter<String, ArrayList<ColumnInfo>  >{



	//int index,String colName,int colWidth,short colDataType,short colAlign,boolean colHidden
	@Override
	public  ArrayList<ColumnInfo>  convert(String source) {
		return ColumnInfoTypeConversion.parseCols(source);
	}

	public static ArrayList<ColumnInfo> parseCols(String source){
		if(source==null || source.equals("")){
			 return null;
		 }
		 JSONArray array = JSONArray.fromObject(source);
		 ArrayList<ColumnInfo> list = new ArrayList<ColumnInfo>();



		 for(int i=0;i<array.size();i++){
			 JSONObject cols = array.getJSONObject(i);

			 int index = cols.getInt("index");
			 String colName = cols.getString("colName");
			 String colData = cols.getString("colData");

			 String colWidths = cols.getString("colWidth");
			 colWidths = colWidths.replace("px", "");
			 if(colWidths.indexOf(".")!=-1){
				 colWidths = colWidths.substring(0,colWidths.indexOf("."));
			 }
			 int colWidth = Integer.valueOf( colWidths);
			 String colDataType = cols.getString("colDataType");
			 short dtype = ColumnInfo.TYPE_STRING;
			 if(colDataType.equals("string")){
				 dtype = ColumnInfo.TYPE_STRING;
			 }
			 if(colDataType.equals("number")){
				 dtype = ColumnInfo.TYPE_INT;
			 }
			 if(colDataType.equals("double")){
				 dtype = ColumnInfo.TYPE_DOUBLE;
			 }
			 if(colDataType.equals("boolean")){
				 dtype = ColumnInfo.TYPE_BOOLEAN;
			 }
			 if(colDataType.equals("percent")){
				 dtype = ColumnInfo.TYPE_PERCENT;
			 }
			 String colAlign = cols.getString("colAlign");
			 short align = ColumnInfo.LEFT;
			 if(colAlign.equals("CENTER")){
				 align = ColumnInfo.CENTER;
			 }
			 if(colAlign.equals("LEFT")){
				 align = ColumnInfo.LEFT;
			 }
			 if(colAlign.equals("RIGHT")){
				 align = ColumnInfo.RIGHT;
			 }
			 boolean colHidden = cols.getBoolean("colHidden");

			 ColumnInfo columnInfo = new ColumnInfo(index,colName,colData, colWidth , dtype,align,colHidden);

			 list.add(columnInfo);
		 }


		 return list;
	}
}
