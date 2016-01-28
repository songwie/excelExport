package com.xr.export.columnEntity;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Component;

/**
 * @author sw
 */

@Component
public class DataBlockInfo {
	/**
	 * 是否添加合计行，默认为true
	 */
	public boolean TOTALLINE = true;
	
	/**
	 * 是否添加列序号，默认为false
	 */
	public boolean ROWNUMBER = true;
	
	/** 
	 * 数据块行高
	 */
	public short DataRowHeight = 15;
	
	private HSSFFont FONT = null;
	/**
	 * 数据块字体
	 */
	public HSSFFont getColumnFont(HSSFWorkbook wb){
		if(FONT == null){
			FONT = wb.createFont() ;
			FONT.setFontHeightInPoints((short)10);
			FONT.setBoldweight(org.apache.poi.hssf.usermodel.HSSFFont.BOLDWEIGHT_NORMAL);
			FONT.setFontName("宋体");
			FONT.setItalic(false);
			FONT.setStrikeout(false);
		}		
	    return FONT;
	}
	
	/**
	 * 得到主体数据样式,包含默认字体
	 */
	public HSSFCellStyle getListCellStyle(HSSFWorkbook wb)
	{
		HSSFCellStyle listCS = wb.createCellStyle();
		listCS.setFont(this.getColumnFont(wb));	
		return listCS;
	}
}
