package com.xr.export.columnEntity;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.springframework.stereotype.Component;

@Component
public class CellInfo {

	/**
	 * 单元格的字体样式
	 */
	private HSSFFont FONT ;
	
	/**
	 * 单元格对齐方式
	 */
	private short colAlign = org.apache.poi.hssf.usermodel.HSSFCellStyle.ALIGN_LEFT;

	/**
	 * 单元格的填充内容
	 */
	private String cellText = "";

	public void setFONT(HSSFFont fONT) {
		FONT = fONT;
	}

	public HSSFFont getFONT() {
		return FONT;
	}

	public void setColAlign(short colAlign) {
		this.colAlign = colAlign;
	}

	public short getColAlign() {
		return colAlign;
	}

	public void setCellText(String cellText) {
		this.cellText = cellText;
	}

	public String getCellText() {
		return cellText;
	}
	
	
}

