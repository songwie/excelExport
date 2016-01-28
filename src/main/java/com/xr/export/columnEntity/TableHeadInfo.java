package com.xr.export.columnEntity;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.springframework.stereotype.Component;

public class TableHeadInfo {

	/**
	 * 构造函数
	 * @param index:序号
	 * @param row:表头行
	 * @param col:表头列
	 * @param rowspan:单元格合并行数
	 * @param colspan:单元格合并列数
	 * @param colAlign:列对齐方式
	 * @param headName:单元格填充内容
	 */
	public TableHeadInfo(int index,int row,int col,int rowspan,int colspan,short colAlign,String headName )
	{
		this.index = index;
		this.row = row;
		this.col = col;
		this.rowspan = rowspan;
		this.colspan = colspan;
		this.colAlign = colAlign ;
		this.headName = headName ;
	}
	/**
	 * 列序号
	 */
	public int index = 0;

	/**
	 * 表头行
	 */
	public int row;
	
	/**
	 * 表头列
	 */
	public int col;
	
	/**
	 * 单元格合并行数
	 */
	public int rowspan;
	
	/**
	 * 单元格合并列数
	 */
	public int colspan;
	
	/**
	 * 单元格对齐方式
	 */
	public short colAlign = org.apache.poi.hssf.usermodel.HSSFCellStyle.ALIGN_LEFT;	
	
	/**
	 * 单元格填充内容
	 */
	public String headName = "undefined";

//==================================
	/**
	 * 每列的样式
	 */
	private HSSFCellStyle cellStyle;
	/**
	 * 每列的样式
	 */
	public HSSFCellStyle getCellStyle() {
		return cellStyle;
	}
	/**
	 * 每列的样式
	 */
	public void setCellStyle(HSSFCellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
}
