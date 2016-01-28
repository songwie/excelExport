package com.xr.export.interfaces;

import java.util.List;

import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.columnEntity.TableHeadInfo;
 

/**
 * @author sw
 * @version 1.0.1 
 *	
 * 异常处理有可改进之处
 */
public interface IExport {	
	
	/**
	 * 
	 * 导出总方法
	 * 	 */
	String export(List<ColumnInfo> listColumnInfo,List<TableHeadInfo> listTableHeadInfo,List data,String title);
	
}
