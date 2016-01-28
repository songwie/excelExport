package com.xr.export.BP;

import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.columnEntity.TableHeadInfo;
import com.xr.export.interfaces.IExport;

import net.sf.json.JSONObject;

@Component
@Scope("prototype")
public class ExportProxy {

	@Autowired
	private IExport clsExport;	

	private String fileAddress = "";

	private List<ColumnInfo> listColumnInfo;
	private List<TableHeadInfo> listHeadInfo;
	private List<Object[]> data;
	private String title = "";
	
	public String export( ArrayList<ColumnInfo>  listColumnInfo,List<Object[]> data,String title){
		JSONObject jsonObject = new JSONObject();
		jsonObject.put("success", true);
		jsonObject.put("error", "");
		try{
			this.listColumnInfo = listColumnInfo;
			//this.listHeadInfo = listHeadInfo;
			this.data = data;
			this.title = title;
			 
			this.fileAddress = clsExport.export(this.listColumnInfo, this.listHeadInfo, this.data,this.title);	
			jsonObject.put("fileAddress", fileAddress);
		}catch(Exception e){
			jsonObject.put("success", false);
			jsonObject.put("error", e.getMessage());

		}
		return jsonObject.toString(); 
	}

	public String export(List<Object[]> data, String title) {
		// TODO Auto-generated method stub
		return null;
	}

	
}
