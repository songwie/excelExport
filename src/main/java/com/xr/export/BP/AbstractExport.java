package com.xr.export.BP;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONObject;

import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.columnEntity.TableHeadInfo;
import com.xr.export.common.ValueStuff;
import com.xr.export.interfaces.IExport;
import com.xr.export.typeconversion.ColumnInfoTypeConversion;

public abstract class AbstractExport   {

	@Autowired
	IExport clsExport;

	@Autowired
	ExportProxy exportProxy;

	private String fileAddress = "";

	private static final Logger LOGGER = Logger.getLogger(AbstractExport.class);

	private List<ColumnInfo> listColumnInfo;
	private List<TableHeadInfo> listHeadInfo;
	private List data;
	private String title = "";

	@RequestMapping("/export")
	@ResponseBody
	public String export(Model model, HttpServletRequest request,
			@RequestParam(value="columnInfo", required=false)ArrayList<ColumnInfo>  listColumnInfo){
		JSONObject jsonObject = new JSONObject();
		jsonObject.put("success", true);
		jsonObject.put("error", "");
		try{
			this.listColumnInfo = listColumnInfo;
			this.listHeadInfo = listHeadInfo;

			this.export(request);

			this.fileAddress = clsExport.export(this.listColumnInfo, this.listHeadInfo, this.data,this.title);
			jsonObject.put("fileAddress", fileAddress);
		}catch(Exception e){
			LOGGER.error("export :",e);
			jsonObject.put("success", false);
			jsonObject.put("error", e.getMessage());

		}

		String key = String.valueOf(Thread.currentThread().getId());
		ArrayList<ColumnInfo> cols = ValueStuff.getValue(key);


		return jsonObject.toString();
	}

	@RequestMapping("/export2")
	@ResponseBody
	public String export2(Model model, HttpServletRequest request,HttpServletResponse response){
		JSONObject jsonObject = new JSONObject();
		jsonObject.put("success", true);
		jsonObject.put("error", "");
		try{
			String source = request.getParameter("columnInfo");
			//System.err.println(request.getParameterMap());
			ArrayList<ColumnInfo> columnInfos = ColumnInfoTypeConversion.parseCols(source);

			this.listColumnInfo = columnInfos;
			this.listHeadInfo = listHeadInfo;

			this.export(request);

			this.fileAddress = clsExport.export(this.listColumnInfo, this.listHeadInfo, this.data,this.title);
			jsonObject.put("fileAddress", fileAddress);
		}catch(Exception e){
			LOGGER.error("export :",e);
			jsonObject.put("success", false);
			jsonObject.put("error", e.getMessage());

		}

		String key = String.valueOf(Thread.currentThread().getId());
		ArrayList<ColumnInfo> cols = ValueStuff.getValue(key);


		return jsonObject.toString();
	}

	public String setData(List  data,String title){
		this.data = data;
		try {
			this.title = new String(title.getBytes("UTF-8"), "UTF-8");
		} catch (UnsupportedEncodingException e) {
			this.title = "excel";
		}
		return null;
	}
	public abstract void export(HttpServletRequest request);

}
