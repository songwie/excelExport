package com.xr.export.action;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.xr.export.common.ExportAnnotation;

@RequestMapping("/excel2")
@Controller 
@Scope("prototype")
public class ExcelTest2 {

	 
	 
	@SuppressWarnings({ "rawtypes", "unchecked" })
	@RequestMapping("/export") 
	@ExportAnnotation
	@ResponseBody
	public Map export(HttpServletRequest resq,HttpServletResponse resp ){ 
		Map<String,Object> map = new HashMap<String,Object>();
		List data = new ArrayList();
		for(int i=0;i<13;i++){
			data.add(new Object[]{"列1","列2","列3",800+i,"列5",10000+i,10000+i,100200+i,10020+i,"列10","列11","列12","列13"});
		}
		String title = "titletest";
		map.put("title", title);
		map.put("data", data);
		//map.put("response", response);

		return map;
	}; 
}
