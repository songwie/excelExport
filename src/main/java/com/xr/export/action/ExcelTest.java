package com.xr.export.action;

import java.util.*;

import javax.servlet.http.HttpServletRequest;

import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.xr.export.BP.AbstractExport;
import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.common.ValueStuff;

@RequestMapping("/excel")
@Controller
@Scope("prototype")
public class ExcelTest extends AbstractExport{

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void export(HttpServletRequest request){
		List list = new ArrayList();
		/*for(int i=0;i<13;i++){
			list.add(new Object[]{"列1","列2","列3",800+i,"列5",10000+i,800+i,10000+i,"列9","列10","列11","列12","列13"});
		}*/
		for(int i=0;i<13;i++){
			Map<String,Object> map = new HashMap<String,Object>();
			map.put("branchSequence", "列1");
			map.put("tradeCode", "列1");
			map.put("tradeStatus", "列1");
			map.put("traderDate", "列1");
			map.put("merchantId", "列1");
			map.put("userInfo", "列1");
			map.put("totalFee", 11111);
			map.put("userFee", 1212122);
			map.put("userFee1", "列1");
			map.put("userFe23e", "列1");
			map.put("user232Fee", "列1") ;
			list.add(map);
		}
		this.setData( list, "testtitle");
	};

	/*@RequestMapping("/export")
	@CheckLoginAnnotation(isLogin=ISLOGIN.NO)
	@ExportAnnotation
	public void export(){
		Map<String,Object> map = new HashMap<String,Object>();
		List data = new ArrayList();
		for(int i=0;i<13;i++){
			data.add(new Object[]{"列1","列2","列3",800+i,"列5",10000+i,"列7","列8","列9","列10","列11","列12","列13"});
		}
		String title = "titletest";
		map.put("data", data);
		map.put("title", title);

	};*/
}
