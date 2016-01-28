package com.xr.export.common;


import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.AfterReturning;
import org.aspectj.lang.annotation.Aspect;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import com.xr.export.BP.ExportProxy;
import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.typeconversion.ColumnInfoTypeConversion;

@Aspect
public class ExportAspect {

	@SuppressWarnings("unchecked")
	//public Object  afterReturning( ProceedingJoinPoint joinPoint ) throws Throwable {
	//public Object  afterReturning( ProceedingJoinPoint joinPoint ) throws Throwable {

	//@Around("execution(* com.xr..*(..))")
	@AfterReturning(value="@annotation(ExportAnnotation)" ,returning="map" )
	public void afterReturning( JoinPoint joinPoint,Map  map ) throws Throwable {
        //Map  map = null;
		ExportProxy export = null;
		try {
			export = SpringContextUtils.getSpringContext().getBean(ExportProxy.class);
			
			HttpServletRequest request = (HttpServletRequest) ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getRequest(); 
			//HttpServletResponse res = ((ServletWebRequest)RequestContextHolder.getRequestAttributes()).getResponse();
			HttpServletResponse res = (HttpServletResponse) joinPoint.getArgs()[1]; 

			//map = (Map) joinPoint.proceed();
			
			String source = request.getParameter("columnInfo");
			ArrayList<ColumnInfo> columnInfos = ColumnInfoTypeConversion.parseCols(source);
			
			List<Object[]> data = (List<Object[]>) map.get("data");
			String title = (String) map.get("title");
			
			String rs = export.export(columnInfos, data, title);
 			
	        res.reset();  
	        res.setHeader("Content-disposition","attachment");
	        res.setContentType("application/msexcel; charset=utf-8");  
	        //res.getWriter().print(rs);
	        
			OutputStream os = res.getOutputStream();  
		    try {  
		        os.write(rs.getBytes());  
		        os.flush();  
		    } finally {  
		        if (os != null) {  
		            os.close();  
		        }  
		    }  
		}catch (Exception e) {
			e.printStackTrace();
		}
		 

        //return map;
	}
	 
	
}
