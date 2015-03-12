package com.xr.export.action;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.Iterator;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.FileFileFilter;
import org.springframework.context.annotation.Scope;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.xr.export.common.ExcelCommon;

/**
 * @author sw
 */


@RequestMapping("/excel")
@Controller
@Scope("prototype")
public class ExcelDownload {

    private String inputPath;
    private String filename;
    private InputStream inputStream;
    private int bufferSize;
    
    
    public String getFilename(String filename) {
    	String fileName = "";
		try {
			fileName =  new String(filename.getBytes("GBK"), "ISO8859-1");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return fileName; 
	}

	public void setFilename(String filename) {
		this.filename = filename;
	}

	public int getBufferSize() {
		return bufferSize;
	}

	public void setBufferSize(int bufferSize) {
		this.bufferSize = bufferSize;
	}

	public String getInputPath() {
		return inputPath;
	}

	public void setInputPath(String inputPath) {
		this.inputPath = inputPath;
    }

	@RequestMapping("/down") 
	public ResponseEntity<byte[]> execute(HttpServletResponse response, 
			@RequestParam(value="filename", required=false) String filename) throws Exception {
		String filePath = "";

    	HttpHeaders headers = new HttpHeaders();  
    	headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);  
    	ResponseEntity<byte[]> responseEntity = null; 

		this.filename = this.getFilename(filename);
		
		if(System.getProperty("os.name").indexOf("Windows")!=-1)
		{
			filePath = ExcelCommon.EXPORT_WINDOW;
		}
		else
		{
			filePath = ExcelCommon.EXPORT_LINUX;
		}
    	Iterator itFiles = FileUtils.iterateFiles(new File(filePath),FileFileFilter.FILE , null);
    	while(itFiles.hasNext())
		{
    		File file = (File)itFiles.next();
			String fileName = file.getName();
			if(fileName.indexOf(this.filename)!=-1)
			{
				int index = fileName.indexOf("$");
				this.setFilename(fileName.substring(0,index)+".xls");
		    	headers.setContentDispositionFormData("attachment",fileName.substring(0,index)+".xls");  
		    	responseEntity = new ResponseEntity<byte[]>(FileUtils.readFileToByteArray(file),headers, HttpStatus.CREATED); 
				break;
			}
		}
    	
    	
    	return responseEntity;
    }
	
	@RequestMapping("/down1") 
	public void downExcel(HttpServletResponse res, 
			@RequestParam(value="filename", required=false) String filename) throws Exception {
		String filePath = "";


		this.filename = this.getFilename(filename);
		
		if(System.getProperty("os.name").indexOf("Windows")!=-1)
		{
			filePath = ExcelCommon.EXPORT_WINDOW;
		}
		else
		{
			filePath = ExcelCommon.EXPORT_LINUX;
		}
    	Iterator itFiles = FileUtils.iterateFiles(new File(filePath),FileFileFilter.FILE , null);
    	while(itFiles.hasNext())
		{
    		File file = (File)itFiles.next();
			String fileName = file.getName();
			if(fileName.indexOf(this.filename)!=-1)
			{
				int index = fileName.indexOf("$");
				this.setFilename(fileName.substring(0,index)+".xls");

				
				OutputStream os = res.getOutputStream();  
			    try {  
			        res.reset();  
			        res.setHeader("Content-Disposition", "attachment; filename="+fileName.substring(0,index)+".xls");  
			        res.setContentType("application/octet-stream; charset=utf-8");  
			        os.write(FileUtils.readFileToByteArray(file));  
			        os.flush();  
			    } finally {  
			        if (os != null) {  
			            os.close();  
			        }  
			    }  

				
				break;
			}
		}
    	
    	
    }

}
