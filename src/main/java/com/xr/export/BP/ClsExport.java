package com.xr.export.BP;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import com.xr.export.columnEntity.ColumnInfo;
import com.xr.export.columnEntity.DataBlockInfo;
import com.xr.export.columnEntity.TableHeadInfo;
import com.xr.export.common.ExcelCommon;
import com.xr.export.interfaces.IExport;


/**
 * @author sw
 * @version 1.0.1b
 */

@Component
@Scope("prototype")
public class ClsExport implements IExport {

	private static Logger log = LoggerFactory.getLogger(ClsExport.class);

	/**
	 * 错误判断
	 */
	@Autowired
	private ExportInvalid invalid;

	/**
	 * 数据块信息
	 */
	@Autowired
	private DataBlockInfo DATABLOCKINFO ;

	/**
	 * Exce文档
	 */
	private HSSFWorkbook WORKBOOK = null;

	/**
	 * 结果集
	 */
	private List<Object[]> RESULTLIST = null;

	/**
	 * 结果集2
	 */
	private List<HashMap<String, Object>> RESULTLIST2 = null;

	/**
	 * 列及列头信息
	 */
	private List<ColumnInfo> COLUMNINFOLIST = null;

	/**
	 * 单页最大行，默认65536
	 */
	private int MAXROW = 65536;

	/**
	 * 当前行号
	 */
	private  int CUR_ROW = -1;

	/**
	 * 当前页号
	 */
	private  int CUR_SHEET = 0;

	/**
	 * 导出文件名
	 */
	private String FILE_FIRSTNAME = "";


	private String filePathOfSave = "";

	/**
	 * 导出文件名
	 */
	private  String FILE_PATH = "";

	/**
	 * 垂直比例数
	 */
	private static final short VRATIO = 20;

	/**
	 * 水平比例数
	 */
	private static final short HRATIO = 40;

	/**
	 * titleName
	 */
	private String titleName="";

	/**
	 * 返回文件流
	 */
	private InputStream inputStream;


	public String export(List<ColumnInfo> listColumnInfo,List<TableHeadInfo> listTableHeadInfo,List data,String title) {
		if(log.isDebugEnabled()){
			log.debug("xuri excel export start...");
		}
		if(log.isDebugEnabled()){
			log.debug("excel : listColumnInfo:" + listColumnInfo);
		}
		this.setColumInfoList(listColumnInfo);
		try {
			this.setTitleName(new String(title.getBytes("UTF-8"), "UTF-8"));
			//this.setTitleName(title);
		} catch (Exception e) {
            this.setTitleName("file");
     	}

		if(listTableHeadInfo!=null){
			//this.setListTableHeadInfo(listTableHeadInfo);
		}
		/*Object[] objects = data.get(0);
		if(data!=null&&data.size()>0){
			if(objects[0] instanceof HashMap<?, ?>){
				this.setResultList2(data);
			}else {
				this.setResultList(data);
			}
		}*/

		if(data!=null&&data.size()>0){
			if(data.get(0) instanceof HashMap<?, ?>){
				this.setResultList2(data);
			}else {
				this.setResultList(data);
			}
		}

		this.setFILE_FIRSTNAME( title);
		String[] SheetName = { title};
		String info = this.Export();

		if(info!=null){
			throw new RuntimeException(info);
		}

		String fileAddress = this.getFilePathOfSave();
		log.info("excel fileAddress :" + fileAddress);

		return fileAddress;
	}

	private String Export() {
		int iRet = 0;
		iRet = this.initialize();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		iRet = this.drawTitle();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		iRet = this.drawHead();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
	    }
		iRet = this.setColumnCellStyle();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		iRet = this.drawMainDataBlock();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		iRet = this.drawFooter();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		iRet = this.drawBorder();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		iRet = this.writeToFile();
		if(iRet <0 ){
			return invalid.Invalid(iRet);
		}
		return null;
	}


	public int drawTitle() {
		try {
			//增加一行
			HSSFRow row = this.addRow();

			//创建风格
			HSSFCellStyle cellStyle = this.getCurrWorkBook().createCellStyle();

			//设置样式,居中
			cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

			//设置样式,居中
			cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

			//设置字体
			HSSFFont font = this.getCurrWorkBook().createFont();
			font.setFontHeightInPoints((short) 12);
			font.setFontName("宋体");
			font.setBoldweight(org.apache.poi.hssf.usermodel.HSSFFont.BOLDWEIGHT_BOLD);

			//关联到style
			cellStyle.setFont(font);
			cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);

			//创建单元格并设置风格
			for (int i = 0; i < this.getColumnCount(); i++) {
				HSSFCell curCell = row.createCell(i);
				curCell.setCellStyle(cellStyle);
			}

			//写入信息
			row.getCell(0).setCellValue(new HSSFRichTextString(titleName));

			//合并
			this.getCurrSheet().addMergedRegion(new CellRangeAddress(row.getRowNum(),
																	 row.getRowNum(),
																	 0,
																	 this.getColumInfoList().size())
												);

		} catch (Exception e) {
			e.printStackTrace();
			return ExportInvalid.ERR_DRAWTITLE;
		}
		return 0;
	}

	public int drawHead() {
		HSSFRow row = addRow();
		if(row == null)
			return ExportInvalid.ERR_DRAWHEAD;
		boolean hasSetHeight = false;
		try{
			if(this.DATABLOCKINFO.ROWNUMBER){
				HSSFCell cell = row.createCell(0);
				HSSFCellStyle cs = WORKBOOK.createCellStyle();
				cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				cs.setHidden(false);
				cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
				cs.setFont(this.DATABLOCKINFO.getColumnFont(WORKBOOK));
				cell.setCellValue(new HSSFRichTextString("序号"));
				cell.setCellStyle(cs);
			}
			for(Iterator<ColumnInfo> it = COLUMNINFOLIST.iterator();it.hasNext();){
				ColumnInfo info = it.next();
				HSSFCell cell;
				if(this.DATABLOCKINFO.ROWNUMBER){
					cell = row.createCell(info.index+1);
					WORKBOOK.getSheetAt(CUR_SHEET).setColumnWidth(info.index+1, info.colWidth*HRATIO);
				} else {
					cell = row.createCell(info.index);
					WORKBOOK.getSheetAt(CUR_SHEET).setColumnWidth(info.index, info.colWidth*HRATIO);
				}

				cell.setCellValue(new HSSFRichTextString(info.colName.
						replaceAll("<br>", "")));
				/**这里额外加入对列名的处理，去掉列中的br标签*/

				HSSFCellStyle cs = WORKBOOK.createCellStyle();
				cs.setAlignment(info.headAlign);
				cs.setDataFormat(info.headStringDataFomate);
				cs.setHidden(info.colHidden);
				cs.setFont(info.getHeadFont(WORKBOOK));
				cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
				if(!hasSetHeight){
					row.setHeight((short)(info.headRowHeight*ClsExport.VRATIO));
					hasSetHeight = true;
				}
				cell.setCellStyle(cs);
			}
		} catch(Exception e){
			e.printStackTrace();
			return ExportInvalid.ERR_DRAWHEAD;
		}
		return 0;
	}
	@SuppressWarnings("unchecked")
	public int drawMainDataBlock() {
		try{
			int columnNum = 1;

			 //序列号
			HSSFCellStyle cs = WORKBOOK.createCellStyle();
			cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			cs.setHidden(false);
			cs.setFont(this.DATABLOCKINFO.getColumnFont(WORKBOOK));

			//根据列号
			if(this.RESULTLIST!=null&&this.RESULTLIST.size()>0){
				for(Iterator<Object[]> itRow = this.RESULTLIST.iterator();itRow.hasNext();){
					HSSFRow row = addRow();

					if(this.DATABLOCKINFO.ROWNUMBER){
						HSSFCell cell = row.createCell(0);
						cell.setCellValue(columnNum++);
						cell.setCellStyle(cs);
					}

					row.setHeight((short)(this.DATABLOCKINFO.DataRowHeight*VRATIO));
					Object[] objList = itRow.next();

					for(int col = 0;col <this.COLUMNINFOLIST.size();col++){
						ColumnInfo info = getInfoAt(col);
						Object obj = objList[col];

						setDataCell(row,obj,info);
					}
				}
			}else {
				//根据列索引
				for(Iterator<HashMap<String, Object>> itRow = this.RESULTLIST2.iterator();itRow.hasNext();){
					HSSFRow row = addRow();

					if(this.DATABLOCKINFO.ROWNUMBER){
						HSSFCell cell = row.createCell(0);
						cell.setCellValue(columnNum++);
						cell.setCellStyle(cs);
					}

					row.setHeight((short)(this.DATABLOCKINFO.DataRowHeight*VRATIO));
					HashMap<String, Object> keys = itRow.next();

					for(int col = 0;col <this.COLUMNINFOLIST.size();col++){
						ColumnInfo info = getInfoAt(col);
						Object obj =  "";

					    if(keys.containsKey(info.colData)){
							obj = keys.get(info.colData);

							setDataCell(row,obj,info);
							//break;
					    }

					}
				}
			}

			if(DATABLOCKINFO.TOTALLINE)
				addTotalLine();

		}catch(Exception e){
			e.printStackTrace();
			return ExportInvalid.ERR_DRAWMAINDATABLOCK;
		}
		return 0;
	}

	public int drawFooter() {
		return 0;
	}


	/**
	 * 创建每列的样式
	 */
	private int setColumnCellStyle()
	{
		try
		{
			for(int i=0;i<this.COLUMNINFOLIST.size();i++)
			{
				HSSFCellStyle cs = this.WORKBOOK.createCellStyle();
				cs.setFont(this.DATABLOCKINFO.getColumnFont(WORKBOOK));
				this.COLUMNINFOLIST.get(i).setCellStyle(cs);
			}
		}
		catch(Exception e)
		{
			e.printStackTrace();
			return ExportInvalid.ERR_CREATECOLUMNSYTLE;
		}
		return 0;
	}
	/**
	 * 画边框
	 * @return 成功返回非负
	 */
	protected int drawBorder()
	{
		try{
		for(int sheet =0;sheet < WORKBOOK.getNumberOfSheets();sheet++)
		{
			for(int row= WORKBOOK.getSheetAt(sheet).getFirstRowNum();row <= WORKBOOK.getSheetAt(sheet).getLastRowNum();row++)
			{

				if(null == WORKBOOK.getSheetAt(sheet).getRow(row))
					continue;

				for(int cell=0;cell < getColumnCount();cell++)
				{

					if(null == WORKBOOK.getSheetAt(sheet).getRow(row).getCell(cell)){
						WORKBOOK.getSheetAt(sheet).getRow(row).createCell(cell);
						HSSFCellStyle cs = WORKBOOK.createCellStyle();
						cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
						cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
						cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
						cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
						WORKBOOK.getSheetAt(sheet).getRow(row).getCell(cell).setCellStyle(cs);
						continue;
					}
					if(null == WORKBOOK.getSheetAt(sheet).getRow(row).getCell(cell).getCellStyle())	{
						HSSFCellStyle cs = WORKBOOK.createCellStyle();
						cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
						cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
						cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
						cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
						WORKBOOK.getSheetAt(sheet).getRow(row).getCell(cell).setCellStyle(cs);
						continue;
					}
					HSSFCellStyle cs = WORKBOOK.getSheetAt(sheet).getRow(row).getCell(cell).getCellStyle();
					cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
					cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
					WORKBOOK.getSheetAt(sheet).getRow(row).getCell(cell).setCellStyle(cs);
				}
			}
		}


		}catch(Exception e)
		{
			e.printStackTrace();
			return ExportInvalid.ERR_DRAWBORDER;
		}
		return 0;
	}

	/**
	 * 初始化Excel文档
	 * @return 成功返回非负值
	 */
	private int initialize() {
		try {
			if(this.getColumInfoList().size() == 0)
				return ExportInvalid.ERR_NOCOLUMNINFO;
			this.WORKBOOK = new HSSFWorkbook();
			this.WORKBOOK.createSheet();
			this.CUR_ROW = -1;
			this.CUR_SHEET = 0;
		} catch (Exception e) {
			e.printStackTrace();
			return ExportInvalid.ERR_INITIALIZE;
		}
		return 0;
	}

	/**
	 * 生成文件
	 * @return 成功返回非负值
	 */
	private int writeToFile(){
		try {
			String strGuid = UUID.randomUUID().toString();
			String filePath = "";
			if(System.getProperty("os.name").indexOf("Windows")!=-1)
			{
				filePath = ExcelCommon.EXPORT_WINDOW;
			}
			else
			{
				filePath = ExcelCommon.EXPORT_LINUX;
			}
			if(log.isDebugEnabled()){
				log.debug("excel filePath :" + filePath);
			}

			this.setFILE_PATH(filePath);
			String strFile =  this.getFILE_PATH() +File.separator+this.FILE_FIRSTNAME+"$"+strGuid+".xls";
			this.setFilePathOfSave(strGuid);
			File file = new File(this.getFILE_PATH());
			file.mkdirs();
			FileOutputStream fs = new FileOutputStream(strFile);
			WORKBOOK.write(fs);
			fs.close();
//			InputStream inputStream = new FileInputStream(strFile);
//			FileUtils.deleteQuietly(new File(strFile));
//			this.setInputStream(inputStream);
		} catch (Exception e) {
			e.printStackTrace();
			return ExportInvalid.ERR_WRITEFILE;
		}
		return 0;
	}

	/**
	 * 增加合计行
	 */
	public void addTotalLine()
	{
		HSSFRow row = addRow();
		HSSFCell cell = row.createCell(0);
		HSSFCellStyle cs = WORKBOOK.createCellStyle();
		cs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		cs.setHidden(false);
		HSSFFont font = WORKBOOK.createFont();
		font.setStrikeout(false);
		font.setBoldweight(org.apache.poi.hssf.usermodel.HSSFFont.BOLDWEIGHT_BOLD);
		cs.setFont(font);
		cell.setCellValue(new HSSFRichTextString("合计："));
		cell.setCellStyle(cs);
		cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);



		for(int i = 0;i <COLUMNINFOLIST.size();i++){
			ColumnInfo info = COLUMNINFOLIST.get(i);
			if(this.DATABLOCKINFO.ROWNUMBER){
				cell = row.createCell(info.index+1);
			}else{
				cell = row.createCell(info.index);
			}
			if((info.colDateType == ColumnInfo.TYPE_DOUBLE)||(info.colDateType == ColumnInfo.TYPE_BIGDECIMAL)||info.colDateType == ColumnInfo.TYPE_INT)
			{
				cs = WORKBOOK.createCellStyle();
				cs.setAlignment(info.colAlign);
				cs.setHidden(false);
				cs.setFont(this.DATABLOCKINFO.getColumnFont(WORKBOOK));
				cs.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				cs.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);

				if(ColumnInfo.TYPE_INT == info.colDateType){
					cs.setDataFormat(info.colGeneralDataFomate);

				}else{
					cs.setDataFormat(info.colNumericDataFomate);
				}
				cell.setCellValue(info.sumValue);
			}
			cell.setCellStyle(cs);
		}

	}

	/**
	 * 设数据块单元值
	 * @param row HSSFRow 当前操作行
	 * @param obj 数据
	 * @param info ColumnInfo 列信息
	 */
	private void setDataCell(HSSFRow row,Object obj,ColumnInfo info){
		HSSFCell cell;
		if(this.DATABLOCKINFO.ROWNUMBER)
			cell = row.createCell(info.index+1);
		else
			cell = row.createCell(info.index);

		HSSFCellStyle cs = info.getCellStyle();
		cs.setAlignment(info.colAlign);
		cs.setHidden(info.colHidden);
		//cs.setFont(this.DATABLOCKINFO.getColumnFont(WORKBOOK));


		if(info.colDateType == ColumnInfo.TYPE_NONE || info.colDateType == ColumnInfo.TYPE_UNKNOW)
			info.colDateType = getObjectType(obj,info);

		String sVal;
		Date dtVal;
		double dVal;
		int iVal;
		boolean bVal;
		short shVal;

		switch(info.colDateType){
			case ColumnInfo.TYPE_STRING :
				sVal = (obj!=null?obj.toString():"");
				sVal = sVal.replaceAll("&nbsp;", " ");
				cell.setCellValue(new HSSFRichTextString(sVal));
				cs.setDataFormat(info.colGeneralDataFomate);
				break;
			case ColumnInfo.TYPE_BOOLEAN :
				bVal = (obj!=null?Boolean.valueOf(obj.toString()).booleanValue():false);
				cell.setCellValue(bVal);
				cs.setDataFormat(info.colGeneralDataFomate);
				break;
			case ColumnInfo.TYPE_DATE :
				dtVal = ((obj!=null&&!obj.equals(""))?(Date)obj:null);
				if(dtVal != null){
					cell.setCellValue(dtVal);
					cs.setDataFormat(info.colDateDataFomate);
				}
				else{
					cell.setCellValue(new HSSFRichTextString(""));
				}
				break;
			case ColumnInfo.TYPE_DOUBLE :
				dVal = (obj!=null?Double.parseDouble(obj.toString()):(double)0);
				cell.setCellValue(dVal);
				cs.setDataFormat(info.colNumericDataFomate);
				info.sumValue += dVal;
				break;
			case ColumnInfo.TYPE_BIGDECIMAL :
				dVal = (obj!=null?Double.parseDouble(((BigDecimal)(obj)).toString()):(double)0);
				cell.setCellValue(dVal);
				cs.setDataFormat(info.colNumericDataFomate);
				info.sumValue += dVal;
				break;
			case ColumnInfo.TYPE_PERCENT:
				dVal = (obj!=null?Double.parseDouble(obj.toString()):(double)0);
				cell.setCellValue(dVal/100);
				cs.setDataFormat(HSSFDataFormat.getBuiltinFormat(info.colPercentDataFomate));
				break;
			case ColumnInfo.TYPE_INT :
				iVal = (obj!=null?Integer.parseInt(obj.toString()):0);
				cell.setCellValue(iVal);
				cs.setDataFormat(info.colGeneralDataFomate);
				info.sumValue += iVal;
				break;
			case ColumnInfo.TYPE_SHORT :
				shVal = (obj!=null?Short.parseShort(obj.toString()):(short)0);
				cell.setCellValue(shVal);
				cs.setDataFormat(info.colGeneralDataFomate);
				info.sumValue += shVal;
				break;
			case ColumnInfo.TYPE_NONE :
				sVal = (obj!=null?obj.toString():"");
				cell.setCellValue(new HSSFRichTextString(sVal));
				cs.setDataFormat(info.colGeneralDataFomate);
				break;
			case ColumnInfo.TYPE_UNKNOW :
				sVal = (obj!=null?obj.toString():"");
				cell.setCellValue(new HSSFRichTextString(sVal));
				cs.setDataFormat(info.colGeneralDataFomate);
				break;
		}
		cell.setCellStyle(cs);
	}

	/**
	 * 增加行
	 * @return HSSFRow
	 */
	protected HSSFRow addRow(){
		if(CUR_ROW == MAXROW-1){
			CUR_ROW = 0;
			CUR_SHEET++;
			WORKBOOK.createSheet();
		} else {
			CUR_ROW++;
		}
		return WORKBOOK.getSheetAt(CUR_SHEET).createRow(CUR_ROW);
	}

	/**
	 * 设置列类型
	 * @return
	 */
	public short  getObjectType(Object obj,ColumnInfo info)
	{
		if(obj == null)
		{
			return ColumnInfo.TYPE_NONE;
		}
		if(obj instanceof java.lang.Integer)
		{
			return ColumnInfo.TYPE_INT;
		}
		if(obj.getClass().equals(java.math.BigInteger.class))
		{
			return ColumnInfo.TYPE_INT;
		}
		//百分比-double
		else if((obj instanceof java.lang.Double)&&(info.colName.indexOf("%")!=-1)&&(info.colName.indexOf("%")==info.colName.lastIndexOf("%")))
		{
			return ColumnInfo.TYPE_PERCENT;
		}
		//百分比-BigDecimal
		else if((obj instanceof java.math.BigDecimal)&&(info.colName.indexOf("%")!=-1)&&(info.colName.indexOf("%")==info.colName.lastIndexOf("%")))
		{
			return ColumnInfo.TYPE_PERCENT;
		}
		else if(obj instanceof java.lang.Double)
		{
			return ColumnInfo.TYPE_DOUBLE;
		}
		else if(obj instanceof java.math.BigDecimal)
		{
			return ColumnInfo.TYPE_BIGDECIMAL;
		}
		else if(obj instanceof java.util.Date)
		{
			return ColumnInfo.TYPE_DATE;
		}
		else if(obj instanceof java.lang.Boolean)
		{
			return ColumnInfo.TYPE_BOOLEAN;
		}
		else if(obj instanceof java.lang.String)
		{
			return ColumnInfo.TYPE_STRING;
		}
		else if(obj instanceof java.lang.Short)
		{
			return ColumnInfo.TYPE_SHORT;
		}
		else
		{
			return ColumnInfo.TYPE_UNKNOW;
		}

	}

	/**
	 * 得到相应index的列信息
	 * @param index
	 * @return ColumnInfo
	 */
	private ColumnInfo getInfoAt(int index){
		ColumnInfo info = null;
		for(Iterator<ColumnInfo> it = COLUMNINFOLIST.iterator();it.hasNext();)
		{
			info = it.next();
			if(info.index == index)
				return info;
		}
		return null;
	}

	/**
	 * 得到当前SHEET
	 *
	 */
	protected HSSFSheet getCurrSheet()
	{
		return WORKBOOK.getSheetAt(CUR_SHEET);
	}

	/**
	 * 得到当前HSSFWorkbook
	 *
	 */
	protected HSSFWorkbook getCurrWorkBook()
	{
		return WORKBOOK;
	}
	/**
	 * 得到当前行
	 *
	 */
	protected HSSFRow getCurrRow()
	{
		return WORKBOOK.getSheetAt(CUR_SHEET).getRow(this.CUR_ROW);
	}
	/**
	 * @return List<Object[]> 结果集
	 */
	public List<Object[]> getResultList() {
		return RESULTLIST;
	}

	/**
	 * 设定结果集
	 * @param resultList
	 */
	public void setResultList(List<Object[]> resultList) {
		this.RESULTLIST = resultList;
	}
	/**
	 * 设定结果集
	 * @param resultList
	 */
	public void setResultList2(List resultList2) {
		this.RESULTLIST2 = resultList2;
	}
	/**
	 * 列及列头信息
	 */
	public List<ColumnInfo> getColumInfoList() {
		return COLUMNINFOLIST;
	}

	/**
	 * 列及列头信息
	 */
	public void setColumInfoList(List<ColumnInfo> columInfoList) {
		this.COLUMNINFOLIST = columInfoList;
	}

	/**
	 * 单页最大行，默认65536
	 */
	public int getMaxRow() {
		return MAXROW;
	}

	/**
	 * 单页最大行，默认65536
	 */
	public void setMaxRow(int maxRow) {
		this.MAXROW = maxRow;
	}

	/**
	 * 数据块信息
	 */
	public DataBlockInfo getDataBolckInfo() {
		return DATABLOCKINFO;
	}

	/**
	 * 数据块信息
	 */
	public void setDataBolckInfo(DataBlockInfo dataBolckInfo) {
		DATABLOCKINFO = dataBolckInfo;
	}

	/**
	 * 导出文件名
	 */
	public String getFILE_FIRSTNAME() {
		return FILE_FIRSTNAME;
	}

	/**
	 * 导出文件名
	 */
	public void setFILE_FIRSTNAME(String file_firstname) {
		FILE_FIRSTNAME = file_firstname;
	}

	/**
	 * 得到文件路径
	 */
	public String getFILE_PATH() {
		return FILE_PATH;
	}

	/**
	 * 设置文件路径
	 * @param file_path
	 */
	public void setFILE_PATH(String file_path) {
		FILE_PATH = file_path;
	}

	public String getFilePathOfSave() {
		return filePathOfSave;
	}

	public void setFilePathOfSave(String filePathOfSave) {
		this.filePathOfSave = filePathOfSave;
	}


	/**
	 * 得到列数
	 */
	protected int getColumnCount(){
		return (int)(this.DATABLOCKINFO.ROWNUMBER == true?(this.COLUMNINFOLIST.size()+1):(this.COLUMNINFOLIST.size()));
	}

	public String getTitleName() {
		return titleName;
	}

	public void setTitleName(String titleName) {
		this.titleName = titleName;
	}

	public InputStream getInputStream() {
		return inputStream;
	}

	public void setInputStream(InputStream inputStream) {
		this.inputStream = inputStream;
	}




}
