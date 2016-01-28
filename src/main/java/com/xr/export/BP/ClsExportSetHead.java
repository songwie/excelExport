package com.xr.export.BP;
/**
 * @author sw 
 */
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import com.xr.export.columnEntity.CellInfo;
import com.xr.export.columnEntity.TableHeadInfo;

@Component
@Scope("prototype")
public class ClsExportSetHead extends ClsExport {
	
	/**
	 * 表头信息
	 */
	private List<TableHeadInfo> listTableHeadInfo;
	/**
	 * 无底纹行的起始行
	 */
	private int noBoderbeginRow = -1;
	/**
	 * 无底纹行的结束行
	 */
	private int noBoderendRow = -1;
	/**
	 * 设置Sheet名字
	 */
	private String[] sheetNames;

	public void setListTableHeadInfo(List<TableHeadInfo> listTableHeadInfo) {
		this.listTableHeadInfo = listTableHeadInfo;
	}

	public List<TableHeadInfo> getListTableHeadInfo() {
		return listTableHeadInfo;
	}
	public void setNoBoderbeginRow(int noBoderbeginRow) {
		this.noBoderbeginRow = noBoderbeginRow;
	}

	public int getNoBoderbeginRow() {
		return noBoderbeginRow;
	}

	public void setNoBoderendRow(int noBoderendRow) {
		this.noBoderendRow = noBoderendRow;
	}

	public int getNoBoderendRow() {
		return noBoderendRow;
	}

	@Override
	public int drawFooter() {
		// TODO Auto-generated method stub
		return super.drawFooter();
	}

	@Override
	public int drawTitle() {

		return setTableHeadList();
	}

	public void setSheetNames(String[] sheetNames) {
		this.sheetNames = sheetNames;
	}

	public String[] getSheetNames() {
		return sheetNames;
	}


	/**
	 * 表头信息
	 */
	public int setTableHeadList() {
		int rowIndex = -1;	//当前行
		int rowLastIndex = -1;	//已创建行
		int colIndex = -1;	//当前列
		int otherRow = 0;	//通过段落标记<p>添加的行
		HSSFRow row = null;
		try {
			if (listTableHeadInfo == null) return 0;
			for (int i = 0; i < listTableHeadInfo.size();i++) {
				//获取第i行第j列的单元格信息
				TableHeadInfo cellInfo = listTableHeadInfo.get(i);
				//获取行内HTML文本信息，即解析单元格的信息
				List<CellInfo> headNames = checkHeadName(cellInfo.headName);
				//设置当前行号
				if (rowIndex != cellInfo.row + otherRow) {
					rowIndex = cellInfo.row + otherRow;
				}
				//如果单元格内有HTML的<p>换行标记，则作为换行处理
				otherRow += headNames.size()-1;
				//创建行以满足当前单元格信息中涉及到的最大行
				while (rowLastIndex < rowIndex + cellInfo.rowspan-1 + headNames.size()-1) {
					//增加一行
					this.addRow();
					rowLastIndex++;
				}
				//循环处理每个单元格的内容，合并后的单元格作为一个单元格处理
				for (int j = 0; j < headNames.size(); j++) {
					//获得当前行号
					row = this.getCurrSheet().getRow(rowIndex);
					//得到当前信息的列号
					colIndex = cellInfo.col;
					//声明单元格样式
					HSSFCellStyle cellStyle = this.getCurrWorkBook().createCellStyle();
									
					//设置样式,居中
					cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
					cellStyle.setBorderBottom(HSSFCellStyle.BORDER_NONE);
					//得到需要设置的单元格的信息
					CellInfo tmpCell = headNames.get(j);
					//获得字体样式
					HSSFFont font = tmpCell.getFONT();
//					//设置对齐方式
					if (cellInfo.colAlign == 1) {
						//在行样式为默认样式的情况下，优先选择单元格的样式；
						cellStyle.setAlignment(tmpCell.getColAlign());
					} else {
						//否则优先选择行样式
						cellStyle.setAlignment(cellInfo.colAlign);
					}
					//设置字体
					cellStyle.setFont(font);
					//创建单元格并设置风格
					for (int k = 0; k < cellInfo.colspan; k++) {
						//生成对应的单元格，并进行样式设置
						HSSFCell curCell = row.createCell(colIndex+k);
						curCell.setCellStyle(cellStyle);
					}
					//写入信息
					row.getCell(colIndex).setCellValue(new HSSFRichTextString(tmpCell.getCellText()));
//					//合并
					this.getCurrSheet().addMergedRegion(new CellRangeAddress(rowIndex,rowIndex+cellInfo.rowspan-1,
																			 colIndex,colIndex+cellInfo.colspan-1)	);
					//换行处理下一行
					rowIndex++;
				}
				
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			return ExportInvalid.ERR_DRAWTITLE;
		}
		return 0;
	}

	//===================================================
	/* 处理标准：严格遵守HTML的<></>的格式标准，要求开始符号<>和结束符号</>必须成对出现；否则解析会有错误，如果确实无法满足该格式，可针对具体需求
	 * 修改解析代码。从开始<tag>到结束</tag>之间的内容目前默认为一行来处理，主要为满足表头内的<p></p>标签，如果有其他情况需要修改解析代码
	 * 目前已经解析的内容有：<b>、<font>、<i>、<p>、<SMALL>,再有需要时要一定要添加解析代码，否则会有异常结果产生
	 * 目前解析内容中不包含单引号以及所有<,> 等特殊标记，所以客户端对于设置部分的文本内容尽量不要使用特殊标记，如果确实有需要，研究后再处理使用情况
	 */
	//===================================================
	//解析表头文本信息，以方便后面的对单元格进行设置
	private List<CellInfo> checkHeadName(String headName){
		//定义一个新的单元格信息列表，用于返回所有的单元格信息
		List<CellInfo> headCells = new ArrayList<CellInfo>();
		//首先去掉首尾的空格
		String tmpHeadName = headName.trim();
		//定义文本内容的格式
		HSSFFont FONT = null;
		//记录符号<的个数，用于多行处理
		int blTagNum = 0;
		//定义一个新的单元格信息
		CellInfo tmpCell = null;
		if (tmpHeadName.indexOf('<') >= 0) {
			//只要解析内容没有结束，就执行下去
			while (tmpHeadName.length() > 0) {
				//每次对文本内容的第一个字符进行判断
				switch (tmpHeadName.charAt(0)) {
					case '<':	//开始符号<的处理
						//开始时
						if (blTagNum == 0) {
							//定义新的单元格信息内容
							tmpCell = new CellInfo();
							//定义新的单元格内字体样式
							FONT = this.getCurrWorkBook().createFont();
							//对样式进行默认初始化设置
							FONT.setFontHeightInPoints((short)9);
							FONT.setFontName("宋体");
							FONT.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
							FONT.setItalic(false);
							FONT.setStrikeout(false);
						}
						//检测到'<'时，记录符号自动加一
						blTagNum++;
						//去掉'<'再继续进行解析
						tmpHeadName = tmpHeadName.substring(1).trim();
						//判断'<'后面的第一个字符
						switch (tmpHeadName.charAt(0)) {
							case 'b':	// 对<b>或<B>进行处理
							case 'B':	// 对<b>或<B>进行处理
								//单元格文本格式设置为粗体
								FONT.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							case 'f':	// 对<FONT>
							case 'F':	// 对<FONT>
								String strFont = tmpHeadName.substring(0,4).toUpperCase();
								//确认是FONT标签，才进行确切处理
								if (strFont.equals("FONT") ) {
									//单独获得FONT标签内的所有信息
									String tmpFn = tmpHeadName.substring(0,tmpHeadName.indexOf('>')+1).toUpperCase();
									//查找是否有SIZE的字体大小的设置
									if ( tmpFn.indexOf("SIZE") >= 0 ) {
										//得到SIZE部分的内容
										String tmpSize = tmpFn.substring(tmpFn.indexOf("SIZE")).trim();
										tmpSize = tmpSize.substring(tmpSize.indexOf("=") + 1).trim();
										if (tmpSize.indexOf((char)0x20) >= 0) {		//以空格结束
											tmpSize = tmpSize.substring(0,tmpSize.indexOf((char)0x20));
										} else if (tmpSize.indexOf('>') >= 0) {		//以结束符号>结束
											tmpSize = tmpSize.substring(0,tmpSize.indexOf('>'));
										}
										//对字体的大小进行设置，根据实际情况添加和处理，不保证所有的都能得到解析，但是有默认值
										switch (Short.parseShort(tmpSize)) {
											case 4:		//4号字体时设置为Excel中的12号字体
												FONT.setFontHeightInPoints((short) 14);
												break;
											case 2:		//4号字体时设置为Excel中的12号字体
												FONT.setFontHeightInPoints((short) 12);
												break;
											default:	//默认情况处理
												FONT.setFontHeightInPoints((short) 9);
											break;
										}
									}
								}
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							case 'i':	// 对<i>或者<I>进行处理
							case 'I':	// 对<i>或者<I>进行处理
								FONT.setItalic(true);
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							case 'p':	// 对<p>或者<P>进行处理
							case 'P':	// 对<p>或者<P>进行处理
								//单元格文本格式设置为斜体
								String tmpPa = tmpHeadName.substring(0,tmpHeadName.indexOf('>')).toUpperCase();
								//检查是否存在ALIGN内容
								if ( tmpPa.indexOf("ALIGN") >= 0 ) {
									//得到ALIGN的内容
									String tmpAlign = tmpPa.substring(tmpPa.indexOf("ALIGN")).trim();
									tmpAlign = tmpAlign.substring(tmpAlign.indexOf("=") + 1).trim();
									//得到ALIGN的实际设计值
									if (tmpAlign.indexOf((char)0x20) >= 0) {
										//以空格结束
										tmpAlign = tmpAlign.substring(0,tmpAlign.indexOf((char)0x20));
									} else if (tmpAlign.indexOf('>') >= 0) {
										//以结束符符号>结束
										tmpAlign = tmpAlign.substring(0,tmpAlign.indexOf('>'));
									}
									
									if ( tmpAlign.indexOf("LEFT") >= 0 ) {
										tmpCell.setColAlign((short) 1);
									}
									if ( tmpAlign.indexOf("CENTER") >= 0 ) {
										tmpCell.setColAlign((short) 2);
									}
									if ( tmpAlign.indexOf("RIGHT") >= 0 ) {
										tmpCell.setColAlign((short) 3);
									}
								} else {
									//默认左对齐
									tmpCell.setColAlign((short) 1);
								}
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							case 's':	// 对<SIZE>
							case 'S':	// 对<SIZE>
								String strSmall = tmpHeadName.substring(0,5).toUpperCase();
								if (strSmall.equals("SMALL") ) {
									FONT.setFontHeightInPoints((short) 9);
								}
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							case 'u':	// 对<UNIT>
							case 'U':	// 对<UNIT>
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							case '/':	//结束符号的处理
								blTagNum = blTagNum - 2;
								//结束时
								if (blTagNum == 0) {
									//把单元格字体样式设置到字体信息对象中
									tmpCell.setFONT(FONT);
									//把当前的字体对象信息插入到列表中，已备后面的处理
									headCells.add(tmpCell);
								}
								//结束当前标签的处理
								tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('>')+1);
								break;
							default:
								break;
						}
						break;
					case (char)0x09:		//水平跳格
					case (char)0x0A:		//换行
					case (char)0x0D:		//回车
					case (char)0x20:		//空格
						//自动去掉上述特殊字符
						tmpHeadName = tmpHeadName.substring(1);
						break;
					default:
						//default的内容作为单元格内的文本内容处理，所以在出现新的标签时一定要处理解析，否则会出现异常解析结果
						
						tmpCell.setCellText(tmpHeadName.substring(0,tmpHeadName.indexOf('<')));
						tmpHeadName = tmpHeadName.substring(tmpHeadName.indexOf('<'));
					
						break;
				}
			}
		}
		//如果内容为空，则设置默认值处理
		if (headCells.size() == 0) {
			FONT = this.getCurrWorkBook().createFont();
			FONT.setFontHeightInPoints((short)9);
			FONT.setFontName("宋体");
			FONT.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
			FONT.setItalic(false);
			FONT.setStrikeout(false);
			tmpCell = new CellInfo();
			tmpCell.setFONT(FONT);
			tmpCell.setCellText(tmpHeadName);
			headCells.add(tmpCell);
		}
		return headCells;
	}
	/**
	 * 设置无底纹的行效果，目前只支持对正行设置，如果需要对局部块设置，需要修改结构，参数可以以数组的形式传递
	 * @return 成功返回非负
	 */
	@Override
	protected int drawBorder()
	{
		//先执行调用父类函数
		super.drawBorder();
		try{
			//获取所有的列数
			int columnCount = getColumnCount();
			//对所有的sheet进行循环处理
			for(int sheet =0;sheet < this.getCurrWorkBook().getNumberOfSheets();sheet++)
			{
				//循环处理设置行
				for(int row= noBoderbeginRow;row <= noBoderendRow;row++)
				{
					//如果当前行不存在，跳过处理
					if(null == this.getCurrWorkBook().getSheetAt(sheet).getRow(row))
						continue;
					//对所有列进行处理
					for(int cell=0;cell <= columnCount;cell++)
					{
						if(null == this.getCurrWorkBook().getSheetAt(sheet).getRow(row).getCell(cell)){
							this.getCurrWorkBook().getSheetAt(sheet).getRow(row).createCell(cell);						
							HSSFCellStyle cs = this.getCurrWorkBook().createCellStyle();			
							//取消边框
							cs.setBorderTop(HSSFCellStyle.BORDER_NONE);
							cs.setBorderBottom(HSSFCellStyle.BORDER_NONE);
							cs.setBorderRight(HSSFCellStyle.BORDER_NONE);
							if (cell == columnCount) {	//最后一列时，设置左边框为默认格式
								cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							} else {
								cs.setBorderLeft(HSSFCellStyle.BORDER_NONE);
									//设置无底纹填充=====无数次的实验结果，原理尚不明确
								cs.setFillPattern((short)65);
								cs.setFillForegroundColor((short)65);
							}
							//设置样式				
							this.getCurrWorkBook().getSheetAt(sheet).getRow(row).getCell(cell).setCellStyle(cs);
							continue;
						}
						HSSFCellStyle cs = this.getCurrWorkBook().getSheetAt(sheet).getRow(row).getCell(cell).getCellStyle();
						//取消边框
						cs.setBorderTop(HSSFCellStyle.BORDER_NONE);
						cs.setBorderBottom(HSSFCellStyle.BORDER_NONE);
						cs.setBorderRight(HSSFCellStyle.BORDER_NONE);
						if (cell == columnCount) {	//最后一列时，设置左边框为默认格式
							cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
						} else {
							cs.setBorderLeft(HSSFCellStyle.BORDER_NONE);
								//设置无底纹填充=====无数次的实验结果，原理尚不明确
							cs.setFillPattern((short)65);
							cs.setFillForegroundColor((short)65);
						}
						//设置样式
						this.getCurrWorkBook().getSheetAt(sheet).getRow(row).getCell(cell).setCellStyle(cs);
					}
				}
			}
		
		}catch(Exception e)
		{
			e.printStackTrace();
			return ExportInvalid.ERR_DRAWBORDER;
		}
		//设置sheet名字
		if (sheetNames.length > 0) {
			setAllSheetName(sheetNames);
		}
		return 0;
	}

	//设置所有Sheet的名字
	private void setAllSheetName(String[] sheetNames) {
		for (int i = 0; i < sheetNames.length;i++) {
			setSheetName(i,sheetNames[i].toString());
		}
			
	}
	//设置第i个Sheet的名字,该方法为内部处理，目前没有对外部暴露，因为Sheet的整个生成过程都在内部处理，所以对于名称的设置目前只满足批量处理，
	//通过索引来设置变得似乎没有必要
	private void setSheetName(int iIndex,String sheetName) {
		//得到当前工作文档
		HSSFWorkbook curWorkBook = this.getCurrWorkBook();
		//判断索引指向的Sheet是否存在，不存在则不做处理
		if (iIndex >= curWorkBook.getNumberOfSheets()) return;
		//检出Sheet的名字是否存在，如果不存在则不做处理
		if (sheetName==null || sheetName.length() <= 0) return;
		//设置Sheet的名字
		curWorkBook.setSheetName(iIndex, sheetName);
	}

}
