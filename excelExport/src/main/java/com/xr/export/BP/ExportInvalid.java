package com.xr.export.BP;

import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

/**
 * @author sw
 */

@Component
@Scope("prototype")
public class ExportInvalid {
	
	public static final int ERR_INITIALIZE = -1001;
	public static final int ERR_DRAWTITLE = -1002;
	public static final int ERR_DRAWHEAD = -1003;
	public static final int ERR_DRAWMAINDATABLOCK = -1004;
	public static final int ERR_DRAWFOOTER = -1005;
	public static final int ERR_WRITEFILE = -1006;
	public static final int ERR_COLUMNTYPE = -1007;
	public static final int ERR_DRAWBORDER = -1008;
	public static final int ERR_NOCOLUMNINFO = -1009;
	public static final int ERR_CREATECOLUMNSYTLE = -1010;
	
	public String Invalid(int err){
		StringBuffer sbError  = new StringBuffer();
		sbError.append("导出出错");
		switch(err){
			case ERR_INITIALIZE :sbError.append(":导出初始化出错!");break;
			case ERR_DRAWTITLE :sbError.append(":画表头出错!");break;
			case ERR_DRAWHEAD :sbError.append(":画列头出错!");break;
			case ERR_DRAWMAINDATABLOCK :sbError.append(":画数据块出错!");break;
			case ERR_DRAWFOOTER :sbError.append(":画表尾出错!");break;
			case ERR_COLUMNTYPE :sbError.append(":转换列头类型出错!");break;
			case ERR_DRAWBORDER :sbError.append(":画边框出错!");break;
			case ERR_NOCOLUMNINFO :sbError.append(":缺少列信息!");break;
			case ERR_CREATECOLUMNSYTLE :sbError.append(":创建每列类型!");break;
			default:sbError.append(":未知错误!");break;
		}
		return sbError.toString();
	}
}
