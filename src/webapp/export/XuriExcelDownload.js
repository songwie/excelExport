
/**
	 * 构造函数
	 * @param index:序号
	 * @param colName:列名
	 * @param colData:列取值索引
	 * @param colWidth:列宽,默认100
	 * @param colDataType:数据类型
	 * @param colAlign:列对齐方式
	 * @param colHidden 是否隐藏
	 */
DownloadExcel_getColumnInfo =  function(datatable)
{
	var tablesId = datatable.selector;
 
	var columnInfo = [];
	if(datatable.dataTableSettings){
		var tables = datatable.dataTableSettings;
		for(var i=0;i<tables.length;i++){
			var table = tables[i];
			if(table.sInstance==tablesId){
				var cums = table.aoColumns;
				for(var i=0;i<cums.length;i++){
					var datacul = cums[i];
					var column = {};
					column.index = i;
					column.colName = datacul.sTitle;
					column.colWidth = datacul.sWidth ;
					column.colDataType = datacul.sType;
					column.colAlign = datacul.sClass;
					var fHidden = false;
					if(datacul.fExportHidden || !datacul.bVisible ){
						fHidden = true;
						column.colWidth = "0px";
					}
					column.colHidden = fHidden ;
					column.colData = datacul.mData;

					columnInfo.push(column);
				}
			}
		}
		
	}

	if(columnInfo.length ==0)
	{
		return false;
	} 

	return columnInfo;
}

DownloadExcel_downloadExcel = function(baseurl,url,datatable, ntype,cols,params)
{
 
	var columnInfo = DownloadExcel_getColumnInfo(datatable);
	if(columnInfo == false)
	{
		return false;
	}
	var data = { };
	if(cols){
	    data.columnInfo =  JSON.stringify(cols) ;
	}else{
	    data.columnInfo =  JSON.stringify(columnInfo) ;
	}
    data.params =  JSON.stringify(params) ;

    $.ajax({
		  type: "post",
		  url: baseurl + url,
		  timeout:999999999,
		  data : data,
		  dataType:'json',
		  contentType: "application/x-www-form-urlencoded; charset=utf-8",
		  success: function(msg){
			  if(typeof msg=='string'){
				  var rs = $.parseJSON(msg);
			  }else{
				  var rs = msg;
			  }
			  if (rs!=null&&rs.success==false){
				  alert("导出失败："+msg.error);
			  }else{
					var data = {
						filename : rs.fileAddress
					};
					var form = document.createElement("form");
					document.body.appendChild(form);
					for(var name in data){
						var i = document.createElement("input");
						i.type = "hidden";
						form.appendChild(i);
						i.value = data[name];
						i.name = name;
					}

					form.action = baseurl + "excel/down1";
					form.method="GET";
					form.target="_self";
					form.submit();

			  }
		  }
	});

}







