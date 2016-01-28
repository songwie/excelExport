# excelExport
excel 公用导出组件，提供统一参数，然后导出excel 组件包
前端 或者导出传递参数，服务端会自动导出。

/**
   * 构造函数
   * @param index:序号
   * @param colName:列名 （必传）
   * @param colData:列取值（必传）
   * @param colWidth:列宽,默认100
   * @param colDataType:数据类型，默认字符串
   * @param colAlign:列对齐方式
   * @param colHidden:是否隐藏列
*/


例：
[

{
   index: 0, 
   colName: "商品ID", 
   colData: "goodsId", 
   colWidth: "47px", 
   colDataType: "numeric", 
   colAlign: "", 
   colHidden: false
 }, 

 {
   index: 1, 
   colName: "商品编号", 
   colData: "goodsNo", 
   colWidth: "84px", 
   colDataType: "string", 
   colAlign: "", 
   colHidden: false
 }
 
]
 
代码引入：
引入com.xuri.export下所有包到工程中
引入export/下所有js代码到工程webapp目录下
