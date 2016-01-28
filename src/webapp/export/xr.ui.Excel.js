(function ($) {
    /**
     * 创建 源码
     * @param $this
     * @returns {boolean}
     */
    function createHtml($this) {   
    	var pageCName = "xr_excel";
        var pagePreBtn = "pre_btn";
        var flatbtn = "btn-default";
 
        var pageDiv = $("<div></div>").addClass(pageCName);
        var PreBtn = $("<a></a>").addClass(flatbtn).attr({"href": "javascript:void(0);"});
        var preBtnName = $("<i></i>").addClass(pagePreBtn).html("导出");
        PreBtn.append(preBtnName);
        pageDiv.append(PreBtn);
        /*
        var PreBtn1 = $("<a></a>").addClass(flatbtn).attr({"href": "javascript:void(0);"});
        var preBtnName1 = $("<i></i>").addClass(pagePreBtn).html("导出全部");
        PreBtn1.append(preBtnName1);
        pageDiv.append(PreBtn1);
        */
        $this.append(pageDiv);

         
    } 

    /**
     *  
     * @param $this
     */
    function formPageClick($this) {
        var allAs = $this.find("a");
        allAs.each(function (i, val) {
            $(val).click(function (e) {  
            	
            	var baseurl = $("#home").val() ;
        		var url = options.url;

        		var datatables = $(options.datatable).dataTable();
        		DownloadExcel_downloadExcel(baseurl,url,datatables);
        		
            });
        });
    };

    //public method
    var methods = {
        init: function (initOptions) {
            options = $.extend({}, $.fn.xrExcel.defaults, initOptions);
            var $this = $(this);
            return this.each(function () {
                createHtml($this);
                formPageClick($this)
            });
        },
        destroy: function () {
            return this.each(function () {
            });
        },
        option: function (key, value) {
            if (arguments.length == 2)
                return this.each(function () {
                    if (options[key]) {
                        options[key] = value;
                    }
                });
            else
                return options[key];
        }
    }

    var methodName = "xrExcel";

    var options = {};

    /**
     *  插件入口
     */
    $.fn.xrExcel = function () {
        var method = arguments[0];
        if (methods[method]) {
            method = methods[method];
            arguments = Array.prototype.slice.call(arguments, 1);
        } else if (typeof method === "object" || !method) {
            method = methods.init;
        } else {
            $.error("Method(" + method + ") does not exist on " + methodName);
            return this;
        }
        return method.apply(this, arguments);
    }


    $.fn.xrExcel.defaults = {

    };
})(jQuery);

