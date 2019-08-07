'use strict';
var fs = require('fs');
var fs = require('fs');
var moment = require('moment');
var path = require('path');
var nodeExcel = require('excel-export');
var sql = require('./sql')


module.exports = function (router) {

    //文件下载
    router.get('/download', async function (ctx) {
		let query = ctx.query;
		var d = moment("2019"+query.selectMonth,"YYYY-MM"); //按照指定的年月字符串和格式解析出一个moment的日期对象
		var firstDate = d.startOf("month").format("YYYY-MM-DD"); //通过startOf函数指定取月份的开始即第一天
		var lastDate = d.endOf("month").format("YYYY-MM-DD"); //通过startOf函数指定取月份的末尾即最后一天
		var totalDay = d.daysInMonth(); //计算总天数
		var sqlString = 'select a.username,a.outdate,a.department,a.place,b.price from outs a left join price b on a.place = b.place where a.place !="" and a.outdate>="'+firstDate+'" and a.outdate<="'+lastDate+'"'
		console.log(sqlString);
        let sql_data = await sql.get(sqlString);
        let user_name = await sql.get('select username from users order by department');
        //检查是否存在文件，不存在则创建：
            let outFilename = "style.xml";
            var conf = {};
            conf.stylesXmlFile = outFilename;//输出文件名
            conf.name = "outprice";//表格名
            //定义表格列字段
            conf.cols = [
                {
                    caption: 'name',
                    type: 'string',
                    width: 20
                }
            ];

            for(let i = 1; i <= (totalDay+1); i++){

                if(i == (totalDay+1)){
                    conf.cols.push({
                        caption: '总计',
                        type: 'number',
                        width: 10
                    })
                }else{
                    conf.cols.push({
                        caption: i,
                        type: 'string',
                        width: 10
                    })
                }

            }
            //填充表格数据
            conf.rows = [];
            for(let i = 0; i < user_name.length; i++){
                let arr = []
                let exUserName = user_name[i].username
                arr.push(exUserName);
                for(let i = 1; i <= totalDay; i++){
                    arr.push('')
                }
                let price = 0;
                for(let d = 0; d < sql_data.length; d++){
                    if(exUserName == sql_data[d].username){
                        let place = sql_data[d];
                        // let num = parseInt(place.outdate.slice(8,10))
						let num = parseInt(moment(place.outdate).date());
						console.log(moment(place.outdate).date())
                        arr[num] = place.place
                        price += place.price || 0;
                    }
                }
                arr.push(price);

                conf.rows.push(arr)
            }

            //生成excel文件
            var result = await nodeExcel.execute(conf);
            //将数据转为二进制输出
            let data = new Buffer(result, 'binary');
            ctx.set('Content-Type', 'application/vnd.openxmlformats');
            ctx.set("Content-Disposition", "attachment; filename=" + "outprice.xlsx");
            ctx.body = data;
        }
    );
	
	router.get('/download2', async function (ctx) {
		let query = ctx.query;
		var d = moment("2019"+query.selectMonth,"YYYY-MM"); //按照指定的年月字符串和格式解析出一个moment的日期对象
		var firstDate = d.startOf("month").format("YYYY-MM-DD"); //通过startOf函数指定取月份的开始即第一天
		var lastDate = d.endOf("month").format("YYYY-MM-DD"); //通过startOf函数指定取月份的末尾即最后一天
		var totalDay = d.daysInMonth(); //计算总天数
		var sqlString = 'select a.username,a.outdate,a.department,a.place,b.price from outs a left join price b on a.place = b.place where a.place !="" and a.outdate>="'+firstDate+'" and a.outdate<="'+lastDate+'"'
		let sql_data = await sql.get(sqlString);
		console.log(sql_data);
		let user_name = await sql.get('select username from users order by department');
		let area = await sql.get('select place from price where place!="上海" order by price');
	    //检查是否存在文件，不存在则创建：
	        let outFilename = "style.xml";
	        var conf = {};
	        conf.stylesXmlFile = outFilename;//输出文件名
	        conf.name = "out";//表格名
	        //定义表格列字段
	        conf.cols = [
	            {
	                caption: '姓名',
	                type: 'string',
	                width: 20
	            }
	        ];
	
	        for(let i = 0; i <= area.length; i++){
	
	            if(i == area.length){
	                conf.cols.push({
	                    caption: '总计',
	                    type: 'number',
	                    width: 10
	                })
	            }else{
	                conf.cols.push({
	                    caption: area[i].place,
	                    type: 'string',
	                    width: 10
	                })
	            }
	
	        }
	        //填充表格数据
	        conf.rows = [];
	        for(let i = 0; i < user_name.length; i++){
	            let arr = []
	            let exUserName = user_name[i].username
	            arr.push(exUserName);
	            for(let i = 1; i <= 31; i++){
	                arr.push('')
	            }
	            let price = 0;
	            for(let d = 0; d < sql_data.length; d++){
	                if(exUserName == sql_data[d].username){
	                    let place = sql_data[d];
	        			console.log(place)
	                    let num = parseInt(place.outdate.slice(8,10))
	                    arr[num] = place.place
	                    price += place.price || 0;
	                }
	            }
	            arr.push(price);
	        
	            conf.rows.push(arr)
	        }
	
	        //生成excel文件
	        var result = await nodeExcel.execute(conf);
	        //将数据转为二进制输出
	        let data = new Buffer(result, 'binary');
	        ctx.set('Content-Type', 'application/vnd.openxmlformats');
	        ctx.set("Content-Disposition", "attachment; filename=" + "out.xlsx");
	        ctx.body = data;
	    }
	);
	
	router.get('/download3', async function (ctx) {
		let query = ctx.query;
		var d = moment("2019"+query.selectMonth,"YYYY-MM"); //按照指定的年月字符串和格式解析出一个moment的日期对象
		var firstDate = d.startOf("month").format("YYYY-MM-DD"); //通过startOf函数指定取月份的开始即第一天
		var lastDate = d.endOf("month").format("YYYY-MM-DD"); //通过startOf函数指定取月份的末尾即最后一天
		var totalDay = d.daysInMonth(); //计算总天数
		var sqlString = 'select username,overtimedate,daynum from overtime where  overtimedate>="'+firstDate+'" and overtimedate<="'+lastDate+'"'
		console.log(sqlString);
	    let sql_data = await sql.get(sqlString);
	    let user_name = await sql.get('select username from users order by department');
	    //检查是否存在文件，不存在则创建：
	        let outFilename = "style.xml";
	        var conf = {};
	        conf.stylesXmlFile = outFilename;//输出文件名
	        conf.name = "overtime";//表格名
	        //定义表格列字段
	        conf.cols = [
	            {
	                caption: '姓名',
	                type: 'string',
	                width: 20
	            }
	        ];
	
	        for(let i = 1; i <= (totalDay+1); i++){
	
	            if(i == (totalDay+1)){
	                conf.cols.push({
	                    caption: '总计',
	                    type: 'number',
	                    width: 10
	                })
	            }else{
	                conf.cols.push({
	                    caption: i,
	                    type: 'string',
	                    width: 10
	                })
	            }
	
	        }
	        //填充表格数据
	        conf.rows = [];
	        for(let i = 0; i < user_name.length; i++){
	            let arr = []
	            let exUserName = user_name[i].username
	            arr.push(exUserName);
	            for(let i = 1; i <= totalDay; i++){
	                arr.push('')
	            }
	            let overtimeSum = 0;
	            for(let d = 0; d < sql_data.length; d++){
	                if(exUserName == sql_data[d].username){
	                    let place = sql_data[d]
						console.log(place.overtimedate)
	                    // let num = parseInt(place.overtimedate.slice(8,10))
						let num = parseInt(moment(place.overtimedate).date());
	                    arr[num] = place.daynum
	                    overtimeSum += place.daynum || 0;
	                }
	            }
	            arr.push(overtimeSum);
	
	            conf.rows.push(arr)
	        }
	
	        //生成excel文件
	        var result = await nodeExcel.execute(conf);
	        //将数据转为二进制输出
	        let data = new Buffer(result, 'binary');
	        ctx.set('Content-Type', 'application/vnd.openxmlformats');
	        ctx.set("Content-Disposition", "attachment; filename=" + "overtime.xlsx");
	        ctx.body = data;
	    }
	);
	
    return router;
}

