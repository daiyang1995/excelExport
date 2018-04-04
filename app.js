var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');

var index = require('./routes/index');
var users = require('./routes/users');

var app = express();

var xlsx = require("node-xlsx");
var fs = require('fs');

//去除所有空格， 包括行内空格
function Trim(str,is_global)

{
    var result;
    result = str.replace(/(^\s+)|(\s+$)/g,"");
    if(is_global.toLowerCase()=="g")
    {
        result = result.replace(/\s/g,"");

    }
    return result;

}

var list = xlsx.parse("test.xlsx");
var data = list[0].data;
var jsonArray = [];
var showTimeList = [] ;

data.forEach(function(item,index){
  var json = {};
  json.id= item[0];
  json.name = Trim(item[1],"g");
  json.grade = item[2];
  jsonArray[index] = json;

  var jsonShowTime = {};
  jsonShowTime.id =item[0];
  jsonShowTime.time = 0 ;
  showTimeList[index] = jsonShowTime;
});


var list1 = xlsx.parse("test1.xlsx");
var length = list1[0].data[0].length;
var data1 = list1[0].data;
var jsonArray1 = [];
var changeValue = 0;
var finishIDArray = [];
data1.forEach((item,index)=>{
    if(index > 2){
        var id = item[4];
        var name = Trim(item[3],"g");
        jsonArray.forEach( (obj , idx) =>{
           if(obj.id == id ){
              data1[index][length+1] = obj.grade;
              var json = showTimeList[idx];
              json.time++;
              showTimeList[idx] = json;
              finishIDArray[changeValue++] = obj.name;
              return;
           }
        });
    }
});
var buffer = xlsx.build([{name: "mySheetName", data: data1}]); // Returns a buffer
fs.writeFile("finish.xlsx", buffer,function () {
    console.log(`本次共有${jsonArray.length}个成绩， 其中更改的为 ${changeValue}个`);
    var index = [];
    showTimeList.forEach( (obj , idx ) => {
        if(obj.time > 1){
            console.log("重复Id 为 ： " + obj.id);
        }
    })



    console.log("success")
})







module.exports = app;
