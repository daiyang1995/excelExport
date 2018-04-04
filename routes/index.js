var express = require('express');
var router = express.Router();
var xlsx = require("node-xlsx");
var fs = require('fs');

/* GET home page. */
router.get('/', function(req, res, next) {
    // var list = xlsx.parse("excel/test.xlsx");
    // praseExcel(list);
    var list =[];
    list[0].name= "1";
    list[1].name="2";
    console.log(list)
});

//解析Excel
function praseExcel(list)
{
    console.log("qqq");
    for (var i = 0; i < list.length; i++)
    {
        var excleData = list[i].data;
        var sheetArray  = [];
        var typeArray =  excleData[1];
        var keyArray =  excleData[2];
        for (var j = 3; j < excleData.length ; j++)
        {
            var curData = excleData[j];
            if(curData.length == 0) continue;
            var item = changeObj(curData,typeArray,keyArray);
            sheetArray.push(item);
        }
        if(sheetArray.length >0)
            writeFile(list[i].name+".json",JSON.stringify(sheetArray));
    }
    console.log("qqq");
}


module.exports = router;
