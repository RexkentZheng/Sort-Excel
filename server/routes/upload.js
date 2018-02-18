let express = require("express");
let router = express.Router();
let multer = require("multer");
let fs = require("fs");
let upload = require("./../utils/upload");
let ejsExcel = require("ejsexcel");
let exlBuf = fs.readFileSync("./model/model.xlsx");
let path = require("path");

function getWrong(res, err) {
  return res.json({
    status: 1,
    msg: err.message,
    result: ""
  });
}

function getRight(res, doc) {
  return res.json({
    status: 0,
    msg: "",
    result: doc
  });
}

router.post("/upload", (req, res, body) => {
  upload.uploadTest(req, res, err => {
    if (err) {
      getWrong(res, err);
    } else {
      let path = req.file.path;
      let exBuf = fs.readFileSync(path);
      ejsExcel.getExcelArr(exBuf).then(exlJson => {
        let workBook = exlJson;
        let workSheets = workBook[0];
        let all = [];
        //输出原始数据表
        for (i in workSheets) {
          if (i != 0) {
            let arr = {};
            for (x in workSheets[i]) {
              arr[workSheets[0][x]] = workSheets[i][x];
            }
            all.push(arr);
          }
        }
        for (i = 0; i < 3; i++) {
          userId = i;
        }

        var getKeyInfo = (arr, onePiece) => {
          let op = {};
          op["客户全称"] = onePiece["购货单位"];
          op["K3销售订单号"] = onePiece["单据编号"];
          op["合同编号"] = onePiece["合同号："];
          op["品牌"] = onePiece["品牌"];
          op["价税合计"] = onePiece["价税合计"];
          op["合同日期"] = onePiece["审核日期"];
          op["销售"] = onePiece["业务员"];
          op["发货RMB"] =
            onePiece["实际含税单价"] * onePiece["常用单位出库数量"];
          op["发货日期"] = onePiece["预计交货日期"];
          op["HKD"] = 0.0;
          arr.push(op);
        };
        var getPriceInfo = (type, arr, onePiece) => {
          let op = {};
          op["客户全称"] = onePiece["客户全称"];
          op["K3销售订单号"] = onePiece["K3销售订单号"];
          op["合同编号"] = onePiece["合同编号"];
          op["" + type + ""] = onePiece["价税合计"];
          op["合同日期"] = onePiece["合同日期"];
          op["销售"] = onePiece["销售"];
          op["发货RMB"] = onePiece["发货RMB"];
          op["发货日期"] = onePiece["发货日期"];
          op["RA销售"] = "无";
          op["HKD"] = 0.0;
          arr.push(op);
        };
        //筛选出个人数据表
        let name = "李彩炜";
        let personalData = [];
        all.forEach(onePiece => {
          if (onePiece.制单 === name) {
            getKeyInfo(personalData, onePiece);
          }
        });
        //合并价格，生成最终结果
        let endData = [];
        for (i in personalData) {
          if (endData.length === 0) {
            getPriceInfo(personalData[i]["品牌"], endData, personalData[i]);
          } else {
            let isInside = false;
            for (j in endData) {
              if (
                endData[j]["K3销售订单号"] === personalData[i]["K3销售订单号"]
              ) {
                isInside = true;
                let type = personalData[i]["品牌"];
                if (endData[j]["" + type + ""]) {
                  endData[j]["" + type + ""] =
                    parseFloat(endData[j]["" + type + ""]) +
                    parseFloat(personalData[i]["价税合计"]);
                } else {
                  endData[j]["" + type + ""] = parseFloat(
                    personalData[i]["价税合计"]
                  );
                }
              }
            }
            if (!isInside) {
              getPriceInfo(personalData[i]["品牌"], endData, personalData[i]);
            }
          }
        }
        //添加剩余数据
        endData.forEach(op => {
          if (!op["AB"]) {
            op["AB"] = 0.0;
          }
          if (!op["第三方"]) {
            op["第三方"] = 0.0;
          }
          op["Total"] = op["AB"] + op["第三方"];
        });
        function fileName() {
          let nowDate = new Date();
          let fileName =
            "北京地区" +
            nowDate.getFullYear() +
            "-" +
            (nowDate.getMonth() + 1) +
            "月报" +
            name;
          return fileName;
        }
        //用数据源(对象)data渲染Excel模板
        ejsExcel
          .renderExcel(exlBuf, endData)
          .then(function(exlBuf2) {
            fs.writeFileSync("./result/" + fileName() + ".xlsx", exlBuf2);
            console.log("end.xlsx");
          })
          .catch(function(err) {
            console.error(err);
          });
        //删除上传数据表
        fs.unlink("./excel/" + path.split("\\")[1], function(err) {
          if (err) {
            return console.log(err);
            console.log("文件删除失败");
          } else {
            console.log("文件删除成功");
          }
        });
        res.json({
          status: 0,
          msg: "",
          result: 123,
          fileName: fileName()
        });
      });
    }
  });
});

router.get("/download", function(req, res, next) {
  var currDir = path.normalize(req.query.dir),
    fileName = req.query.name,
    currFile = path.join(currDir, fileName),
    fReadStream;

  fs.exists(currFile, function(exist) {
    if (exist) {
      res.set({
        "Content-type": "application/octet-stream",
        "Content-Disposition": "attachment;filename=" + encodeURI(fileName)
      });
      fReadStream = fs.createReadStream(currFile);
      fReadStream.on("data", chunk => res.write(chunk, "binary"));
      fReadStream.on("end", function() {
        res.end();
      });
    } else {
      res.set("Content-type", "text/html");
      res.send("file not exist!");
      res.end();
    }
  });
});
//暴露路由
module.exports = router;
