var multer = require('multer');


//测试接口
var storageTest = multer.diskStorage({
	destination: function(req, file, callback) {
		callback(null, "./excel");
	},
	filename: function(req, file, callback) {
		callback(null, file.fieldname + "_" + Date.now() + "_" + file.originalname);
	},
});

//定义upload方法
var uploadTest = multer({
	storage: storageTest
}).single("test"); 

exports.uploadTest = uploadTest;
