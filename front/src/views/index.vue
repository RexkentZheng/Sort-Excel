<template>
  <div>
    <div class="contain-body">
	      <div class="wraper">
					<el-upload
			  		ref="upload"
					  action="/sortExcel/upload"
					  name="test"
					  :multiple="false"
					  :limit="1"
					  accept=".xls,.xlsx"
					  :on-success="test"
					  :show-file-list="true"
					  :auto-upload="false"
					  :file-list="fileList">
					  <a slot="trigger" class="cbtn" href="javascript:;" type="primary">点击导入Excel文件</a>
					  <a class="cbtn" type="success" @click="submitUpload">确认上传Excel文件</a>
					</el-upload>
				</div>
			</div>
  </div>
</template>
<script>
import axios from "axios";

export default {
  data() {
    return {
      fileList: [],
      fileName: "",
      dir: "result"
    };
  },
  mounted(){
    this.changeTitle();
  },
  methods: {
    changeTitle(){
      document.title = 'For PP';
    },
    submitUpload() {
      this.$refs.upload.submit();
    },
    test(response, file, fileList) {
      let res = response;
      console.log(res);
      if (res.status == 0) {
        console.log("导入成功");
        this.fileName = res.fileName;
        //设置定时器，否则文件未修改就会被下载
        setTimeout(() => {
          this.download();
        }, 1000);
        console.log(this.fileName);
      } else {
        console.log("导入失败");
      }
    },
    download() {
      window.location.href =
        "http://39.106.140.189:8182/sortExcel/download?dir=result&name=" +
        this.fileName +
        ".xlsx";
    }
  }
};
</script>
<style>
.contain-body {
  margin-top: 20px;
  margin-bottom: 50px;
}
.wraper {
  width: 500px;
  margin: 0 auto;
}
.cbtn {
  width: 300px;
  padding: 0px 85px;
  height: 60px;
  line-height: 60px;
  position: relative;
  cursor: pointer;
  color: #888;
  background: #fafafa;
  border: 1px solid #ddd;
  border-radius: 4px;
  overflow: hidden;
  display: inline-block;
  margin: 10px 0;
  text-align: center;
}
</style>

