<template>
  <el-button type="primary" @click="exportSheet" style="display: black;">导出</el-button>
  <div class="excelItem">
    <div id="luckysheet" class="luckysheet"></div>
  </div>
</template>

<script setup>
import { onMounted, ref } from "vue";
import LuckyExcel from "luckyexcel";
import exportExcel from "./export.js"
const uploadExcel = (evt) => {
  const files = evt.target.files;
  if (files == null || files.length == 0) {
    alert("No files wait for import");
    return;
  }

  let name = files[0].name;
  let suffixArr = name.split("."),
    suffix = suffixArr[suffixArr.length - 1];
  if (suffix != "xlsx") {
    alert("Currently only supports the import of xlsx files");
    return;
  }
  LuckyExcel.transformExcelToLucky(
    files[0],
    (exportJson, luckysheetfile) => {
      if (exportJson.sheets == null || exportJson.sheets.length == 0) {
        alert(
          "Failed to read the content of the excel file, currently does not support xls files!"
        );
        return;
      }
      window.luckysheet.destroy();

      window.luckysheet.create({
        container: "luckysheet", //luckysheet is the container id
        showinfobar: false,
        data: exportJson.sheets,
        title: exportJson.info.name,
        lang: "zh", // 设定表格语言
        plugins: ["chart"],
        userInfo: exportJson.info.name.creator,
      });
    }
  );
}
const exportSheet = () => {
  exportExcel(window.luckysheet.getluckysheetfile(), "报表名称");
}
// 打开用户弹窗组件
const handleOpen = (evl) => {
  uploadExcel(evl);
}

// 子组件默认包含是私有
defineExpose({
  handleOpen
})
</script>

<style>
.excelItem {
  width: 100%;
  height: 800px;
}
.luckysheet {
  width: 100%;
  height: 800px;
}
</style>
