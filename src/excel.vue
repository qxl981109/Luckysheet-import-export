<template>
  <div class="viewExcel">
    <input class="import" type="file" @change="importFile" />
    <SheetView ref="useSheet" />
  </div>
</template>

<script setup>
import { onMounted, ref } from "vue";
import SheetView from "./components/luckysheetView.vue";
const useSheet = ref(null)
const importFile = (file) => {
  useSheet.value.handleOpen(file)
};

</script>

<style>
*{
  margin: 0;
  padding: 0;
}
.viewExcel{
  width: 100%;
  height: 800px;
}
.import {
  white-space: nowrap;
  cursor: pointer;
  background: #FFF;
  border: 1px solid #08090a;
  color: #606266;
  text-align: center;
  margin: 20px;
  margin-right: 60px;
  transition: .1s;
  font-weight: 500;
  font-size: 16px;
  border-radius: 4px;
}
</style>
