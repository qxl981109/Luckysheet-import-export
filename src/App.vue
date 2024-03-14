<script setup>
import { onMounted, ref } from "vue";
const pdfurl =''
</script>

<template>
  <div class="main-page">
    <div class="pdf-container">
      <iframe :src="`pdf/web/viewer.html?file=${pdfurl}`" class="pdf-window" width="100%" height="100%"
        frameborder="0"></iframe>
    </div>
  </div>
</template>
<style>
.pdf-container {
  width: 100%;
  height: 100vh;
}
</style>