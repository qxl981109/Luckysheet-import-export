import { createApp } from 'vue';
import App from './App.vue';
// import { createPinia } from 'pinia';
// const pinia = createPinia();
import { i18n } from './lang';
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
// 创建 Vue 实例
const app = createApp(App);
//  app.use(pinia);

// 注册对象
app.use(i18n);
app.use(ElementPlus)
// 挂载到 Dom 元素中
app.mount('#app');
