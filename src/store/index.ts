// import { defineStore } from 'pinia'
// import axios from 'axios'
// // 自定义国际化配置
// import { createI18n } from 'vue-i18n';
// const useStore = defineStore('store', {
//   state: () => ({
//     en: {
//       navigateBar: {
//         hotspot: "Hotspot",
//         experience: "Experience",
//         focus: "Focus",
//         recommend: "Recommend"
//       },
//       tabs: {
//         work: "Work",
//         private: "Private",
//         collect: "Collect",
//         like: "Like"
//       }
//     },
//     zh: {
//       navigateBar: {
//         hotspot: "热点",
//         experience: "经验",
//         focus: "关注",
//         recommend: "推荐"
//       },
//       tabs: {
//         work: "作品",
//         private: "私密",
//         collect: "收藏",
//         like: "喜欢"
//       }
//     }
//   }),
//   getters: {},
//   actions: {
//     getLang(lang = 'en') {
//       const i18n = createI18n({
//         legacy: false,  // 设置为 false，启用 composition API 模式
//         messages: { ...this.en, ...this.zh },
//         locale: lang
//       })
//       return i18n
//     },
//     changeLang(lang: string) {
//       let that = this;
//       if (lang === 'zh') {
//         axios({
//           method: "get",
//           url: "src/zh-cn.json",
//         })
//           .then(function (res) {
//             that.zh = res.data;
//             that.getLang(lang);
//           })
//       } else if (lang === 'en') {
//         axios({
//           method: "get",
//           url: "src/en.json",
//         })
//           .then(function (res) {
//             that.en = res.data;
//             that.getLang(lang);
//           })
//       }
//     }
//   }
// })
// export default useStore
