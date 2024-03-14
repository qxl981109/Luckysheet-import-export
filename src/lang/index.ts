// 自定义国际化配置
import { createI18n } from 'vue-i18n';
import axios from 'axios'
let en = {}
let zh = {}
let messages = {
  en,
  zh
}

// 创建实例对象
const i18n = createI18n({
  legacy: false,  // 设置为 false，启用 composition API 模式
  messages,
  locale: 'en'
})
// 向后台发请求
const setEn = async (lang: string, elelocale: any) => {
  let localeMessage = await axios({
    method: "get",
    url: "src/lang-server/en.json",
  });
  i18n.global.setLocaleMessage(lang, localeMessage.data)
  elelocale.value = localeMessage.data;
}
const setZh = async (lang: string, elelocale: any) => {
  let localeMessage = await axios({
    method: "get",
    url: "src/lang-server/zh-cn.json",
  });
  i18n.global.setLocaleMessage(lang, localeMessage.data);
  elelocale.value = localeMessage.data;
}
// 默认是英文
setEn('en', {});

export { i18n, setEn, setZh };
