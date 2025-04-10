/* global window */
import zhCN from "./zh-cn";

// 修改默认语言设置
let $languages = ['zh-cn'];  // 使用正确的语言代码
const $messages = {
  'zh-cn': zhCN,  // 键名要和语言代码一致
};

function translate(key, messages) {
  if (messages) {
    // Return the translation from the first language in the languages array
    // that has a value for the provided key.
    for (const lang of $languages) {
      if (!messages[lang]) break;

      let message = messages[lang];

      // Splits the key at '.' except where escaped as '\.'
      const keys = key.match(/(?:\\.|[^.])+/g);

      for (let i = 0; i < keys.length; i += 1) {
        const property = keys[i];
        const value = message[property];

        // If value doesn't exist, try next language
        if (!value) break;

        if (i === keys.length - 1) return value;

        // Move down to the next level of the messages object
        message = value;
      }
    }
  }

  return undefined;
}

function t(key) {
  let v = translate(key, $messages);
  if (!v && window && window.x_spreadsheet && window.x_spreadsheet.$messages) {
    v = translate(key, window.x_spreadsheet.$messages);
  }
  return v || "";
}

function tf(key) {
  return () => t(key);
}

// If clearLangList is set to false, lang will be added to the front of the
// languages array. The languages in the language array are searched in order
// to find a translation. This allows the use of other languages as a fallback
// if lang is missing some keys. The language array is preloaded with English.
// To set the languages array to only include lang, set clearLangList to true.
function locale(lang, message, clearLangList = false) {
  // 确保语言代码统一使用小写
  const normalizedLang = lang.toLowerCase();
  
  if (clearLangList) {
    $languages = [normalizedLang];
  } else {
    // 避免重复添加
    if (!$languages.includes(normalizedLang)) {
      $languages.unshift(normalizedLang);
    }
  }

  if (message) {
    $messages[normalizedLang] = message;
  }
}

export default {
  t,
};

export { locale, t, tf };
