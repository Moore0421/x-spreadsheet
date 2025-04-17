import { tf } from "../locale/locale";

// 检查值是否为有效数字
const isValidNumber = (v) => {
  if (v === null || v === undefined || v === '') return false;
  return !isNaN(Number(v));
};

// 通用数字格式化函数
const formatNumber = (v, options = {}) => {
  if (!isValidNumber(v)) return v;
  try {
    return new Intl.NumberFormat(undefined, options).format(Number(v));
  } catch (e) {
    return v;
  }
};

// 常规格式化函数
const formatGeneral = (v) => {
  if (isValidNumber(v)) {
    // 如果是小数，保留两位小数
    const num = Number(v);
    return Number.isInteger(num) ? num.toString() : num.toFixed(2);
  }
  return v ? v.toString() : '';
};

// 数值格式化函数
const formatNumberRender = (v) => {
  if (!isValidNumber(v)) return v;
  return formatNumber(v, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
};

// 百分比格式化函数 (将数值乘以100并加上%)
const formatPercent = (v) => {
  if (!isValidNumber(v)) return v;
  return formatNumber(Number(v) * 100, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + '%';
};

// 千分比格式化函数 (将数值乘以1000并加上‰)
const formatPermille = (v) => {
  if (!isValidNumber(v)) return v;
  return formatNumber(Number(v) * 1000, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + '‰';
};

// 货币格式化函数
const formatCurrency = (v, currency, showSymbol = false) => {
  if (!isValidNumber(v)) return v;
  try {
    if (!showSymbol) {
      // 不显示货币符号，直接使用数字格式
      return formatNumber(Number(v), {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    }
    return new Intl.NumberFormat(undefined, {
      style: 'currency',
      currency: currency,
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(Number(v));
  } catch (e) {
    // 如果出错，回退到简单的格式
    return formatNumberRender(v);
  }
};

// 日期格式化函数
const formatDate = (v, options = {}) => {
  if (!v) return '';
  try {
    const date = new Date(v);
    if (isNaN(date.getTime())) return v;
    return new Intl.DateTimeFormat(undefined, options).format(date);
  } catch (e) {
    return v;
  }
};

const baseFormats = [
  {
    key: "general",
    title: tf("format.general"),
    type: "string",
    render: formatGeneral,
  },
  {
    key: "number",
    title: tf("format.number"),
    type: "number",
    render: formatNumberRender,
  },
  {
    key: "text",
    title: tf("format.text"),
    type: "string",
    render: (v) => v ? v.toString() : '',
  },
  {
    key: "percent",
    title: tf("format.percent"),
    type: "percent",
    render: formatPercent,
  },
  {
    key: "permille",
    title: tf("format.permille"),
    type: "permille",
    render: formatPermille,
  },
  {
    key: "rmb",
    title: tf("format.rmb"),
    type: "number",
    render: (v) => formatCurrency(v, 'CNY'),
  },
  {
    key: "usd",
    title: tf("format.usd"),
    type: "number",
    render: (v) => formatCurrency(v, 'USD'),
  },
  {
    key: "eur",
    title: tf("format.eur"),
    type: "number",
    render: (v) => formatCurrency(v, 'EUR'),
  },
  {
    key: "date",
    title: tf("format.date"),
    type: "date",
    render: (v) => formatDate(v, {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    }),
  },
  {
    key: "time",
    title: tf("format.time"),
    type: "date",
    render: (v) => formatDate(v, {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }),
  },
  {
    key: "datetime",
    title: tf("format.datetime"),
    type: "date",
    render: (v) => formatDate(v, {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }),
  },
  {
    key: "duration",
    title: tf("format.duration"),
    type: "date",
    render: (v) => {
      if (!v) return '';
      try {
        const date = new Date(v);
        if (isNaN(date.getTime())) return v;
        // 只显示时间部分
        const hours = date.getHours().toString().padStart(2, '0');
        const minutes = date.getMinutes().toString().padStart(2, '0');
        const seconds = date.getSeconds().toString().padStart(2, '0');
        return `${hours}:${minutes}:${seconds}`;
      } catch (e) {
        return v;
      }
    },
  },
];

const formatm = {};
baseFormats.forEach((f) => {
  formatm[f.key] = f;
});

export default {};
export { formatm, baseFormats };
