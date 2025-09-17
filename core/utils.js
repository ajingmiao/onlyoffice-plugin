export function tryWrap(fn, fallback = null) {
    try { return fn(); } catch (e) { return fallback; }
  }
  
  export function isFn(f) { return typeof f === 'function'; }
  
// core/utils.js

/**
 * 函数防抖：在指定延迟内多次触发，只执行最后一次
 * @param {Function} fn - 需要防抖的函数
 * @param {number} ms - 延迟时间（毫秒）
 */
export const debounce = (fn, ms = 100) => {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
};

/**
 * 函数节流：在指定间隔内只执行一次
 * @param {Function} fn - 需要节流的函数
 * @param {number} ms - 间隔时间（毫秒）
 */
export const throttle = (fn, ms = 100) => {
  let last = 0;
  return (...args) => {
    const now = Date.now();
    if (now - last >= ms) {
      last = now;
      fn(...args);
    }
  };
};

/**
 * 安全 JSON 解析
 * @param {string} v - JSON 字符串
 * @param {*} fallback - 解析失败时返回的默认值
 */
export const safeJson = (v, fallback = null) => {
  try {
    return JSON.parse(v);
  } catch {
    return fallback;
  }
};

/**
 * 颜色转 RGB 数组
 * @param {string|Array|undefined} c - hex 字符串 / RGB 数组 / undefined
 * @returns {Array} [r,g,b]
 */
export const toRgb = (c) => {
  if (!c) return [25, 118, 210]; // 默认蓝色 #1976D2
  if (Array.isArray(c) && c.length === 3) return c;
  if (typeof c === 'string' && /^#?[0-9a-fA-F]{6}$/.test(c)) {
    const hex = c[0] === '#' ? c.slice(1) : c;
    return [
      parseInt(hex.slice(0, 2), 16),
      parseInt(hex.slice(2, 4), 16),
      parseInt(hex.slice(4, 6), 16)
    ];
  }
  return [25, 118, 210];
};

