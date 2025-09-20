// services/editor-service.js
export class EditorService {
  constructor(pluginBridge) {
    this.plugin = pluginBridge;
  }

  runInDoc(fn, options = {}) {
    // 支持两种调用方式：
    // 1. runInDoc(fn, callback) - 旧式回调
    // 2. runInDoc(fn, { async, cb, scope }) - 新式选项对象

    if (typeof options === 'function') {
      // 旧式回调方式：runInDoc(fn, callback)
      return this.plugin.callCommand(fn, false, options);
    }

    // 新式选项对象方式
    const { async = false, cb, scope } = options;

    // 把参数挂到插件侧的 Asc.scope（会被文档端读取）
    if (scope) {
      try {
        window.Asc = window.Asc || {};
        window.Asc.scope = Object.assign({}, window.Asc.scope || {}, scope);
      } catch (e) {}
    }

    return this.plugin.callCommand(fn, async, cb);
  }
}
