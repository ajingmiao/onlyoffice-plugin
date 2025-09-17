// services/editor-service.js
export class EditorService {
  constructor(pluginBridge) {
    this.plugin = pluginBridge;
  }

  runInDoc(fn, { async = false, cb, scope } = {}) {
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
