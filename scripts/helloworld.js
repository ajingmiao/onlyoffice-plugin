/* globals Api, Asc */

// —— 回父页：简易回执 —— 
function info(cmd, data) {
  try {
    if (window.parent?.Common?.Gateway) {
      window.parent.Common.Gateway.sendInfo({ command: cmd, data });
    }
  } catch (e) { }
}
 

function sendToHost(msg){
  try { window.parent?.Common?.Gateway?.sendInfo(msg); } catch(e){}
}


// 插件就绪
window.Asc.plugin.init = function () {
  console.log('[plugin] init');
  // 主动告诉宿主：我初始化完成，可以收指令了
  info('pluginInitialized', { guid: window.Asc.plugin.guid });
 
};


// 只保留 serviceCommand 通道
if (window.parent?.Common?.Gateway?.on) {
  window.parent.Common.Gateway.on('internalcommand', function (data) {
    if (data.command !== 'insertText') return;

    // 兼容字符串或 {text:'...'}
    var payload = data.data;
    var text = (typeof payload === 'string') ? payload : (payload && payload.text) || '';

    // ★ 关键：直接把参数写到 Asc.scope 上（不要依赖第三参）
    Asc.scope = Asc.scope || {};
    Asc.scope.text = text;

    window.Asc.plugin.callCommand(function () {

      // let doc = Api.GetDocument();
      // let paragraph = doc.GetElement(0);

      // // 创建行内内容控件
      // let inlineLvlSdt = Api.CreateInlineLvlSdt();

      // // 在 SDT 上附加数据
      // inlineLvlSdt.SetTag(JSON.stringify({
      //   id: "customerName",
      //   type: "string",
      //   required: true,
      //   source: "/api/customer/123"
      // }));
      // inlineLvlSdt.SetAlias("客户名称");

      // // 添加占位文本
      // let run = Api.CreateRun();
      // run.AddText("【插入占位符】");
      // inlineLvlSdt.AddElement(run, 0);

      // // 把控件放进段落
      // paragraph.AddInlineLvlSdt(inlineLvlSdt);
      // paragraph.Push(inlineLvlSdt);

      var doc = Api.GetDocument();

        // 1) 在光标位置添加一个内联组合框控件（行内）
        var sdt = doc.AddComboBoxContentControl();   // 或 doc.AddDropDownListContentControl()

        // 2) 自定义显示文字（作为占位提示）
        var content = sdt.GetContent();
        content.RemoveAllElements();
        var p = content.GetElement(0);
        p.AddText("{{customer_name}}");

        // 3) 设置标识/标题
        sdt.SetTag("bind:customer_name");
        sdt.SetTitle("客户姓名");



    }, false);

    // 回执（可选）
    try { window.parent.Common.Gateway.sendInfo({ command: 'pluginAck', data: { op: 'insertText', text } }); } catch (e) { }
  });
}


window.Asc.plugin.event_onSelectionChanged = function () {
  // 在插件 iframe 与父页都打日志，便于排查
  console.log('[plugin] onSelectionChanged fired');
  try { window.parent?.Common?.Gateway?.sendInfo({ command: 'onSelectionChanged-fired' }); } catch (e) { }

  window.Asc.plugin.callCommand(function () {
    var doc = Api.GetDocument();
    var range = doc.GetRangeBySelect();
    if (!range) return null;

    var ctrls = doc.GetAllContentControls();
    for (var i = 0; i < ctrls.length; i++) {
      var c = ctrls[i];
      var hit =
        (typeof range.IsInContentControl === 'function' && range.IsInContentControl(c)) ||
        (typeof c.IsRangeIn === 'function' && c.IsRangeIn(range)) ||
        (typeof c.IsSelected === 'function' && c.IsSelected());

      if (hit) {
        return {
          tag: c.GetTag ? c.GetTag() : "",
          alias: c.GetAlias ? c.GetAlias() : ""
        };
      }
    }
    return null;
  }, false, function (res) {
    // 打到插件控制台
    console.log('[plugin] active SDT:', res);

    // 同时回父页，让你在宿主页 Console 也能看到
    try { window.parent?.Common?.Gateway?.sendInfo({ command: 'activeSdt', data: res }); } catch (e) { }
  });
};