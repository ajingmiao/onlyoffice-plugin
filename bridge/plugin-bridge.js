import { logger } from '../core/logger.js';
import { EditorService } from '../services/editor-service.js';

// =====================================================
// ChartBinding - 使用"图表指纹 + 文档自定义属性"实现图表⇄绑定数据的关联
// =====================================================
(function (global) {
  var TAG_PREFIX = 'chart-binding:';

  // --- 简易日志器 ---
  var _logger = (global.logger && typeof global.logger.info === 'function') ? global.logger : console;

  // --- 事件回调（可选） ---
  var _onEvent = null; // function(payload) {}
  var _editorService = null; // EditorService 实例

  function safeCb(payload) {
    try {
      _logger.info('🔄 safeCb called with payload:', payload);
      if (typeof _onEvent === 'function') {
        _logger.info('✅ _onEvent is available, calling...');
        _onEvent(payload);
      } else {
        _logger.warn('❌ _onEvent is not a function:', typeof _onEvent);
      }
    } catch (e) {
      try { _logger.error('[cb error]', e && e.message ? e.message : e); } catch (_) {}
    }
  }

  // ===========================
  // 安装：设置 logger / 回调
  // ===========================
  function install(opts) {
    opts = opts || {};
    if (opts.logger && typeof opts.logger.info === 'function') {
      _logger = opts.logger;
    }
    if (typeof opts.onEvent === 'function') {
      _onEvent = opts.onEvent;
    }

    // 初始化 EditorService
    if (!_editorService) {
      _editorService = new EditorService({
        callCommand: function(fn, isAsync, callback) {
          return global.Asc?.plugin?.callCommand?.(fn, isAsync, callback);
        }
      });
    }

    // 注册点击事件（OnlyOffice 编辑器内部 onClick）
    global.Asc = global.Asc || {};
    global.Asc.plugin = global.Asc.plugin || {};
    global.Asc.plugin.event_onClick = handleOnClick;
  }

  // ===========================
  // 图表检测工具函数
  // ===========================
  function getChartType(sel) {
    try {
      var ch = (sel && typeof sel.GetChart === 'function') ? sel.GetChart() : null;
      if (ch && typeof ch.GetChartType === 'function') return ch.GetChartType();
      if (sel && typeof sel.GetChartType === 'function') return sel.GetChartType();
    } catch (e) {
      console.log('getChartType失败:', e);
    }
    return null;
  }

  function findHostParagraph(el, maxHop) {
    var hop = 0, cur = el;
    while (cur && hop < (maxHop || 24)) {
      if (cur.GetClassType && cur.GetClassType() === 'paragraph') return cur;
      if (!cur.GetParent) break;
      try { cur = cur.GetParent(); } catch (e) { break; }
      hop++;
    }
    return null;
  }

  function collectAncestorChain(el, limit) {
    var chain = [], hop = 0, cur = el;
    while (cur && hop < (limit || 8)) {
      chain.push(cur.GetClassType ? cur.GetClassType() : '?');
      if (!cur.GetParent) break;
      try { cur = cur.GetParent(); } catch (e) { break; }
      hop++;
    }
    return chain;
  }

  // 核心：生成图表指纹
  function buildChartFingerprint(sel, doc) {

    // 1) 优先稳定 ID
    var stableId = null;
    try { if (sel && typeof sel.GetId === 'function') stableId = sel.GetId(); } catch(e){ console.log('GetId失败:', e); }
    try { if (!stableId && sel && typeof sel.GetInnerId === 'function') stableId = sel.GetInnerId(); } catch(e){ console.log('GetInnerId失败:', e); }
    try { if (!stableId && sel && typeof sel.GetRId === 'function') stableId = sel.GetRId(); } catch(e){ console.log('GetRId失败:', e); }
    if (stableId) {
      console.log('✅ 使用稳定ID:', stableId);
      return 'id:' + stableId;
    }

    var parts = [];

    // 2.1 类型
    var type = getChartType(sel);
    if (type) parts.push('type:' + type);

    // 2.2 祖先链（前3级）
    try {
      var anc = collectAncestorChain(sel, 8);
      if (anc && anc.length) parts.push('chain:' + anc.slice(0,3).join('-'));
    } catch(_){ console.log('祖先链失败'); }

    // 2.3 宿主段落文本哈希（前20字符）
    try {
      var host = findHostParagraph(sel, 24);
      if (host && typeof host.GetText === 'function') {
        var t = '';
        try { t = host.GetText() || ''; } catch(_){}
        if (t) parts.push('text:' + t.substring(0,20).replace(/\s+/g, '_'));
      }
    } catch(_){ console.log('宿主段落失败'); }

    // 2.4 在全部绘图对象里的大致序位百分比
    try {
      var all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
      if (all && all.length) {
        for (var i=0; i<all.length; i++) {
          if (all[i] === sel) {
            var p = Math.floor((i / all.length) * 100);
            parts.push('pos:' + p);
            break;
          }
        }
      }
    } catch(_){ console.log('绘图对象位置失败'); }

    if (!parts.length) parts.push('rand:' + Date.now().toString(36));
    var fingerprint = parts.join('|');
    console.log('🔖 生成指纹:', fingerprint);
    return fingerprint;
  }

  // 文档自定义属性读写
  function getBindingByFingerprint(doc, fp) {
    try {
      var props = doc.GetCustomProperties();
      var key = 'chart-binding:' + fp;
      var val = props.Get(key);
      if (!val) return null;
      try { return JSON.parse(val); } catch(_){ return null; }
    } catch(e){ console.log('读取绑定失败:', e); return null; }
  }

  function setBindingByFingerprint(doc, fp, meta) {
    try {
      var props = doc.GetCustomProperties();
      var key = 'chart-binding:' + fp;
      props.Add(key, JSON.stringify(meta)); // 存在同名时覆盖
    } catch(e){ console.log('写入绑定失败:', e); }
  }

  // 图表检测主函数
  function detectAndBindChart() {
    console.log('🏁 图表检测开始执行');

    // 从 Asc.scope 中获取参数
    var isSelectionUse = false;
    try {
      isSelectionUse = !!(Asc && Asc.scope && Asc.scope.preferDrawing);
    } catch(e) {}

    // 在沙箱内部重新定义工具函数
    function getChartType(sel) {
      try {
        var ch = (sel && typeof sel.GetChart === 'function') ? sel.GetChart() : null;
        if (ch && typeof ch.GetChartType === 'function') return ch.GetChartType();
        if (sel && typeof sel.GetChartType === 'function') return sel.GetChartType();
      } catch (e) {
        console.log('getChartType失败:', e);
      }
      return null;
    }

    function findHostParagraph(el, maxHop) {
      var hop = 0, cur = el;
      while (cur && hop < (maxHop || 24)) {
        if (cur.GetClassType && cur.GetClassType() === 'paragraph') return cur;
        if (!cur.GetParent) break;
        try { cur = cur.GetParent(); } catch (e) { break; }
        hop++;
      }
      return null;
    }

    function collectAncestorChain(el, limit) {
      var chain = [], hop = 0, cur = el;
      while (cur && hop < (limit || 8)) {
        chain.push(cur.GetClassType ? cur.GetClassType() : '?');
        if (!cur.GetParent) break;
        try { cur = cur.GetParent(); } catch (e) { break; }
        hop++;
      }
      return chain;
    }

    function buildChartFingerprint(sel, doc) {
      // 1) 优先稳定 ID
      var stableId = null;
      try { if (sel && typeof sel.GetId === 'function') stableId = sel.GetId(); } catch(e){ console.log('GetId失败:', e); }
      try { if (!stableId && sel && typeof sel.GetInnerId === 'function') stableId = sel.GetInnerId(); } catch(e){ console.log('GetInnerId失败:', e); }
      try { if (!stableId && sel && typeof sel.GetRId === 'function') stableId = sel.GetRId(); } catch(e){ console.log('GetRId失败:', e); }
      if (stableId) {
        console.log('✅ 使用稳定ID:', stableId);
        return 'id:' + stableId;
      }

      var parts = [];

      // 2.1 类型
      var type = getChartType(sel);
      if (type) parts.push('type:' + type);

      // 2.2 祖先链（前3级）
      try {
        var anc = collectAncestorChain(sel, 8);
        if (anc && anc.length) parts.push('chain:' + anc.slice(0,3).join('-'));
      } catch(_){ console.log('祖先链失败'); }

      // 2.3 宿主段落文本哈希（前20字符）
      try {
        var host = findHostParagraph(sel, 24);
        if (host && typeof host.GetText === 'function') {
          var t = '';
          try { t = host.GetText() || ''; } catch(_){}
          if (t) parts.push('text:' + t.substring(0,20).replace(/\s+/g, '_'));
        }
      } catch(_){ console.log('宿主段落失败'); }

      // 2.4 在全部绘图对象里的大致序位百分比
      try {
        var all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
        if (all && all.length) {
          for (var i=0; i<all.length; i++) {
            if (all[i] === sel) {
              var p = Math.floor((i / all.length) * 100);
              parts.push('pos:' + p);
              break;
            }
          }
        }
      } catch(_){ console.log('绘图对象位置失败'); }

      if (!parts.length) parts.push('rand:' + Date.now().toString(36));
      var fingerprint = parts.join('|');
      console.log('🔖 生成指纹:', fingerprint);
      return fingerprint;
    }

    function getBindingByFingerprint(doc, fp) {
      try {
        var props = doc.GetCustomProperties();
        var key = 'chart-binding:' + fp;
        var val = props.Get(key);
        if (!val) return null;
        try { return JSON.parse(val); } catch(_){ return null; }
      } catch(e){ console.log('读取绑定失败:', e); return null; }
    }

    function setBindingByFingerprint(doc, fp, meta) {
      try {
        var props = doc.GetCustomProperties();
        var key = 'chart-binding:' + fp;
        props.Add(key, JSON.stringify(meta)); // 存在同名时覆盖
      } catch(e){ console.log('写入绑定失败:', e); }
    }

    var doc = Api.GetDocument();
    if (!doc) {
      console.log('❌ 无法获取文档对象');
      return { ok: false, action: 'error', message: '无法获取文档对象' };
    }

    // 先看是否选中图表
    console.log('🎯 检查选中的绘图对象...');
    var selected = null;
    try {
      selected = doc.GetSelectedDrawings ? doc.GetSelectedDrawings() : null;
    } catch(e){
      console.log('GetSelectedDrawings失败:', e);
    }

    if (!selected || !selected.length) {
      // 没选中图表 → 再看文本选区（避免误触）
      var range = null;
      try { range = doc.GetRangeBySelect && doc.GetRangeBySelect(); } catch(_){}
      if (range) {
        return { ok: true, action: 'text-click', message: '点击文本区域，跳过图表绑定' };
      }
      return { ok: true, action: 'no-user-chart', message: '未检测到用户选中的图表' };
    }

    console.log('✅ 发现选中的图表，开始处理...');
    var sel = selected[selected.length - 1];
    var fp = buildChartFingerprint(sel, doc);

    console.log('🔍 检查是否已存在绑定...');
    var existed = getBindingByFingerprint(doc, fp);
    if (existed) {
      console.log('✅ 已存在绑定，数据详情:', JSON.stringify(existed, null, 2));
      console.log('📊 绑定数据解析:');
      console.log('  - 图表ID:', existed.chartId);
      console.log('  - 创建时间:', existed.createdAt);
      console.log('  - 图表类型:', existed.chartType);
      console.log('  - 指纹:', existed.fingerprint);
      if (existed.datasetId) console.log('  - 数据集ID:', existed.datasetId);
      if (existed.mapping) console.log('  - 数据映射:', JSON.stringify(existed.mapping));

      return {
        ok: true,
        action: 'exists',
        meta: existed,
        fingerprint: fp,
        message: '已存在图表绑定（自定义属性）'
      };
    }

    // 新建并写入
    console.log('🆕 创建新的绑定...');
    var meta = {
      chartId: 'chart-' + Date.now().toString(36) + '-' + Math.floor(Math.random()*1e6).toString(36),
      createdAt: new Date().toISOString(),
      chartType: getChartType(sel),
      fingerprint: fp
    };

    setBindingByFingerprint(doc, fp, meta);
    console.log('✅ 图表检测执行完成，新建绑定数据详情:', JSON.stringify(meta, null, 2));
    console.log('📊 新建绑定数据解析:');
    console.log('  - 图表ID:', meta.chartId);
    console.log('  - 创建时间:', meta.createdAt);
    console.log('  - 图表类型:', meta.chartType);
    console.log('  - 指纹:', meta.fingerprint);

    return {
      ok: true,
      action: 'created',
      meta: meta,
      fingerprint: fp,
      message: '已写入绑定到文档自定义属性'
    };
  }

  // ===========================
  // 点击事件主处理
  // ===========================
  function handleOnClick(isSelectionUse) {
    try { _logger.info('[onClick] triggered', { isSelectionUse: !!isSelectionUse }); } catch (_){}

    // 1) 先看当前是否点到内容控件（兼容旧锚点/其它 CC）
    try {
      _logger.info('🔍 开始执行 GetCurrentContentControlPr...');
      global.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function (ccPr) {
        try {
          _logger.info('📋 GetCurrentContentControlPr 回调执行，ccPr:', ccPr);

          if (ccPr && typeof ccPr.Tag === 'string' && ccPr.Tag.indexOf(TAG_PREFIX) === 0) {
            // 命中旧锚点 SDT（Tag 里就是 JSON 元数据）
            try {
              var anchorMeta = JSON.parse(ccPr.Tag.slice(TAG_PREFIX.length));
              _logger.info('🎯 命中旧版图表锚点SDT', anchorMeta);
              safeCb({
                type: 'chart-anchor-clicked',
                meta: anchorMeta,
                anchor: {
                  tag: ccPr.Tag,
                  alias: ccPr.Alias || '',
                  internalId: ccPr.InternalId || ''
                }
              });
            } catch (parseErr) {
              _logger.warn('旧锚点解析失败:', parseErr && parseErr.message ? parseErr.message : parseErr);
              safeCb({ type: 'error', message: '旧锚点数据解析失败' });
            }
            return;
          }

          if (ccPr && ccPr.Tag) {
            // 其它内容控件
            _logger.info('🏷️ 内容控件点击:', ccPr.Tag);
            safeCb({
              type: 'content-control-clicked',
              tag: ccPr.Tag,
              alias: ccPr.Alias || '',
              internalId: ccPr.InternalId || '',
              appearance: ccPr.Appearance || 0
            });
            return;
          }

        } catch (ccError) {
          _logger.error('处理内容控件时出错:', ccError);
        }

        // 2) 未命中任何 CC → 使用新的简洁图表检测方式
        _logger.info('🔍 准备开始图表检测流程...');
        try {
          _logger.info('🔍 开始图表检测...');

          if (!_editorService) {
            _logger.error('❌ EditorService 未初始化');
            safeCb({ type: 'error', message: 'EditorService 未初始化' });
            return;
          }

          _logger.info('✅ EditorService 可用，准备执行图表检测...');

          // 使用全局回调方式执行图表检测
          _logger.info('🚀 准备执行图表检测（全局回调方式）...');

          // 设置全局回调函数
          global._chartDetectionCallback = function(result) {
            try {
              _logger.info('🎉 全局回调被触发，结果:', result);

              if (result && result.ok) {
                // 根据检测结果触发相应事件
                if (result.action === 'exists' && result.meta) {
                  _logger.info('📊 触发 chart-binding-exists 事件');
                  safeCb({
                    type: 'chart-binding-exists',
                    meta: result.meta,
                    fingerprint: result.fingerprint
                  });
                } else if (result.action === 'created' && result.meta) {
                  _logger.info('📊 触发 chart-binding-created 事件');
                  safeCb({
                    type: 'chart-binding-created',
                    meta: result.meta,
                    fingerprint: result.fingerprint
                  });
                } else if (result.action === 'text-click') {
                  _logger.info('📝 检测到文本点击');
                  safeCb({
                    type: 'text-click',
                    message: result.message
                  });
                } else if (result.action === 'no-user-chart') {
                  _logger.info('⚠️ 未检测到选中的图表');
                  safeCb({
                    type: 'no-user-chart',
                    message: result.message
                  });
                } else {
                  safeCb({
                    type: 'chart-detection-info',
                    message: result.message || '图表检测完成',
                    action: result.action
                  });
                }
              } else {
                _logger.warn('图表检测失败或无结果');
                safeCb({ type: 'error', message: result?.message || '图表检测失败' });
              }
            } catch (callbackError) {
              _logger.error('🚨 全局回调执行出错:', callbackError);
            }
          };

          _logger.info('📝 全局回调函数已设置');

          // 创建修改后的函数字符串，在最后调用全局回调
          const funcStr = `
            ${detectAndBindChart.toString()}

            // 执行检测并获取结果
            var result = detectAndBindChart();

            // 调用全局回调
            if (typeof window._chartDetectionCallback === 'function') {
              window._chartDetectionCallback(result);
            }

            return result;
          `;

          _logger.info('🔧 函数字符串已生成，准备调用callCommand...');

          // 直接使用callCommand执行修改后的函数
          global.Asc = global.Asc || {};
          global.Asc.scope = { preferDrawing: !!isSelectionUse };

          _logger.info('⚡ 正在调用 callCommand...');
          global.Asc.plugin.callCommand(new Function(funcStr), false);
          _logger.info('✅ callCommand 调用完成');

        } catch (detectionError) {
          _logger.error('🚨 图表检测流程出错:', detectionError);
          safeCb({ type: 'error', message: '图表检测失败: ' + detectionError.message });
        }
    }); // GetCurrentContentControlPr 回调结束
    } catch (executeError) {
      _logger.error('🚨 executeMethod 调用失败:', executeError);
      safeCb({ type: 'error', message: 'executeMethod 调用失败: ' + executeError.message });
    }

    _logger.info('✅ handleOnClick 函数执行完成');
  }

  // ===========================
  // 工具：列出所有绑定（自定义属性中以 chart-binding: 开头的）
  // 返回：[{ key, meta }, ...]
  // ===========================
  function listBindings(cb) {
    var funcStr = (function(){/*
      var out = { ok: true, bindings: [], logs: [] };
      function dbg(){ try { out.logs.push(Array.prototype.join.call(arguments, ' ')); } catch (_e){} }
      var doc = null;
      try { doc = Api.GetDocument(); } catch(e){}
      if (!doc) { out.ok = false; out.message = '无法获取文档对象'; return out; }
      try {
        var props = doc.GetCustomProperties();
        // API 没有直接"列举全部"的方法时，这里无法遍历；
        // 若当前版本 props 支持枚举（不同版本行为不同），你可以替换为真实枚举逻辑。
        // 兼容策略：维护一个索引 key（chart-binding:index），里面存 fingerprint 数组。
        var idx = null;
        try { idx = props.Get('chart-binding:index'); } catch(e){}
        if (idx) {
          try {
            var fps = JSON.parse(idx);
            for (var i=0;i<fps.length;i++){
              var key = 'chart-binding:' + fps[i];
              var val = props.Get(key);
              if (!val) continue;
              try { out.bindings.push({ key:key, meta: JSON.parse(val) }); } catch(_){}
            }
          } catch(e){}
        }
      } catch (e) {
        out.ok = false; out.message = e && e.message ? e.message : String(e);
      }
      return out;
    */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

    global.Asc.plugin.callCommand(new Function(funcStr), function (info) {
      if (!info || info.ok === false) {
        cb && cb({ ok:false, message: info && info.message });
        return;
      }
      cb && cb({ ok:true, bindings: info.bindings || [] });
    });
  }

  // ===========================
  // 工具：为"当前选中图表"写入/更新绑定（手动 Upsert）
  // payload: 你的业务数据（对象），将合并写入 meta
  // ===========================
  function upsertBindingForSelectedChart(payload, cb) {
    payload = payload || {};
    var funcStr = (function(){/*
      var out = { ok:true, action:'none', meta:null, fingerprint:null, message:'', logs:[] };
      function dbg(){ try { out.logs.push(Array.prototype.join.call(arguments, ' ')); } catch (_e){} }

      function getDoc(){ try { return Api.GetDocument(); } catch(e){ return null; } }
      function getChartType(sel){
        try{
          var ch = (sel && typeof sel.GetChart === 'function') ? sel.GetChart() : null;
          if (ch && typeof ch.GetChartType === 'function') return ch.GetChartType();
          if (sel && typeof sel.GetChartType === 'function') return sel.GetChartType();
        }catch(e){}
        return null;
      }
      function findHostParagraph(el, maxHop){
        var hop=0, cur=el;
        while(cur && hop < (maxHop||24)){
          if (cur.GetClassType && cur.GetClassType()==='paragraph') return cur;
          if (!cur.GetParent) break;
          try{ cur = cur.GetParent(); }catch{ break; }
          hop++;
        }
        return null;
      }
      function collectAncestorChain(el, limit){
        var chain=[], hop=0, cur=el;
        while(cur && hop < (limit||8)){
          chain.push(cur.GetClassType?cur.GetClassType():'?');
          if(!cur.GetParent) break;
          try{ cur = cur.GetParent(); }catch{ break; }
          hop++;
        }
        return chain;
      }
      function buildChartFingerprint(sel, doc){
        var stableId=null;
        try{ if(sel && typeof sel.GetId==='function') stableId = sel.GetId(); }catch(e){}
        try{ if(!stableId && sel && typeof sel.GetInnerId==='function') stableId = sel.GetInnerId(); }catch(e){}
        try{ if(!stableId && sel && typeof sel.GetRId==='function') stableId = sel.GetRId(); }catch(e){}
        if(stableId) return 'id:'+stableId;

        var parts=[];
        var type=getChartType(sel); if(type) parts.push('type:'+type);
        try{ var anc=collectAncestorChain(sel,8); if(anc&&anc.length) parts.push('chain:'+anc.slice(0,3).join('-')); }catch(_){}
        try{
          var host=findHostParagraph(sel,24);
          if(host && typeof host.GetText==='function'){
            var t=''; try{ t=host.GetText()||''; }catch(_){}
            if(t) parts.push('text:'+t.substring(0,20).replace(/\s+/g, '_'));
          }
        }catch(_){}
        try{
          var all=doc.GetAllDrawingObjects?doc.GetAllDrawingObjects():null;
          if(all && all.length){
            for(var i=0;i<all.length;i++){
              if(all[i]===sel){
                var p=Math.floor((i/all.length)*100);
                parts.push('pos:'+p); break;
              }
            }
          }
        }catch(_){}
        if(!parts.length) parts.push('rand:'+Date.now().toString(36));
        return parts.join('|');
      }

      var doc=getDoc();
      if(!doc){ out.ok=false; out.message='无法获取文档对象'; return out; }
      var selected=null;
      try{ selected = doc.GetSelectedDrawings ? doc.GetSelectedDrawings() : null; }catch(_){}
      if(!selected || !selected.length){ out.ok=false; out.message='未选中图表'; return out; }
      var sel = selected[selected.length-1];

      var fp=buildChartFingerprint(sel, doc);
      out.fingerprint=fp;

      var props=doc.GetCustomProperties();
      var key='chart-binding:'+fp;
      var old=null;
      try{ old = props.Get(key); }catch(_){}
      var meta=null;
      if(old){
        try{ meta=JSON.parse(old); }catch(_){ meta=null; }
      }
      if(!meta){
        meta={
          chartId:'chart-'+Date.now().toString(36)+'-'+Math.floor(Math.random()*1e6).toString(36),
          createdAt:new Date().toISOString(),
          chartType:getChartType(sel),
          fingerprint:fp
        };
        out.action='created';
      }else{
        out.action='updated';
      }

      // 合并业务数据
      var payload = Asc && Asc.scope && Asc.scope._cb_payload ? Asc.scope._cb_payload : null;
      if (payload && typeof payload === 'object') {
        for (var k in payload) { if (Object.prototype.hasOwnProperty.call(payload, k)) meta[k]=payload[k]; }
      }

      props.Add(key, JSON.stringify(meta));
      out.meta=meta;
      out.message = (out.action==='created'?'已创建':'已更新') + '绑定';
      return out;
    */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

    try { global.Asc.scope = global.Asc.scope || {}; global.Asc.scope._cb_payload = payload; } catch(_){}

    global.Asc.plugin.callCommand(new Function(funcStr), function(info){
      if (!info || info.ok === false) { cb && cb({ ok:false, message: info && info.message }); return; }
      cb && cb({ ok:true, action: info.action, meta: info.meta, fingerprint: info.fingerprint });
    });
  }

  // ===========================
  // 工具：按 fingerprint "软删除"绑定（设为空串）
  // （如你确认 API 支持真正 Remove，可改为 Remove(key)）
  // ===========================
  function removeBindingByFingerprint(fingerprint, cb) {
    if (!fingerprint) { cb && cb({ ok:false, message:'fingerprint 为空' }); return; }

    var funcStr = (function(){/*
      var out = { ok:true, message:'' };
      var doc=null; try{ doc=Api.GetDocument(); }catch(e){}
      if(!doc){ out.ok=false; out.message='无法获取文档对象'; return out; }
      var props = doc.GetCustomProperties();
      var key = 'chart-binding:' + (Asc && Asc.scope && Asc.scope._fp ? Asc.scope._fp : '');
      try {
        props.Add(key, ''); // 软删除：置空串
        out.message='已置空该绑定（软删除）';
      } catch(e){
        out.ok=false; out.message=e && e.message ? e.message : String(e);
      }
      return out;
    */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

    try { global.Asc.scope = global.Asc.scope || {}; global.Asc.scope._fp = fingerprint; } catch(_){}

    global.Asc.plugin.callCommand(new Function(funcStr), function(info){
      if (!info || info.ok === false) { cb && cb({ ok:false, message: info && info.message }); return; }
      cb && cb({ ok:true, message: info.message });
    });
  }

  // 导出
  var API = {
    install: install,
    listBindings: listBindings,
    upsertBindingForSelectedChart: upsertBindingForSelectedChart,
    removeBindingByFingerprint: removeBindingByFingerprint
  };

  global.ChartBinding = API;
})(window);

export class PluginBridge {
  onInit(cb) {
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};
    window.Asc.plugin.init = () => {
      logger.info('plugin init');
      cb?.();
    };
  }

  callCommand(fn, isAsync = false, callback) {
    return window.Asc?.plugin?.callCommand?.(fn, isAsync, callback);
  }

  onSelectionChanged(cb) {
    // 初始化 ChartBinding 系统，但保留原有的选择变化逻辑
    window.ChartBinding.install({
      logger: logger,
      onEvent: function(payload) {
        logger.info('ChartBinding event:', payload);

        // 将 ChartBinding 事件转换为兼容的回调格式
        if (payload.type === 'chart-binding-exists') {
          cb?.({
            type: 'chart-anchor-exists',
            meta: payload.meta,
            fingerprint: payload.fingerprint
          });
          // 图表绑定存在时也触发选择变化检测
          logger.info('Chart binding exists, triggering selection change detection');
          setTimeout(() => cb?.(), 10);
        } else if (payload.type === 'chart-binding-created') {
          cb?.({
            type: 'chart-anchor-created',
            meta: payload.meta,
            fingerprint: payload.fingerprint
          });
          // 图表绑定创建后也触发选择变化检测
          logger.info('Chart binding created, triggering selection change detection');
          setTimeout(() => cb?.(), 10);
        } else if (payload.type === 'content-control-clicked') {
          cb?.({
            tag: payload.tag,
            alias: payload.alias,
            internalId: payload.internalId,
            appearance: payload.appearance
          });
        } else if (payload.type === 'chart-anchor-clicked') {
          cb?.({
            type: 'chart-anchor-clicked',
            meta: payload.meta,
            anchor: payload.anchor
          });
          // 图表锚点点击后也触发选择变化检测
          logger.info('Chart anchor clicked, triggering selection change detection');
          setTimeout(() => cb?.(), 10);
        } else if (payload.type === 'no-user-chart') {
          cb?.({
            type: 'no-user-chart',
            message: payload.message
          });
          // 即使没有检测到用户图表，也触发选择变化检测以便其他检测逻辑执行
          logger.info('No user chart detected, triggering selection change detection');
          setTimeout(() => cb?.(), 10);
        } else if (payload.type === 'text-click') {
          // 文本点击时，仍然触发选择变化检测
          logger.info('Text click detected, triggering selection change detection');
          cb?.();
        } else if (payload.type === 'error') {
          cb?.({
            type: 'error',
            message: payload.message
          });
        } else {
          // 其他事件直接传递
          cb?.(payload);
        }
      }
    });

    // 保留原有的 onSelectionChanged 事件处理（用于光标移动等常规选择变化）
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};
    window.Asc.plugin.event_onSelectionChanged = () => {
      logger.info('Selection changed event triggered - running standard detection');
      cb?.();
    };

    // 如果有 connector API，也使用它
    if (window.connector && typeof window.connector.attachEvent === 'function') {
      logger.info('Using connector.attachEvent for enhanced selection tracking');

      // 监听内容控件变化事件
      window.connector.attachEvent("onChangeContentControl", (obj) => {
        logger.info('Content control changed via connector:', obj);
        cb?.();
      });

      // 监听目标位置变化（光标移动）
      window.connector.attachEvent("onTargetPositionChanged", (obj) => {
        logger.info('Target position changed via connector:', obj);
        cb?.();
      });
    }

    logger.info('✅ ChartBinding system initialized and selection change handlers configured');
  }

  onHyperLinkClick(cb) {
    window.Asc = window.Asc || {};
    window.Asc.plugin = window.Asc.plugin || {};
    window.Asc.plugin.event_onHyperLinkClick = (data) => {
      logger.info('hyperlink click detected:', data);
      cb?.(data);
    };
  }
}
