import { logger } from '../core/logger.js';

// =====================================================
// ChartBinding - 使用"图表指纹 + 文档自定义属性"实现图表⇄绑定数据的关联
// =====================================================
(function (global) {
  var TAG_PREFIX = 'chart-binding:';

  // --- 简易日志器 ---
  var _logger = (global.logger && typeof global.logger.info === 'function') ? global.logger : console;

  // --- 事件回调（可选） ---
  var _onEvent = null; // function(payload) {}

  function safeCb(payload) {
    try {
      if (typeof _onEvent === 'function') _onEvent(payload);
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
    // 注册点击事件（OnlyOffice 编辑器内部 onClick）
    global.Asc = global.Asc || {};
    global.Asc.plugin = global.Asc.plugin || {};
    global.Asc.plugin.event_onClick = handleOnClick;
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

          _logger.info('❌ 没有命中内容控件，开始图表检测...');
        } catch (ccError) {
          _logger.error('处理内容控件时出错:', ccError);
        }

        // 2) 未命中任何 CC → 进入沙箱：识别图表并进行绑定（文档自定义属性）
        try {
          _logger.info('🔍 准备执行图表检测沙箱代码...');
          global.Asc.scope = global.Asc.scope || {};
          global.Asc.scope.preferDrawing = !!isSelectionUse; // 仅作为参考
        } catch (_){}

        var funcStr = (function () {/*
        // =============== 沙箱开始 ===============
        console.log('🏁 沙箱代码开始执行');

        // 测试沙箱环境的能力
        console.log('🧪 测试沙箱环境:');
        console.log('  typeof window:', typeof window);
        console.log('  typeof global:', typeof global);
        console.log('  typeof this:', typeof this);

        var out = { ok: true, action: 'none', message: '', meta: null, fingerprint: null, logs: [] };
        function dbg(){ try { out.logs.push(Array.prototype.join.call(arguments, ' ')); } catch (_e){} }

        function getDoc() { try { return Api.GetDocument(); } catch(e){ console.log('getDoc失败:', e); return null; } }

        function getChartType(sel) {
          try {
            var ch = (sel && typeof sel.GetChart === 'function') ? sel.GetChart() : null;
            if (ch && typeof ch.GetChartType === 'function') return ch.GetChartType();
            if (sel && typeof sel.GetChartType === 'function') return sel.GetChartType();
          } catch (e) { console.log('getChartType失败:', e); }
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
          console.log('🔖 开始生成图表指纹');
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

        // 文档自定义属性读/写
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

        console.log('📄 获取文档对象...');
        var doc = getDoc();
        if (!doc) {
          console.log('❌ 无法获取文档对象');
          out.ok = false;
          out.action = 'error';
          out.message = '无法获取文档对象';
          return out;
        }

        // 先看是否选中图表
        console.log('🎯 检查选中的绘图对象...');
        var selected = null;
        try { selected = doc.GetSelectedDrawings ? doc.GetSelectedDrawings() : null; } catch(e){ console.log('GetSelectedDrawings失败:', e); }
        dbg('🎯 选中的绘图对象数量:', selected ? selected.length : 0);

        if (!selected || !selected.length) {
          // 没选中图表 → 再看文本选区（避免误触）
          var range = null;
          try { range = doc.GetRangeBySelect && doc.GetRangeBySelect(); } catch(_){}
          if (range) {
            out.action = 'text-click';
            out.message = '点击文本区域，跳过图表绑定';
            console.log('📝 检测到文本点击');
            return out;
          }
          out.action = 'no-user-chart';
          out.message = '未检测到用户选中的图表';
          console.log('⚠️ 未检测到选中的图表');
          return out;
        }

        console.log('✅ 发现选中的图表，开始处理...');
        var sel = selected[selected.length - 1];
        var fp = buildChartFingerprint(sel, doc);
        out.fingerprint = fp;
        dbg('🔖 指纹:', fp);

        console.log('🔍 检查是否已存在绑定...');
        var existed = getBindingByFingerprint(doc, fp);
        if (existed) {
          out.action = 'exists';
          out.meta = existed;
          out.message = '已存在图表绑定（自定义属性）';
          console.log('✅ 发现已存在的绑定');

          // 在沙箱内部直接通知外部
          try {
            console.log('🔗 尝试沙箱内部通知:');
            console.log('  typeof window:', typeof window);
            console.log('  window.ChartBindingNotify:', typeof (typeof window !== 'undefined' ? window.ChartBindingNotify : 'undefined'));

            if (typeof window !== 'undefined' && window.ChartBindingNotify) {
              window.ChartBindingNotify({
                type: 'chart-binding-exists',
                meta: existed,
                fingerprint: fp
              });
              console.log('✅ 已通过 ChartBindingNotify 发送存在事件');
            } else {
              console.log('❌ 沙箱内部无法访问 window.ChartBindingNotify');
              console.log('  window:', typeof window);
              console.log('  ChartBindingNotify:', typeof window !== 'undefined' ? typeof window.ChartBindingNotify : 'window不存在');
            }
          } catch (notifyError) {
            console.log('⚠️ ChartBindingNotify 调用失败:', notifyError);
          }
          console.log('已经绑定的返回数据对象...',out);
          return out;
        }

        // 新建并写入
        console.log('🆕 创建新的绑定...');
        var meta = {
          chartId: 'chart-' + Date.now().toString(36) + '-' + Math.floor(Math.random()*1e6).toString(36),
          createdAt: new Date().toISOString(),
          chartType: (function(){
            try {
              var c = (sel && typeof sel.GetChart === 'function') ? sel.GetChart() : null;
              if (c && typeof c.GetChartType === 'function') return c.GetChartType();
              if (sel && typeof sel.GetChartType === 'function') return sel.GetChartType();
            } catch(e){ console.log('获取图表类型失败:', e); }
            return null;
          })(),
          fingerprint: fp
          // 你可以在这里扩展你的业务字段，比如 bindingPayload / datasetId / mapping 等
        };
        setBindingByFingerprint(doc, fp, meta);
        console.log('沙箱中获取到的图表ID...',meta.chartId);
        console.log('沙箱中获取到的图表类型...',meta.chartType);

        // 由于 callCommand 回调不可靠，我们在沙箱内部直接触发事件
        out.action = 'created';
        out.meta = meta;
        out.message = '已写入绑定到文档自定义属性';
        console.log('✅ 沙箱代码执行完成，准备返回结果');

        // 在沙箱内部直接通知外部
        try {
          if (typeof window !== 'undefined' && window.ChartBindingNotify) {
            window.ChartBindingNotify({
              type: 'chart-binding-created',
              meta: meta,
              fingerprint: fp
            });
            console.log('✅ 已通过 ChartBindingNotify 发送创建事件');
          }
        } catch (notifyError) {
          console.log('⚠️ ChartBindingNotify 调用失败:', notifyError);
        }

        return out;
        // =============== 沙箱结束 ===============
      */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

      try {
        // 设置沙箱内部可以直接调用的通知函数
        window.ChartBindingNotify = function(payload) {
          _logger.info('📨 沙箱内部通知:', payload);
          safeCb(payload);
        };

        // 先测试沙箱代码是否有语法问题
        try {
          var testFunc = new Function(funcStr);
          _logger.info('✅ 沙箱代码语法检查通过');
        } catch (syntaxError) {
          _logger.error('🚨 沙箱代码语法错误:', syntaxError.message);
          _logger.info('funcStr 前100字符:', funcStr.substring(0, 100));
          safeCb({ type: 'error', message: '沙箱代码语法错误: ' + syntaxError.message });
          return;
        }

        // 尝试多种 callCommand 调用方式来解决回调问题
        _logger.info('📞 尝试不同的 callCommand 调用方式...');

        // 方法1：标准调用（你说这个不回调）
        try {
          global.Asc.plugin.callCommand(new Function(funcStr), function(info) {
            _logger.info('🎉 方法1回调执行:', info);
            if (info && info.ok) {
              _logger.info('📊 图表检测成功，通知外部');
              if (info.action === 'exists' && info.meta) {
                safeCb({
                  type: 'chart-binding-exists',
                  meta: info.meta,
                  fingerprint: info.fingerprint
                });
              } else if (info.action === 'created' && info.meta) {
                safeCb({
                  type: 'chart-binding-created',
                  meta: info.meta,
                  fingerprint: info.fingerprint
                });
              }
            }
          });
        } catch (e1) {
          _logger.error('方法1失败:', e1);
        }

        // 方法2：异步调用
        try {
          global.Asc.plugin.callCommand(new Function(funcStr), true, function(info) {
            _logger.info('🎉 方法2异步回调执行:', info);
            if (info && info.ok) {
              if (info.action === 'exists' && info.meta) {
                safeCb({
                  type: 'chart-binding-exists',
                  meta: info.meta,
                  fingerprint: info.fingerprint
                });
              } else if (info.action === 'created' && info.meta) {
                safeCb({
                  type: 'chart-binding-created',
                  meta: info.meta,
                  fingerprint: info.fingerprint
                });
              }
            }
          });
        } catch (e2) {
          _logger.error('方法2失败:', e2);
        }

        // 方法3：直接执行函数（不通过 callCommand）
        try {
          _logger.info('📞 方法3：直接执行沙箱函数...');
          var directResult = (new Function(funcStr))();
          _logger.info('📊 直接执行结果:', directResult);
          if (directResult && directResult.ok) {
            if (directResult.action === 'exists' && directResult.meta) {
              safeCb({
                type: 'chart-binding-exists',
                meta: directResult.meta,
                fingerprint: directResult.fingerprint
              });
            } else if (directResult.action === 'created' && directResult.meta) {
              safeCb({
                type: 'chart-binding-created',
                meta: directResult.meta,
                fingerprint: directResult.fingerprint
              });
            }
          }
        } catch (e3) {
          _logger.error('方法3失败:', e3);
        }

      } catch (callError) {
        _logger.error('🚨 执行失败:', callError);
        safeCb({ type: 'error', message: '执行失败: ' + callError.message });
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
        } else if (payload.type === 'chart-binding-created') {
          cb?.({
            type: 'chart-anchor-created',
            meta: payload.meta,
            fingerprint: payload.fingerprint
          });
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
        } else if (payload.type === 'no-user-chart') {
          cb?.({
            type: 'no-user-chart',
            message: payload.message
          });
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
