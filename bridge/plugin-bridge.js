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

      if (typeof _onEvent === 'function') {
        _onEvent(payload);
      } else {
        _logger.warn('📨 _onEvent 不是函数，无法发送事件');
      }
    } catch (e) {
      _logger.error('📨 safeCb 调用出错:', e);
      try { _logger.error('[cb error]', e && e.message ? e.message : e); } catch (_) { }
    }
  }

  // ===========================
  // 安装：设置 logger / 回调
  // ===========================
  function install(opts) {
    opts = opts || {};
    _logger.info('📦 ChartBinding.install 被调用，opts:', opts);

    if (opts.logger && typeof opts.logger.info === 'function') {
      _logger = opts.logger;
    }
    if (typeof opts.onEvent === 'function') {
      _onEvent = opts.onEvent;
      _logger.info('📦 onEvent 回调设置完成，类型:', typeof _onEvent);
    } else {
      _logger.warn('📦 onEvent 回调未设置或不是函数，类型:', typeof opts.onEvent);
    }

    // 注册点击事件（OnlyOffice 编辑器内部 onClick）
    global.Asc = global.Asc || {};
    global.Asc.plugin = global.Asc.plugin || {};
    global.Asc.plugin.event_onClick = handleOnClick;
    _logger.info('📦 onClick 事件处理器注册完成');
  }

  // ===========================
  // 点击事件主处理
  // ===========================
  function handleOnClick(isSelectionUse) {
    try { _logger.info('[onClick] triggered', { isSelectionUse: !!isSelectionUse }); } catch (_) { }

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
        } catch (_) { }

        var funcStr = (function () {/*
        // =============== 沙箱开始 ===============

        var out = { ok: true, action: 'none', message: '', meta: null, fingerprint: null, logs: [] };
        function dbg(){ try { out.logs.push(Array.prototype.join.call(arguments, ' ')); } catch (_e){} }

        function getDoc() { try { return Api.GetDocument(); } catch(e){ console.log('getDoc失败:', e); return null; } }

        function getChartType(sel) {
          try {

            var detectedType = null;
            var ch = null;

            // 步骤1: 首先尝试通过API获取图表类型
            try {
              if (sel && typeof sel.GetChart === 'function') {
                ch = sel.GetChart();

                if (ch && typeof ch.GetChartType === 'function') {
                  detectedType = ch.GetChartType();
                }
              } else {
                
              }

              // 也尝试直接从选择对象获取
              if (!detectedType && sel && typeof sel.GetChartType === 'function') {
                detectedType = sel.GetChartType();
              }

              // 如果上面的方法都失败，尝试其他获取图表对象的方法
              if (!ch) {

                // 方法1: 检查sel本身是否就是图表对象
                if (sel && typeof sel.GetChartType === 'function') {
                  ch = sel;
                }

                // 方法2: 尝试通过GetDrawingObjectsController获取
                if (!ch && sel && typeof sel.GetDrawingObjectsController === 'function') {
                  try {
                    var drawingController = sel.GetDrawingObjectsController();
                    if (drawingController && typeof drawingController.GetChart === 'function') {
                      ch = drawingController.GetChart();
                    }
                  } catch (e) {
                    console.log('❌ GetDrawingObjectsController方法失败:', e);
                  }
                }

                // 方法3: 检查是否有Chart属性
                if (!ch && sel && sel.Chart) {
                  ch = sel.Chart;
                }

                // 方法4: 检查是否有chartSpace属性
                if (!ch && sel && sel.chartSpace) {
                  ch = sel.chartSpace;
                }
              }

            } catch (e) {
              console.log('❌ API获取图表类型失败:', e);
            }

            // 步骤2: 只有在API返回"unknown"或获取失败时，才执行OOXML分析
            if (!detectedType || detectedType === 'unknown' || detectedType === null || detectedType === undefined) {
              // 特殊处理：如果检测到"unknown"，很可能是雷达图
              if (detectedType === 'unknown') {
                return 'radar-chart';
              }

              // 如果detectedType完全为空，尝试XML分析（简化版）
              try {
                if (ch && typeof ch.ToXML === 'function') {
                  try {
                    var xmlData = ch.ToXML();
                    if (xmlData && typeof xmlData === 'string' && xmlData.length > 0) {
                      var chartType = analyzeChartTypeFromXML(xmlData);
                      if (chartType && chartType !== 'unknown') {
                        return chartType;
                      }
                    }
                  } catch (xmlError) {
                    console.log('❌ XML分析失败:', xmlError.message);
                  }
                }

                // 备选方案：从选择对象获取XML
                if (sel && typeof sel.ToXML === 'function') {
                  try {
                    var selXml = sel.ToXML();
                    if (selXml && typeof selXml === 'string' && selXml.length > 0) {
                      var selChartType = analyzeChartTypeFromXML(selXml);
                      if (selChartType && selChartType !== 'unknown') {
                        return selChartType;
                      }
                    }
                  } catch (selXmlError) {
                    console.log('❌ 选择对象XML分析失败:', selXmlError.message);
                  }
                }
              } catch (e) {
                console.log('❌ OOXML分析失败:', e.message);
              }
            } else {
            }

            // 返回检测到的类型或默认值
            return detectedType || 'chart-generic';

          } catch (e) {
            console.log('getChartType失败:', e);
            return 'error';
          }
        }

        // OOXML图表类型分析函数
        function analyzeChartTypeFromXML(xmlString) {
          try {

            // 转换为小写便于匹配
            var xml = xmlString.toLowerCase();

            // 雷达图特征检测
            if (xml.indexOf('c:radarchart') !== -1 ||
                xml.indexOf('radarChart') !== -1 ||
                xml.indexOf('radar') !== -1) {
              return 'radar-chart';
            }

            // 饼图检测
            if (xml.indexOf('c:piechart') !== -1 ||
                xml.indexOf('pieChart') !== -1) {
              return 'pie-chart';
            }

            // 条形图检测
            if (xml.indexOf('c:barchart') !== -1 ||
                xml.indexOf('barChart') !== -1) {
              return 'bar-chart';
            }

            // 折线图检测
            if (xml.indexOf('c:linechart') !== -1 ||
                xml.indexOf('lineChart') !== -1) {
              return 'line-chart';
            }

            // 面积图检测
            if (xml.indexOf('c:areachart') !== -1 ||
                xml.indexOf('areaChart') !== -1) {
              return 'area-chart';
            }

            // 散点图检测
            if (xml.indexOf('c:scatterchart') !== -1 ||
                xml.indexOf('scatterChart') !== -1) {
              return 'scatter-chart';
            }

            // 股票图检测
            if (xml.indexOf('c:stockchart') !== -1 ||
                xml.indexOf('stockChart') !== -1) {
              return 'stock-chart';
            }

            // 气泡图检测
            if (xml.indexOf('c:bubblechart') !== -1 ||
                xml.indexOf('bubbleChart') !== -1) {
              return 'bubble-chart';
            }

            // 环形图检测
            if (xml.indexOf('c:doughnutchart') !== -1 ||
                xml.indexOf('doughnutChart') !== -1) {
              return 'doughnut-chart';
            }

            // 瀑布图检测
            if (xml.indexOf('c:waterfallchart') !== -1 ||
                xml.indexOf('waterfallChart') !== -1) {
              return 'waterfall-chart';
            }

            // 如果包含图表相关的XML但没有匹配到具体类型
            if (xml.indexOf('chart') !== -1 || xml.indexOf('c:') !== -1) {
              return 'chart-from-xml';
            }

            return 'unknown';

          } catch (e) {
            console.log('❌ XML分析出错:', e);
            return 'xml-error';
          }
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

        // 获取共享的图表指纹生成函数
        var buildChartFingerprint = null;
        try {
          if (window.ChartBinding && typeof window.ChartBinding.createBuildChartFingerprintFunction === 'function') {
            buildChartFingerprint = window.ChartBinding.createBuildChartFingerprintFunction();
            console.log('✅ 成功获取共享指纹生成函数');
          } else {
            console.log('❌ ChartBinding.createBuildChartFingerprintFunction 不可用');
          }
        } catch (e) {
          console.log('❌ 获取共享指纹生成函数失败:', e.message);
        }

        // 如果共享函数不可用，直接在沙箱内定义
        if (!buildChartFingerprint) {
          console.log('🔧 使用沙箱内置指纹生成函数');
          buildChartFingerprint = function(sel, doc) {
            var parts = [];
            console.log('🔍 开始生成指纹，图表对象:', sel);

            // 1) 精确位置索引（必须找到，否则报错）
            var exactIndex = -1;
            var all = null;

            try {
              all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
              console.log('📊 文档中图表总数:', all ? all.length : 0);

              if (!all || all.length === 0) {
                throw new Error('无法获取文档中的绘图对象列表');
              }

              // 先尝试严格匹配
              for (var i = 0; i < all.length; i++) {
                if (all[i] === sel) {
                  exactIndex = i;
                  console.log('✅ 找到精确索引(严格匹配):', exactIndex);
                  break;
                }
              }

              // 如果严格匹配失败，尝试ID匹配
              if (exactIndex < 0) {
                console.log('⚠️ 严格匹配失败，尝试ID匹配...');
                var selId = null;
                try {
                  if (sel && typeof sel.GetId === 'function') {
                    selId = sel.GetId();
                    console.log('🔍 当前图表ID:', selId);
                  }
                } catch(e) {}

                if (selId) {
                  for (var j = 0; j < all.length; j++) {
                    try {
                      if (all[j] && typeof all[j].GetId === 'function') {
                        var allId = all[j].GetId();
                        if (allId === selId) {
                          exactIndex = j;
                          console.log('✅ 找到精确索引(ID匹配):', exactIndex);
                          break;
                        }
                      }
                    } catch(e) {}
                  }
                }
              }

              // 如果ID匹配也失败，尝试属性匹配
              if (exactIndex < 0) {
                console.log('⚠️ ID匹配失败，尝试属性匹配...');
                for (var k = 0; k < all.length; k++) {
                  try {
                    // 比较多个属性来判断是否是同一个对象
                    var match = true;

                    // 比较类型
                    if (sel && typeof sel.GetClassType === 'function' &&
                        all[k] && typeof all[k].GetClassType === 'function') {
                      if (sel.GetClassType() !== all[k].GetClassType()) {
                        match = false;
                      }
                    }

                    // 比较尺寸
                    if (match && sel && typeof sel.GetWidth === 'function' &&
                        all[k] && typeof all[k].GetWidth === 'function') {
                      var selW = sel.GetWidth();
                      var allW = all[k].GetWidth();
                      if (Math.abs(selW - allW) > 1) { // 允许1像素误差
                        match = false;
                      }
                    }

                    if (match && sel && typeof sel.GetHeight === 'function' &&
                        all[k] && typeof all[k].GetHeight === 'function') {
                      var selH = sel.GetHeight();
                      var allH = all[k].GetHeight();
                      if (Math.abs(selH - allH) > 1) { // 允许1像素误差
                        match = false;
                      }
                    }

                    if (match) {
                      exactIndex = k;
                      console.log('✅ 找到精确索引(属性匹配):', exactIndex);
                      break;
                    }
                  } catch(e) {}
                }
              }

              if (exactIndex < 0) {
                throw new Error('无法在文档绘图对象列表中找到当前图表的索引位置');
              }

            } catch(e) {
              console.error('❌ 获取图表索引失败:', e.message);
              throw new Error('图表索引获取失败: ' + e.message);
            }

            // 必须的索引部分
            parts.push('idx:' + exactIndex);
            console.log('✅ 添加索引部分:', 'idx:' + exactIndex);

            // 2) 内部ID（可选，增强唯一性）
            try {
              if (sel && typeof sel.GetId === 'function') {
                var stableId = sel.GetId();
                if (stableId) {
                  parts.push('id:' + stableId);
                  console.log('✅ 添加内部ID:', stableId);
                }
              }
            } catch(e){
              console.log('⚠️ 获取内部ID失败，继续执行:', e.message);
            }

            // 3) 图表类型（可选）
            try {
              if (sel && typeof sel.GetChart === 'function') {
                var ch = sel.GetChart();
                if (ch && typeof ch.GetChartType === 'function') {
                  var type = ch.GetChartType();
                  if (type && type !== 'error') {
                    parts.push('type:' + type);
                    console.log('✅ 添加图表类型:', type);
                  }
                }
              } else if (sel && typeof sel.GetChartType === 'function') {
                var type = sel.GetChartType();
                if (type && type !== 'error') {
                  parts.push('type:' + type);
                  console.log('✅ 添加图表类型:', type);
                }
              }
            } catch(e) {
              console.log('⚠️ 获取图表类型失败，继续执行:', e.message);
            }

            // 4) 尺寸信息（可选）
            try {
              if (sel && typeof sel.GetWidth === 'function' && typeof sel.GetHeight === 'function') {
                var w = sel.GetWidth();
                var h = sel.GetHeight();
                if (w && h && w > 0 && h > 0) {
                  var sizePart = 'size:' + Math.floor(w) + 'x' + Math.floor(h);
                  parts.push(sizePart);
                  console.log('✅ 添加尺寸信息:', sizePart);
                }
              }
            } catch(e){
              console.log('⚠️ 获取尺寸失败，继续执行:', e.message);
            }

            // 必须至少有索引，如果没有就报错
            if (parts.length === 0) {
              throw new Error('无法生成任何有效的指纹部分');
            }

            var fingerprint = parts.join('|');
            console.log('🔖 最终生成指纹:', fingerprint);
            return fingerprint;
          };
        }

        // 文档自定义属性读/写
        function getBindingByFingerprint(doc, fp) {
          try {
            var props = doc.GetCustomProperties();
            var key = 'chart-binding:' + fp;

            var val = props.Get(key);
            if (!val) return null;
            try {
              var result = JSON.parse(val);
              return result;
            } catch(_){
              return null;
            }
          } catch(e){
            console.log('❌ 读取绑定失败:', e);
            return null;
          }
        }
        function setBindingByFingerprint(doc, fp, meta) {
          try {
            var props = doc.GetCustomProperties();
            var key = 'chart-binding:' + fp;
            props.Add(key, JSON.stringify(meta)); // 存在同名时覆盖
          } catch(e){
            console.log('❌ 写入绑定失败:', e);
          }
        }

        var doc = getDoc();
        if (!doc) {
          out.ok = false;
          out.action = 'error';
          out.message = '无法获取文档对象';
          return out;
        }

        var selected = null;
        try {
          selected = doc.GetSelectedDrawings ? doc.GetSelectedDrawings() : null;
        } catch(e){
          console.log('GetSelectedDrawings失败:', e);
        }
        dbg('🎯 选中的绘图对象数量:', selected ? selected.length : 0);

        if (!selected || !selected.length) {
          // 没选中图表 → 再看文本选区（避免误触）
          var range = null;
          try { range = doc.GetRangeBySelect && doc.GetRangeBySelect(); } catch(_){}
          if (range) {
            out.action = 'text-click';
            out.message = '点击文本区域，跳过图表绑定';
            return out;
          }
          out.action = 'no-user-chart';
          out.message = '未检测到用户选中的图表';
          return out;
        }

        var sel = selected[selected.length - 1];

        // 严格模式：必须能生成指纹，否则报错
        var fp = null;
        try {
          fp = buildChartFingerprint(sel, doc);
          console.log('✅ 成功生成指纹:', fp);
        } catch (fpError) {
          console.error('❌ 指纹生成失败:', fpError.message);
          out.ok = false;
          out.action = 'error';
          out.message = '指纹生成失败: ' + fpError.message;
          return out;
        }

        out.fingerprint = fp;
        dbg('🔖 指纹:', fp);

        var existed = getBindingByFingerprint(doc, fp);
        if (existed) {
          out.action = 'exists';
          out.meta = existed;
          out.message = '已存在图表绑定（自定义属性）';

          return out;
        }

        var meta = {
          chartId: 'chart-' + fp.replace(/[^a-zA-Z0-9]/g, '_') + '_' + Math.abs(hashCode(fp)).toString(36),
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

        // 简单哈希函数，基于指纹生成稳定的ID
        function hashCode(str) {
          var hash = 0;
          if (str.length === 0) return hash;
          for (var i = 0; i < str.length; i++) {
            var char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // 转换为32位整数
          }
          return hash;
        }
        setBindingByFingerprint(doc, fp, meta);

        // 由于 callCommand 回调不可靠，我们在沙箱内部直接触发事件
        out.action = 'created';
        out.meta = meta;
        out.message = '已写入绑定到文档自定义属性';

        // 沙箱通信方案：在文档自定义属性中写入临时结果（使用稳定的key）
        try {
          var tempKey = 'temp-chart-result-' + Math.abs(hashCode(fp || 'default')).toString(36);
          var props = doc.GetCustomProperties();
          props.Add(tempKey, JSON.stringify({
            timestamp: new Date().toISOString(),
            result: out
          }));
          console.log('📝 沙箱结果已写入临时属性:', tempKey);
        } catch (writeErr) {
          console.log('写入临时结果失败:', writeErr);
        }

        return out;
        // =============== 沙箱结束 ===============
      */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

        try {

          window.Asc.scope.ccPr = ccPr;
          const P = window.Asc.plugin;

          try {
            var testFunc = new Function(funcStr);
            _logger.info('✅ 沙箱代码语法检查通过');
          } catch (syntaxError) {
            _logger.error('🚨 沙箱代码语法错误:', syntaxError.message);
            safeCb({ type: 'error', message: '沙箱代码语法错误: ' + syntaxError.message });
            return;
          }
          
          //把原始沙箱也改成延迟执行
          function callInNextTick(fn, cb) {
            setTimeout(() => P.callCommand(fn, false, false, cb), 0);
          }

          callInNextTick(function () { return 2; }, function (ret) {
            console.log('✅ 测试回调成功:', ret);

            callInNextTick(new Function(funcStr), function (info) {

              // 处理沙箱返回的结果
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
                } else if (info.action === 'text-click') {
                  _logger.info('📊 检测到文本点击，不发送图表事件');
                } else if (info.action === 'no-user-chart') {
                  _logger.info('📊 检测到无图表选择');
                } else {
                  _logger.info('📊 其他沙箱结果 action:', info.action, 'meta存在:', !!info.meta);
                }
              } else {
                _logger.warn('📊 沙箱执行失败或ok为假:', info);
              }
            });
          });

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
    var funcStr = (function () {/*
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
        cb && cb({ ok: false, message: info && info.message });
        return;
      }
      cb && cb({ ok: true, bindings: info.bindings || [] });
    });
  }

  // ===========================
  // 工具：为"当前选中图表"写入/更新绑定（手动 Upsert）
  // payload: 你的业务数据（对象），将合并写入 meta
  // ===========================
  function upsertBindingForSelectedChart(payload, cb) {
    payload = payload || {};
    var funcStr = (function () {/*
      var out = { ok:true, action:'none', meta:null, fingerprint:null, message:'', logs:[] };
      function dbg(){ try { out.logs.push(Array.prototype.join.call(arguments, ' ')); } catch (_e){} }

      function getDoc(){ try { return Api.GetDocument(); } catch(e){ return null; } }
      function getChartType(sel){
        try{
          console.log('🔍 [upsert] 开始检测图表类型，对象类型:', sel ? sel.GetClassType() : 'null');

          var detectedType = null;
          var ch = null;

          // 步骤1: 首先尝试通过API获取图表类型
          try {
            if (sel && typeof sel.GetChart === 'function') {
              ch = sel.GetChart();

              if (ch && typeof ch.GetChartType === 'function') {
                detectedType = ch.GetChartType();
                console.log('📊 [upsert] API返回的图表类型:', detectedType);
              }
            }

            // 也尝试直接从选择对象获取
            if (!detectedType && sel && typeof sel.GetChartType === 'function') {
              detectedType = sel.GetChartType();
              console.log('📊 [upsert] 直接从选择对象获取图表类型:', detectedType);
            }
          } catch (e) {
            console.log('❌ [upsert] API获取图表类型失败:', e);
          }

          // 步骤2: 只有在API返回"unknown"或获取失败时，才执行OOXML分析
          if (!detectedType || detectedType === 'unknown' || detectedType === null || detectedType === undefined) {
            console.log('🔍 [upsert] API无法获取准确类型，开始OOXML分析...');

            
            if (detectedType === 'unknown') {
              console.log('[upsert] 检测到unknown类型，根据测试经验，很可能是雷达图');
              return 'radar-chart';
            }

            // 如果detectedType完全为空，尝试XML分析
            try {
              if (ch && typeof ch.ToXML === 'function') {
                try {
                  var xmlData = ch.ToXML();
                  if (xmlData && typeof xmlData === 'string' && xmlData.length > 0) {
                    var chartType = analyzeChartTypeFromXML_upsert(xmlData);
                    if (chartType && chartType !== 'unknown') {
                      console.log('✅ [upsert] 从XML分析得到图表类型:', chartType);
                      return chartType;
                    }
                  }
                } catch (xmlError) {
                  console.log('❌ [upsert] XML分析失败:', xmlError.message);
                }
              }
            } catch (e) {
              console.log('❌ [upsert] OOXML分析失败:', e.message);
            }
          } else {
            console.log('✅ [upsert] API成功获取到图表类型，跳过XML分析');
          }

          // 返回检测到的类型或默认值
          return detectedType || 'chart-generic';

        }catch(e){
          console.log('[upsert] getChartType失败:', e);
          return 'error';
        }
      }

      // upsert版本的XML分析函数
      function analyzeChartTypeFromXML_upsert(xmlString) {
        try {
          console.log('🔍 [upsert] 开始分析XML数据...');
          var xml = xmlString.toLowerCase();

          // 雷达图特征检测
          if (xml.indexOf('c:radarchart') !== -1 ||
              xml.indexOf('radarChart') !== -1 ||
              xml.indexOf('radar') !== -1) {
            console.log('🎯 [upsert] 检测到雷达图特征');
            return 'radar-chart';
          }

          // 饼图检测
          if (xml.indexOf('c:piechart') !== -1 ||
              xml.indexOf('pieChart') !== -1) {
            console.log('🥧 [upsert] 检测到饼图特征');
            return 'pie-chart';
          }

          // 条形图检测
          if (xml.indexOf('c:barchart') !== -1 ||
              xml.indexOf('barChart') !== -1) {
            console.log('📊 [upsert] 检测到条形图特征');
            return 'bar-chart';
          }

          // 折线图检测
          if (xml.indexOf('c:linechart') !== -1 ||
              xml.indexOf('lineChart') !== -1) {
            console.log('📈 [upsert] 检测到折线图特征');
            return 'line-chart';
          }

          // 面积图检测
          if (xml.indexOf('c:areachart') !== -1 ||
              xml.indexOf('areaChart') !== -1) {
            console.log('📊 [upsert] 检测到面积图特征');
            return 'area-chart';
          }

          // 散点图检测
          if (xml.indexOf('c:scatterchart') !== -1 ||
              xml.indexOf('scatterChart') !== -1) {
            console.log('⚡ [upsert] 检测到散点图特征');
            return 'scatter-chart';
          }

          // 其他图表类型...
          if (xml.indexOf('chart') !== -1 || xml.indexOf('c:') !== -1) {
            console.log('📊 [upsert] 检测到通用图表XML特征');
            return 'chart-from-xml';
          }

          console.log('❌ [upsert] XML中未找到已知的图表类型特征');
          return 'unknown';

        } catch (e) {
          console.log('❌ [upsert] XML分析出错:', e);
          return 'xml-error';
        }
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
        // 获取共享的图表指纹生成函数
        var sharedFunc = (window.ChartBinding && window.ChartBinding.createBuildChartFingerprintFunction)
          ? window.ChartBinding.createBuildChartFingerprintFunction()
          : null;

        if (sharedFunc) {
          return sharedFunc(sel, doc);
        }

        // 降级方案：使用与主版本完全一致的逻辑
        var parts = [];
        console.log('🔍 开始生成指纹，图表对象:', sel);

        // 1) 精确位置索引（必须找到，否则报错）
        var exactIndex = -1;
        var all = null;

        try {
          all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
          console.log('📊 文档中图表总数:', all ? all.length : 0);

          if (!all || all.length === 0) {
            throw new Error('无法获取文档中的绘图对象列表');
          }

          // 先尝试严格匹配
          for (var i = 0; i < all.length; i++) {
            if (all[i] === sel) {
              exactIndex = i;
              console.log('✅ 找到精确索引(严格匹配):', exactIndex);
              break;
            }
          }

          // 如果严格匹配失败，尝试ID匹配
          if (exactIndex < 0) {
            console.log('⚠️ 严格匹配失败，尝试ID匹配...');
            var selId = null;
            try {
              if (sel && typeof sel.GetId === 'function') {
                selId = sel.GetId();
                console.log('🔍 当前图表ID:', selId);
              }
            } catch(e) {}

            if (selId) {
              for (var j = 0; j < all.length; j++) {
                try {
                  if (all[j] && typeof all[j].GetId === 'function') {
                    var allId = all[j].GetId();
                    if (allId === selId) {
                      exactIndex = j;
                      console.log('✅ 找到精确索引(ID匹配):', exactIndex);
                      break;
                    }
                  }
                } catch(e) {}
              }
            }
          }

          // 如果ID匹配也失败，尝试属性匹配
          if (exactIndex < 0) {
            console.log('⚠️ ID匹配失败，尝试属性匹配...');
            for (var k = 0; k < all.length; k++) {
              try {
                // 比较多个属性来判断是否是同一个对象
                var match = true;

                // 比较类型
                if (sel && typeof sel.GetClassType === 'function' &&
                    all[k] && typeof all[k].GetClassType === 'function') {
                  if (sel.GetClassType() !== all[k].GetClassType()) {
                    match = false;
                  }
                }

                // 比较尺寸
                if (match && sel && typeof sel.GetWidth === 'function' &&
                    all[k] && typeof all[k].GetWidth === 'function') {
                  var selW = sel.GetWidth();
                  var allW = all[k].GetWidth();
                  if (Math.abs(selW - allW) > 1) { // 允许1像素误差
                    match = false;
                  }
                }

                if (match && sel && typeof sel.GetHeight === 'function' &&
                    all[k] && typeof all[k].GetHeight === 'function') {
                  var selH = sel.GetHeight();
                  var allH = all[k].GetHeight();
                  if (Math.abs(selH - allH) > 1) { // 允许1像素误差
                    match = false;
                  }
                }

                if (match) {
                  exactIndex = k;
                  console.log('✅ 找到精确索引(属性匹配):', exactIndex);
                  break;
                }
              } catch(e) {}
            }
          }

          if (exactIndex < 0) {
            throw new Error('无法在文档绘图对象列表中找到当前图表的索引位置');
          }

        } catch(e) {
          console.error('❌ 获取图表索引失败:', e.message);
          throw new Error('图表索引获取失败: ' + e.message);
        }

        // 必须的索引部分
        parts.push('idx:' + exactIndex);
        console.log('✅ 添加索引部分:', 'idx:' + exactIndex);

        // 2) 内部ID（可选，增强唯一性）
        try {
          if (sel && typeof sel.GetId === 'function') {
            var stableId = sel.GetId();
            if (stableId) {
              parts.push('id:' + stableId);
              console.log('✅ 添加内部ID:', stableId);
            }
          }
        } catch(e){
          console.log('⚠️ 获取内部ID失败，继续执行:', e.message);
        }

        // 3) 图表类型（可选）
        try {
          var type = getChartType(sel);
          if (type && type !== 'error') {
            parts.push('type:' + type);
            console.log('✅ 添加图表类型:', type);
          }
        } catch(e) {
          console.log('⚠️ 获取图表类型失败，继续执行:', e.message);
        }

        // 4) 尺寸信息（可选）
        try {
          if (sel && typeof sel.GetWidth === 'function' && typeof sel.GetHeight === 'function') {
            var w = sel.GetWidth();
            var h = sel.GetHeight();
            if (w && h && w > 0 && h > 0) {
              var sizePart = 'size:' + Math.floor(w) + 'x' + Math.floor(h);
              parts.push(sizePart);
              console.log('✅ 添加尺寸信息:', sizePart);
            }
          }
        } catch(e){
          console.log('⚠️ 获取尺寸失败，继续执行:', e.message);
        }

        // 必须至少有索引，如果没有就报错
        if (parts.length === 0) {
          throw new Error('无法生成任何有效的指纹部分');
        }

        var fingerprint = parts.join('|');
        console.log('🔖 最终生成指纹:', fingerprint);
        return fingerprint;
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
        // 简单哈希函数，基于指纹生成稳定的ID
        function hashCode(str) {
          var hash = 0;
          if (str.length === 0) return hash;
          for (var i = 0; i < str.length; i++) {
            var char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // 转换为32位整数
          }
          return hash;
        }

        meta={
          chartId: 'chart-' + fp.replace(/[^a-zA-Z0-9]/g, '_') + '_' + Math.abs(hashCode(fp)).toString(36),
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

    try { global.Asc.scope = global.Asc.scope || {}; global.Asc.scope._cb_payload = payload; } catch (_) { }

    global.Asc.plugin.callCommand(new Function(funcStr), function (info) {
      if (!info || info.ok === false) { cb && cb({ ok: false, message: info && info.message }); return; }
      cb && cb({ ok: true, action: info.action, meta: info.meta, fingerprint: info.fingerprint });
    });
  }

  // ===========================
  // 工具：按 fingerprint "软删除"绑定（设为空串）
  // （如你确认 API 支持真正 Remove，可改为 Remove(key)）
  // ===========================
  function removeBindingByFingerprint(fingerprint, cb) {
    if (!fingerprint) { cb && cb({ ok: false, message: 'fingerprint 为空' }); return; }

    var funcStr = (function () {/*
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

    try { global.Asc.scope = global.Asc.scope || {}; global.Asc.scope._fp = fingerprint; } catch (_) { }

    global.Asc.plugin.callCommand(new Function(funcStr), function (info) {
      if (!info || info.ok === false) { cb && cb({ ok: false, message: info && info.message }); return; }
      cb && cb({ ok: true, message: info.message });
    });
  }

  // 共享的图表指纹生成函数（严格模式 - 无备用方案）
  function createBuildChartFingerprintFunction() {
    return function buildChartFingerprint(sel, doc) {
      var parts = [];
      console.log('🔍 开始生成指纹，图表对象:', sel);

      // 1) 精确位置索引（必须找到，否则报错）
      var exactIndex = -1;
      var all = null;

      try {
        all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
        console.log('📊 文档中图表总数:', all ? all.length : 0);

        if (!all || all.length === 0) {
          throw new Error('无法获取文档中的绘图对象列表');
        }

        // 先尝试严格匹配
        for (var i = 0; i < all.length; i++) {
          if (all[i] === sel) {
            exactIndex = i;
            console.log('✅ 找到精确索引(严格匹配):', exactIndex);
            break;
          }
        }

        // 如果严格匹配失败，尝试ID匹配
        if (exactIndex < 0) {
          console.log('⚠️ 严格匹配失败，尝试ID匹配...');
          var selId = null;
          try {
            if (sel && typeof sel.GetId === 'function') {
              selId = sel.GetId();
              console.log('🔍 当前图表ID:', selId);
            }
          } catch(e) {}

          if (selId) {
            for (var j = 0; j < all.length; j++) {
              try {
                if (all[j] && typeof all[j].GetId === 'function') {
                  var allId = all[j].GetId();
                  if (allId === selId) {
                    exactIndex = j;
                    console.log('✅ 找到精确索引(ID匹配):', exactIndex);
                    break;
                  }
                }
              } catch(e) {}
            }
          }
        }

        // 如果ID匹配也失败，尝试属性匹配
        if (exactIndex < 0) {
          console.log('⚠️ ID匹配失败，尝试属性匹配...');
          for (var k = 0; k < all.length; k++) {
            try {
              // 比较多个属性来判断是否是同一个对象
              var match = true;

              // 比较类型
              if (sel && typeof sel.GetClassType === 'function' &&
                  all[k] && typeof all[k].GetClassType === 'function') {
                if (sel.GetClassType() !== all[k].GetClassType()) {
                  match = false;
                }
              }

              // 比较尺寸
              if (match && sel && typeof sel.GetWidth === 'function' &&
                  all[k] && typeof all[k].GetWidth === 'function') {
                var selW = sel.GetWidth();
                var allW = all[k].GetWidth();
                if (Math.abs(selW - allW) > 1) { // 允许1像素误差
                  match = false;
                }
              }

              if (match && sel && typeof sel.GetHeight === 'function' &&
                  all[k] && typeof all[k].GetHeight === 'function') {
                var selH = sel.GetHeight();
                var allH = all[k].GetHeight();
                if (Math.abs(selH - allH) > 1) { // 允许1像素误差
                  match = false;
                }
              }

              if (match) {
                exactIndex = k;
                console.log('✅ 找到精确索引(属性匹配):', exactIndex);
                break;
              }
            } catch(e) {}
          }
        }

        if (exactIndex < 0) {
          throw new Error('无法在文档绘图对象列表中找到当前图表的索引位置');
        }

      } catch(e) {
        console.error('❌ 获取图表索引失败:', e.message);
        throw new Error('图表索引获取失败: ' + e.message);
      }

      // 必须的索引部分
      parts.push('idx:' + exactIndex);
      console.log('✅ 添加索引部分:', 'idx:' + exactIndex);

      // 2) 内部ID（可选，增强唯一性）
      try {
        if (sel && typeof sel.GetId === 'function') {
          var stableId = sel.GetId();
          if (stableId) {
            parts.push('id:' + stableId);
            console.log('✅ 添加内部ID:', stableId);
          }
        }
      } catch(e){
        console.log('⚠️ 获取内部ID失败，继续执行:', e.message);
      }

      // 3) 图表类型（可选）
      try {
        var type = getChartType(sel);
        if (type && type !== 'error') {
          parts.push('type:' + type);
          console.log('✅ 添加图表类型:', type);
        }
      } catch(e) {
        console.log('⚠️ 获取图表类型失败，继续执行:', e.message);
      }

      // 4) 尺寸信息（可选）
      try {
        if (sel && typeof sel.GetWidth === 'function' && typeof sel.GetHeight === 'function') {
          var w = sel.GetWidth();
          var h = sel.GetHeight();
          if (w && h && w > 0 && h > 0) {
            var sizePart = 'size:' + Math.floor(w) + 'x' + Math.floor(h);
            parts.push(sizePart);
            console.log('✅ 添加尺寸信息:', sizePart);
          }
        }
      } catch(e){
        console.log('⚠️ 获取尺寸失败，继续执行:', e.message);
      }

      // 必须至少有索引，如果没有就报错
      if (parts.length === 0) {
        throw new Error('无法生成任何有效的指纹部分');
      }

      var fingerprint = parts.join('|');
      console.log('🔖 最终生成指纹:', fingerprint);
      return fingerprint;
    };

    // 简单哈希函数（仅用于辅助功能）
    function simpleHash(str) {
      var hash = 0;
      if (str.length === 0) return hash.toString(36);
      for (var i = 0; i < str.length; i++) {
        var char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // 转换为32位整数
      }
      return Math.abs(hash).toString(36);
    }

    // 获取图表类型
    function getChartType(sel) {
      try {
        var detectedType = null;
        var ch = null;

        // 步骤1: 首先尝试通过API获取图表类型
        try {
          if (sel && typeof sel.GetChart === 'function') {
            ch = sel.GetChart();
            if (ch && typeof ch.GetChartType === 'function') {
              detectedType = ch.GetChartType();
            }
          }

          // 也尝试直接从选择对象获取
          if (!detectedType && sel && typeof sel.GetChartType === 'function') {
            detectedType = sel.GetChartType();
          }
        } catch (e) {
          console.log('❌ API获取图表类型失败:', e);
        }

        // 步骤2: 只有在API返回"unknown"或获取失败时，才执行OOXML分析
        if (!detectedType || detectedType === 'unknown' || detectedType === null || detectedType === undefined) {
          // 特殊处理：如果检测到"unknown"，很可能是雷达图
          if (detectedType === 'unknown') {
            return 'radar-chart';
          }
        }

        // 返回检测到的类型或默认值
        return detectedType || 'chart-generic';

      } catch (e) {
        console.log('getChartType失败:', e);
        return 'error';
      }
    }

    // 查找宿主段落
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

    // 收集祖先链
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
  }

  // 导出
  var API = {
    install: install,
    listBindings: listBindings,
    upsertBindingForSelectedChart: upsertBindingForSelectedChart,
    removeBindingByFingerprint: removeBindingByFingerprint,
    createBuildChartFingerprintFunction: createBuildChartFingerprintFunction
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
    if (callback) {
      // 有回调函数时，使用与plugin-bridge.js一致的callInNextTick模式
      const P = window.Asc?.plugin;
      if (!P || !P.callCommand) {
        console.error('❌ OnlyOffice plugin API不可用');
        callback && callback({ success: false, error: 'OnlyOffice plugin API不可用' });
        return;
      }

      setTimeout(() => {
        P.callCommand(fn, false, false, callback);
      }, 0);
    } else {
      // 无回调函数时，直接调用
      return window.Asc?.plugin?.callCommand?.(fn, isAsync);
    }
  }

  onSelectionChanged(cb) {
    // 初始化 ChartBinding 系统，但保留原有的选择变化逻辑
    window.ChartBinding.install({
      logger: logger,
      onEvent: function (payload) {
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
