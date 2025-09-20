// services/chart-binding-service.js
import { ChartTypeDetector } from './chart-type-detector.js';
import { ChartDataBinder } from './chart-data-binder.js';
import { ChartDetector } from './chart-detector.js';

export class ChartBindingService {
    constructor(editorService) {
        this.editor = editorService;

        // 注入专门的服务类
        this.typeDetector = new ChartTypeDetector();
        this.dataBinder = new ChartDataBinder();
        this.chartDetector = new ChartDetector();
    }

    /**
     * 为文档中的图表绑定隐藏数据
     * @param {Object} chartData - 包含chartId和数据的对象，支持两种格式：
     *   1. 嵌套格式: {chartId: "xxx", data: {...}, metadata: {...}}
     *   2. 扁平格式: {type: "chart", rid: "xxx", chartType: "xxx", tag: "...", chartId: "xxx"}
     */
    async bindDataToChart(chartData) {
        console.log('开始图表数据绑定流程');
        console.log('从宿主页面接收到的图表数据:', chartData);

        // 检查是否从宿主页面传递了有效数据
        if (!chartData || !chartData.chartId) {
            console.log('❌ 未收到有效的图表ID，无法进行绑定');
            return {
                success: false,
                error: '未从宿主页面接收到有效的chartId，请确保调用时传递了chartId',
                data: {
                    timestamp: new Date().toLocaleString()
                }
            };
        }

        // 适配扁平数据结构 - 检测是否为新的扁平格式
        let normalizedData = chartData;
        if (chartData.type && chartData.rid && !chartData.data && !chartData.metadata) {
            console.log('🔄 检测到扁平数据格式，进行数据结构转换...');

            // 解析tag字段中的JSON数据
            let tagData = {};
            if (chartData.tag) {
                try {
                    tagData = JSON.parse(chartData.tag);
                    console.log('📋 解析tag数据:', tagData);
                } catch (e) {
                    console.log('⚠️ tag数据解析失败，使用原始字符串:', chartData.tag);
                    tagData = { rawTag: chartData.tag };
                }
            }

            // 转换为标准的嵌套格式
            normalizedData = {
                chartId: chartData.chartId,
                data: {
                    type: chartData.type,
                    rid: chartData.rid,
                    chartType: chartData.chartType,
                    tag: chartData.tag,
                    tagData: tagData,
                    originalFormat: 'flat'
                },
                metadata: {
                    sourceFormat: 'flat',
                    convertedAt: new Date().toISOString(),
                    trackId: tagData.trackId || null,
                    groupFields: tagData.groupFields || [],
                    timestamp: tagData._t || Date.now()
                }
            };

            console.log('✅ 数据格式转换完成:', normalizedData);
        }

        // 验证转换后的数据结构
        if (!normalizedData.data || !normalizedData.metadata) {
            console.log('❌ 数据结构验证失败，缺少必要字段');
            return {
                success: false,
                error: '数据结构不完整，请确保包含data和metadata字段，或提供有效的扁平格式数据',
                data: {
                    timestamp: new Date().toLocaleString(),
                    receivedData: chartData
                }
            };
        }

        const targetChartId = normalizedData.chartId;
        console.log('🎯 目标图表ID:', targetChartId);

        // 由于runInDoc是隔离环境，将数据直接嵌入到函数体中
        const chartDataJson = JSON.stringify(normalizedData);
        console.log('🔍 准备嵌入的数据JSON长度:', chartDataJson.length);

        // 创建包含数据的函数字符串，使用与plugin-bridge.js一致的方式
        const funcStr = (function(){/*
            var out = { ok: true, bindings: [], logs: [] };
            function dbg(){ try { out.logs.push(Array.prototype.join.call(arguments, ' ')); } catch (_e){} }

            var doc = Api.GetDocument();
            console.log('=== 图表数据绑定开始 ===');

            try {
                // 使用与plugin-bridge.js一致的方式获取数据
                var bindingData = null;
                try {
                    bindingData = Asc && Asc.scope && Asc.scope._cb_chartData ? Asc.scope._cb_chartData : null;
                    if (bindingData) {
                        console.log('📊 从scope获取绑定数据:', bindingData);
                    } else {
                        console.log('❌ 无法从scope获取绑定数据');
                        return { success: false, error: '无法获取绑定数据' };
                    }
                } catch (scopeError) {
                    console.log('❌ 获取scope数据失败:', scopeError.message);
                    return { success: false, error: '获取scope数据失败: ' + scopeError.message };
                }

                var targetChartId = bindingData.chartId;
                console.log('🎯 查找图表ID:', targetChartId);

                // 通过chartId从文档自定义属性中查找图表
                var props = doc.GetCustomProperties();
                var foundChart = null;
                var foundFingerprint = null;

                // 遍历所有图表绑定属性，查找匹配的chartId
                console.log('🔍 开始在文档自定义属性中查找图表...');

                // 1. 先扫描当前选中的图表
                var selectedDrawings = doc.GetSelectedDrawings();
                if (selectedDrawings && selectedDrawings.length > 0) {
                    console.log('📋 检查当前选中的图表...');
                    for (var i = 0; i < selectedDrawings.length; i++) {
                        var drawing = selectedDrawings[i];

                        // 生成指纹来查找绑定
                        var fingerprint = generateChartFingerprint(drawing, doc);
                        var bindingKey = 'chart-binding:' + fingerprint;

                        try {
                            var bindingValue = props.Get(bindingKey);
                            if (bindingValue) {
                                var binding = JSON.parse(bindingValue);
                                console.log('📊 检查绑定:', binding);

                                if (binding.chartId === targetChartId) {
                                    console.log('✅ 找到匹配的图表!');
                                    foundChart = drawing;
                                    foundFingerprint = fingerprint;
                                    break;
                                }
                            }
                        } catch (e) {
                            console.log('解析绑定数据失败:', e);
                        }
                    }
                }

                // 2. 如果当前选中的图表中没找到，扫描所有图表
                if (!foundChart) {
                    console.log('📋 在所有文档图表中查找...');
                    var allDrawings = doc.GetAllDrawingObjects();

                    if (allDrawings && allDrawings.length > 0) {
                        for (var j = 0; j < allDrawings.length; j++) {
                            var drawing = allDrawings[j];

                            // 生成指纹来查找绑定
                            var fingerprint = generateChartFingerprint(drawing, doc);
                            var bindingKey = 'chart-binding:' + fingerprint;

                            try {
                                var bindingValue = props.Get(bindingKey);
                                if (bindingValue) {
                                    var binding = JSON.parse(bindingValue);
                                    console.log('📊 检查绑定:', binding);

                                    if (binding.chartId === targetChartId) {
                                        console.log('✅ 找到匹配的图表!');
                                        foundChart = drawing;
                                        foundFingerprint = fingerprint;
                                        break;
                                    }
                                }
                            } catch (e) {
                                console.log('解析绑定数据失败:', e);
                            }
                        }
                    }
                }

                if (!foundChart) {
                    console.log('❌ 未找到匹配的图表');
                    return {
                        success: false,
                        error: '未找到chartId为 ' + targetChartId + ' 的图表',
                        chartId: targetChartId
                    };
                }

                console.log('🎯 成功找到目标图表，开始绑定数据...');

                // 更新图表绑定数据
                var bindingKey = 'chart-binding:' + foundFingerprint;
                var existingBindingValue = props.Get(bindingKey);
                var updatedBinding;

                if (existingBindingValue) {
                    updatedBinding = JSON.parse(existingBindingValue);
                } else {
                    updatedBinding = {
                        chartId: targetChartId,
                        createdAt: new Date().toISOString(),
                        fingerprint: foundFingerprint
                    };
                }

                // 合并新的数据
                console.log('🔄 开始更新绑定数据...');
                console.log('🆔 旧绑定数据:', updatedBinding);
                console.log('📦 新绑定数据 bindingData.data:', bindingData.data);
                console.log('📦 新绑定数据 bindingData.metadata:', bindingData.metadata);

                updatedBinding.boundData = bindingData.data;
                updatedBinding.metadata = bindingData.metadata;
                updatedBinding.lastUpdated = new Date().toISOString();

                // 保存重要的tag信息
                if (bindingData.data && bindingData.data.tag) {
                    console.log('📌 更新tag数据: 从', updatedBinding.tag, '到', bindingData.data.tag);
                    updatedBinding.tag = bindingData.data.tag;
                    updatedBinding.tagData = bindingData.data.tagData;
                }

                // 保存图表类型信息
                if (bindingData.data && bindingData.data.chartType) {
                    console.log('📊 更新图表类型: 从', updatedBinding.chartType, '到', bindingData.data.chartType);
                    updatedBinding.chartType = bindingData.data.chartType;
                }

                // 保存rid信息
                if (bindingData.data && bindingData.data.rid) {
                    console.log('🆔 更新rid: 从', updatedBinding.rid, '到', bindingData.data.rid);
                    updatedBinding.rid = bindingData.data.rid;
                }

                // 保存更新后的绑定
                try {
                    props.Add(bindingKey, JSON.stringify(updatedBinding));
                    console.log('✅ 图表数据绑定完成');
                } catch (saveError) {
                    console.log('❌ 保存绑定数据失败:', saveError);
                    throw saveError;
                }

                return {
                    success: true,
                    message: '图表数据绑定成功',
                    chartId: targetChartId,
                    boundData: bindingData.data,
                    timestamp: new Date().toISOString()
                };

            } catch (error) {
                console.log('❌ 图表数据绑定失败:', error);
                return {
                    success: false,
                    error: '图表数据绑定失败: ' + error.message,
                    timestamp: new Date().toISOString()
                };
            }

            // 辅助函数：生成图表指纹（严格模式，与plugin-bridge.js保持完全一致）
            function generateChartFingerprint(sel, doc) {
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

            // 辅助函数：获取图表类型
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

            // 辅助函数：查找宿主段落
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

            // 辅助函数：收集祖先链
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
        */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

        console.log('🔍 准备执行图表绑定沙箱代码...');

        // 将数据设置到scope中，使用与plugin-bridge.js一致的方式
        try {
            if (!window.Asc) window.Asc = {};
            if (!window.Asc.scope) window.Asc.scope = {};
            window.Asc.scope._cb_chartData = normalizedData;
            console.log('✅ 数据已设置到 Asc.scope._cb_chartData');
        } catch (scopeError) {
            console.error('❌ 设置scope数据失败:', scopeError);
            return {
                success: false,
                error: '设置scope数据失败: ' + scopeError.message,
                chartId: targetChartId
            };
        }

         try {
            var testFunc = new Function(funcStr);
            console.log('✅ 沙箱代码语法检查通过');
          } catch (syntaxError) {
            console.error('🚨 沙箱代码语法错误:', syntaxError.message);
            return {
                success: false,
                error: '沙箱代码语法错误: ' + syntaxError.message,
                chartId: targetChartId
            };
          }
          
        return new Promise((resolve) => {
            this.editor.runInDoc(new Function(funcStr), {
                async: false,
                cb: (result) => {
                    console.log('📋 图表绑定沙箱执行结果:', result);

                    if (result && result.success) {
                        console.log('✅ 图表数据绑定成功');
                        resolve({
                            success: true,
                            message: result.message,
                            chartId: result.chartId,
                            boundData: result.boundData,
                            timestamp: result.timestamp
                        });
                    } else {
                        console.log('❌ 图表数据绑定失败');
                        resolve({
                            success: false,
                            error: result ? result.error : '图表绑定执行失败',
                            chartId: targetChartId
                        });
                    }
                }
            });
        });
    }

    /**
     * 检测图表点击并获取绑定的数据
     */
    async detectChartClick() {
        // 使用动态函数生成，完全避免作用域问题和SDK错误
        const funcStr = `
            const doc = Api.GetDocument();

            try {
                // 🔒 完全内联的安全图表扫描，避免调用外部服务

                const documentLevelCharts = [];

                // 直接内联图表扫描逻辑
                try {
                    if (typeof doc.GetAllDrawingObjects === 'function') {
                        const docDrawingObjects = doc.GetAllDrawingObjects();

                        if (docDrawingObjects && docDrawingObjects.length > 0) {
                            for (let j = 0; j < docDrawingObjects.length; j++) {
                                const drawingObj = docDrawingObjects[j];

                                let drawingType = 'unknown';
                                try {
                                    if (typeof drawingObj.GetClassType === 'function') {
                                        drawingType = drawingObj.GetClassType();
                                    }
                                } catch (typeError) {
                                    drawingType = 'unknown';
                                }

                                // 检查是否是图表类型
                                if (drawingType === 'chart' ||
                                    drawingType.includes('Chart') ||
                                    drawingType.includes('Drawing') ||
                                    drawingType.includes('Shape') ||
                                    drawingType.includes('Image')) {

                                    const chartInfo = {
                                        element: drawingObj,
                                        elementType: drawingType,
                                        index: 'doc_drawing_' + j,
                                        drawingIndex: j,
                                        isDocumentLevel: true,
                                        source: 'document-level'
                                    };

                                    documentLevelCharts.push(chartInfo);
                                }
                            }
                        }
                    } else {
                        console.log('文档不支持GetAllDrawingObjects方法');
                    }
                } catch (scanError) {
                    console.log('🚨 图表扫描失败:', scanError.message);
                }

                const chartDetectionResults = [];

                // 处理找到的图表 - 完全安全模式
                documentLevelCharts.forEach(function(chartInfo) {

                    try {
                        // 完全安全的图表类型定义
                        let detailedChartType = {
                            category: 'chart',
                            specificType: 'chart_detected',
                            description: '图表（超级安全检测）',
                            properties: {
                                superSafeMode: true,
                                noApiCalls: true
                            },
                            confidence: 0.8
                        };

                        // 基于字符串的简单类型判断
                        if (chartInfo.elementType && chartInfo.elementType.toLowerCase().includes('chart')) {
                            detailedChartType.specificType = 'chart_confirmed';
                            detailedChartType.description = '图表（已确认）';
                            detailedChartType.confidence = 0.9;
                        }


                        // 生成稳定的唯一标识符（基于图表索引）
                        const uniqueId = 'chart_safe_' + chartInfo.drawingIndex + '_stable';

                        const chartResult = {
                            chartIndex: chartInfo.index,
                            chartType: chartInfo.elementType,
                            detailedChartType: detailedChartType,
                            uniqueId: uniqueId,
                            boundData: null,
                            hasBindingData: false,
                            isDocumentLevel: true,
                            drawingIndex: chartInfo.drawingIndex,
                            superSafeMode: true
                        };

                        // 🔒 从文档自定义属性获取绑定数据而不是内存存储

                        // 生成图表指纹（与plugin-bridge.js保持完全一致）
                        var fingerprint = generateChartFingerprint(drawingObj, doc);

                        // 辅助函数：生成图表指纹（与plugin-bridge.js保持完全一致）
                        function generateChartFingerprint(sel, doc) {
                            var parts = [];

                            // 1) 精确位置索引（最重要的稳定唯一性标识）
                            var exactIndex = -1;
                            try {
                                var all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
                                if (all && all.length) {
                                    for (var i = 0; i < all.length; i++) {
                                        if (all[i] === sel) {
                                            exactIndex = i;
                                            break;
                                        }
                                    }
                                }
                            } catch(_){}

                            // 索引是最重要的，必须包含（除非真的找不到）
                            if (exactIndex >= 0) {
                                parts.push('idx:' + exactIndex); // 精确索引，稳定且唯一
                            } else {
                                // 如果找不到精确索引，尝试其他方式生成唯一标识
                                try {
                                    if (all && all.length) {
                                        // 基于对象引用生成一个稳定的标识
                                        for (var j = 0; j < all.length; j++) {
                                            if (all[j] === sel) {
                                                parts.push('idx:' + j);
                                                break;
                                            }
                                        }
                                        // 如果还是没找到，使用序列号
                                        if (parts.length === 0) {
                                            parts.push('seq:' + simpleHash(sel ? sel.toString() : 'unknown'));
                                        }
                                    }
                                } catch(_){
                                    parts.push('noidx:' + simpleHash(sel ? sel.toString() : 'unknown'));
                                }
                            }

                            // 2) 尝试获取内部ID（稳定的内部标识）
                            try {
                                if (sel && typeof sel.GetId === 'function') {
                                    var stableId = sel.GetId();
                                    if (stableId) parts.push('id:' + stableId);
                                }
                            } catch(e){}

                            // 3) 图表类型（稳定）
                            var type = getChartType(sel);
                            if (type) parts.push('type:' + type);

                            // 4) 类名链（稳定）
                            try {
                                var anc = collectAncestorChain(sel, 8);
                                if (anc && anc.length) parts.push('chain:' + anc.slice(0,3).join('-'));
                            } catch(_){}

                            // 5) 宿主段落文本（稳定的上下文）
                            try {
                                var host = findHostParagraph(sel, 24);
                                if (host && typeof host.GetText === 'function') {
                                    var t = '';
                                    try { t = host.GetText() || ''; } catch(_){}
                                    if (t) {
                                        // 使用文本的前20字符和长度（稳定）
                                        var textHash = t.substring(0,20).replace(/\s+/g, '_');
                                        parts.push('text:' + textHash + ':len' + t.length);
                                    }
                                }
                            } catch(_){}

                            // 6) 图表尺寸信息（相对稳定）
                            try {
                                if (sel && typeof sel.GetWidth === 'function' && typeof sel.GetHeight === 'function') {
                                    var w = sel.GetWidth();
                                    var h = sel.GetHeight();
                                    if (w && h) {
                                        parts.push('size:' + Math.floor(w) + 'x' + Math.floor(h));
                                    }
                                }
                            } catch(_){}

                            // 7) 只有在缺乏足够信息时才使用备选方案
                            if (!parts.length || (parts.length === 1 && parts[0].indexOf('type:') === 0)) {
                                // 只有图表类型信息不够，需要更多稳定的标识符
                                try {
                                    // 使用图表在文档中的总数量作为上下文
                                    var all = doc.GetAllDrawingObjects ? doc.GetAllDrawingObjects() : null;
                                    if (all && all.length) {
                                        parts.push('total:' + all.length);
                                    }
                                } catch(_){}

                                // 增加更多区分因素确保唯一性
                                try {
                                    // 尝试获取对象的字符串表示
                                    if (sel && typeof sel.toString === 'function') {
                                        var objStr = sel.toString();
                                        if (objStr && objStr !== '[object Object]') {
                                            var objHash = simpleHash(objStr);
                                            parts.push('obj:' + objHash);
                                        }
                                    }

                                    // 尝试获取类名作为区分
                                    if (sel && sel.constructor && sel.constructor.name) {
                                        parts.push('class:' + sel.constructor.name);
                                    }

                                    // 尝试获取更多属性
                                    if (sel && typeof sel.GetClassType === 'function') {
                                        var classType = sel.GetClassType();
                                        if (classType) parts.push('ctype:' + classType);
                                    }

                                } catch(_){}

                                // 最后的备选方案：使用一个基于内容的哈希值（但不包含时间）
                                if (!parts.length) {
                                    parts.push('fallback:' + (type || 'unknown') + ':count:' + (all ? all.length : 0));
                                }
                            }

                            var fingerprint = parts.join('|');
                            console.log('🔖 生成稳定指纹:', fingerprint);
                            return fingerprint;

                            // 简单哈希函数
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
                        }

                        // 辅助函数：获取图表类型
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

                        // 辅助函数：查找宿主段落
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

                        // 辅助函数：收集祖先链
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
                        var bindingKey = 'chart-binding:' + fingerprint;

                        try {
                            var props = doc.GetCustomProperties();
                            var bindingValue = props.Get(bindingKey);

                            if (bindingValue) {
                                var binding = JSON.parse(bindingValue);
                                console.log('✅ 从文档属性找到绑定数据:', binding);

                                chartResult.boundData = binding;
                                chartResult.hasBindingData = true;
                                chartResult.bindingMethod = 'document-properties';

                                // 记录找到的重要信息
                                console.log('📋 绑定数据包含字段:', Object.keys(binding));
                                if (binding.tag) {
                                    console.log('🔖 找到tag数据:', binding.tag);
                                }
                                if (binding.chartType) {
                                    console.log('📊 找到图表类型:', binding.chartType);
                                }
                                if (binding.rid) {
                                    console.log('🆔 找到rid:', binding.rid);
                                }
                            } else {
                                console.log('⚠️ 未找到绑定数据，bindingKey:', bindingKey);
                            }
                        } catch (bindingError) {
                            console.log('❌ 读取绑定数据失败:', bindingError.message);
                        }

                        if (!chartResult.hasBindingData) {
                            console.log('⚠️ 图表 ' + chartInfo.index + ' 未找到绑定数据');
                        }

                        chartDetectionResults.push(chartResult);

                    } catch (processError) {

                        // 添加一个最基础的记录
                        chartDetectionResults.push({
                            chartIndex: chartInfo.index,
                            chartType: chartInfo.elementType,
                            detailedChartType: {
                                category: 'chart',
                                specificType: 'error_safe',
                                description: '图表（错误安全模式）',
                                confidence: 0.5
                            },
                            uniqueId: 'chart_error_' + chartInfo.drawingIndex + '_stable',
                            boundData: null,
                            hasBindingData: false,
                            isDocumentLevel: true,
                            drawingIndex: chartInfo.drawingIndex,
                            errorSafeMode: true
                        });
                    }
                });


                // 选择返回的图表 - 优先有数据的图表
                let targetChart = null;
                if (chartDetectionResults.length > 0) {
                    // 优先选择有绑定数据的图表
                    for (let i = chartDetectionResults.length - 1; i >= 0; i--) {
                        if (chartDetectionResults[i].hasBindingData) {
                            targetChart = chartDetectionResults[i];
                            break;
                        }
                    }
                    // 如果没有有数据的图表，选择最后一个
                    if (!targetChart) {
                        targetChart = chartDetectionResults[chartDetectionResults.length - 1];
                    }
                }

                // 返回结果
                if (targetChart && targetChart.hasBindingData) {
                    // 提取图表类型信息
                    const chartTypeInfo = {
                        chartType: targetChart.detailedChartType?.specificType || targetChart.elementType || 'unknown',
                        category: targetChart.detailedChartType?.category || 'unknown',
                        description: targetChart.detailedChartType?.description || '未知图表类型',
                        confidence: targetChart.detailedChartType?.confidence || 0.8
                    };

                    const result = {
                        success: true,
                        message: 'Chart with bound data detected in super safe mode',
                        data: {
                            clickType: 'chart',
                            chartInfo: targetChart,
                            chartTypeInfo: chartTypeInfo,
                            boundData: targetChart.boundData.boundData || targetChart.boundData,
                            // 包含完整的绑定信息，特别是tag数据
                            bindingMetadata: {
                                bindingId: targetChart.boundData.chartId,
                                boundAt: targetChart.boundData.createdAt,
                                bindingMethod: targetChart.bindingMethod,
                                tag: targetChart.boundData.tag,
                                tagData: targetChart.boundData.tagData,
                                chartType: targetChart.boundData.chartType,
                                rid: targetChart.boundData.rid,
                                fingerprint: targetChart.boundData.fingerprint
                            },
                            detectionSummary: {
                                totalChartsFound: chartDetectionResults.length,
                                chartsWithData: chartDetectionResults.filter(c => c.hasBindingData).length,
                                safeMode: true
                            },
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };

                    console.log('✅ 超级安全模式检测到图表点击，包含绑定数据!');
                    console.log('📊 图表标题:', result.data.boundData?.title);
                    console.log('📊 图表类型:', result.data.chartTypeInfo?.chartType);
                    console.log('📊 数据源:', result.data.boundData?.dataSource);
                    console.log('🔖 Tag数据:', result.data.bindingMetadata?.tag);
                    console.log('🆔 RID:', result.data.bindingMetadata?.rid);

                    return result;

                } else if (chartDetectionResults.length > 0) {
                    // 有图表但没有绑定数据 - 也要返回图表类型信息
                    const chart = targetChart || chartDetectionResults[0];
                    const chartTypeInfo = {
                        chartType: chart.detailedChartType?.specificType || chart.elementType || 'unknown',
                        category: chart.detailedChartType?.category || 'unknown',
                        description: chart.detailedChartType?.description || '未知图表类型',
                        confidence: chart.detailedChartType?.confidence || 0.8
                    };

                    return {
                        success: true,
                        message: 'Chart detected but no bound data',
                        data: {
                            clickType: 'chart',
                            chartInfo: chart,
                            chartTypeInfo: chartTypeInfo,
                            boundData: null,
                            detectionSummary: {
                                totalChartsFound: chartDetectionResults.length,
                                chartsWithData: 0,
                                safeMode: true
                            },
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };

                } else {
                    return {
                        success: false,
                        error: 'No charts found in document',
                        data: {
                            clickType: 'other',
                            safeMode: true,
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };
                }

            } catch (error) {
                return {
                    success: false,
                    error: error.message,
                    data: {
                        safeMode: true,
                        timestamp: new Date().toLocaleString('zh-CN')
                    }
                };
            }
        `;

        // 使用new Function执行完全隔离的安全代码
        const dynamicFunction = new Function(funcStr);
        return this.editor.runInDoc(dynamicFunction, { async: false });
    }

    /**
     * 获取当前选中图表的类型信息
     * @returns {Object} 图表类型信息
     */
    async getChartType() {
        console.log('📈 开始获取图表类型信息...');

        try {
            const result = await this.editor.runInDoc(function() {
                console.log('🏁 图表类型检测开始执行');

                const doc = Api.GetDocument();
                if (!doc) {
                    return { success: false, error: '无法获取文档对象' };
                }

                // 检查是否选中图表
                const selected = doc.GetSelectedDrawings ? doc.GetSelectedDrawings() : null;
                if (!selected || !selected.length) {
                    return { success: false, error: '请先选中一个图表' };
                }

                const chart = selected[selected.length - 1];
                console.log('✅ 发现选中的图表');

                // 生成图表指纹（与plugin-bridge.js保持完全一致）
                function buildChartFingerprint(sel, doc) {
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
                        var type = detectChartType(sel);
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

                // 辅助函数：查找宿主段落
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

                // 辅助函数：收集祖先链
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

                // 检测图表类型
                function detectChartType(chart) {
                    try {
                        var ch = (chart && typeof chart.GetChart === 'function') ? chart.GetChart() : null;
                        if (ch && typeof ch.GetChartType === 'function') {
                            return ch.GetChartType();
                        }
                        if (chart && typeof chart.GetChartType === 'function') {
                            return chart.GetChartType();
                        }
                    } catch (e) {
                        console.log('图表类型检测失败:', e);
                    }
                    return null;
                }

                // 检查绑定状态
                function checkBindingStatus(doc, fingerprint) {
                    try {
                        var props = doc.GetCustomProperties();
                        var key = 'chart-binding:' + fingerprint;
                        var val = props.Get(key);
                        if (val) {
                            try { return JSON.parse(val); } catch(_){ return null; }
                        }
                    } catch(e) {}
                    return null;
                }

                const fingerprint = buildChartFingerprint(chart, doc);
                const chartType = detectChartType(chart);
                const bindingData = checkBindingStatus(doc, fingerprint);

                console.log('📊 图表类型:', chartType);
                console.log('🔖 图表指纹:', fingerprint);

                // 构建详细信息
                const detailedInfo = {
                    category: chartType ? (chartType.includes('bar') ? '柱状图' :
                             chartType.includes('line') ? '线图' :
                             chartType.includes('pie') ? '饼图' :
                             chartType.includes('area') ? '面积图' : '其他') : '未知',
                    specificType: chartType || '未知',
                    description: chartType ? `检测到${chartType}类型图表` : '无法确定图表具体类型',
                    confidence: chartType ? 0.9 : 0.3
                };

                return {
                    success: true,
                    chartType: chartType || 'unknown',
                    fingerprint: fingerprint,
                    timestamp: new Date().toLocaleString('zh-CN'),
                    detailedInfo: detailedInfo,
                    bindingInfo: {
                        isBound: !!bindingData,
                        dataType: bindingData?.data?.type || null,
                        boundAt: bindingData?.createdAt || null
                    }
                };

            }, { async: false });

            console.log('📈 图表类型检测完成:', result);

            if (result && result.success) {
                return {
                    success: true,
                    data: result,
                    message: '图表类型检测成功'
                };
            } else {
                return {
                    success: false,
                    error: result?.error || '图表类型检测失败',
                    data: {
                        timestamp: new Date().toLocaleString('zh-CN')
                    }
                };
            }

        } catch (error) {
            console.error('❌ 图表类型检测异常:', error);
            return {
                success: false,
                error: '图表类型检测过程中发生异常: ' + error.message,
                data: {
                    timestamp: new Date().toLocaleString('zh-CN')
                }
            };
        }
    }

    /**
     * 清理临时数据
     */
    cleanupTempData() {
        if (this.dataBinder && typeof this.dataBinder.cleanupTempData === 'function') {
            this.dataBinder.cleanupTempData();
        }
    }
}