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
     * @param {Object} chartData - 包含chartId和数据的对象 {chartId: "xxx", data: {...}, metadata: {...}}
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

        if (!chartData.data || !chartData.metadata) {
            console.log('❌ 未收到有效的图表数据，无法进行绑定');
            return {
                success: false,
                error: '未从宿主页面接收到有效的图表数据，请确保调用时传递了完整的数据结构',
                data: {
                    timestamp: new Date().toLocaleString()
                }
            };
        }

        const targetChartId = chartData.chartId;
        console.log('🎯 目标图表ID:', targetChartId);

        // 由于runInDoc是隔离环境，将数据直接嵌入到函数体中
        const chartDataJson = JSON.stringify(chartData);
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
                updatedBinding.boundData = bindingData.data;
                updatedBinding.metadata = bindingData.metadata;
                updatedBinding.lastUpdated = new Date().toISOString();

                // 保存更新后的绑定
                props.Add(bindingKey, JSON.stringify(updatedBinding));

                console.log('✅ 图表数据绑定完成');

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

            // 辅助函数：生成图表指纹（简化版）
            function generateChartFingerprint(sel, doc) {
                // 简化的指纹生成逻辑
                var parts = [];

                // 获取图表类型
                var type = 'unknown';
                try {
                    if (sel && typeof sel.GetChartType === 'function') {
                        type = sel.GetChartType() || 'unknown';
                    }
                } catch (e) {
                    // 忽略错误
                }

                if (type) parts.push('type:' + type);

                // 获取类名
                try {
                    var className = sel ? sel.GetClassType() : null;
                    if (className) parts.push('chain:' + className);
                } catch (e) {
                    // 忽略错误
                }

                if (!parts.length) parts.push('rand:' + Date.now().toString(36));
                return parts.join('|');
            }
        */}).toString().replace(/^function\s*\(\)\s*\{\/\*|\*\/\}\s*$/g, '');

        console.log('🔍 准备执行图表绑定沙箱代码...');

        // 将数据设置到scope中，使用与plugin-bridge.js一致的方式
        try {
            if (!window.Asc) window.Asc = {};
            if (!window.Asc.scope) window.Asc.scope = {};
            window.Asc.scope._cb_chartData = chartData;
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
            this.editor.runInDoc(new Function(funcStr), (result) => {
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
            });
        });
    }

    /**
     * 检测图表点击并获取绑定的数据
     */
    async detectChartClick() {
        console.log('开始图表点击检测流程');

        // 使用动态函数生成，完全避免作用域问题和SDK错误
        const funcStr = `
            const doc = Api.GetDocument();
            console.log('=== 图表点击检测（超级安全模式）===');

            try {
                // 🔒 完全内联的安全图表扫描，避免调用外部服务
                console.log('🔒 开始超级安全模式图表扫描...');

                const documentLevelCharts = [];

                // 直接内联图表扫描逻辑
                try {
                    if (typeof doc.GetAllDrawingObjects === 'function') {
                        const docDrawingObjects = doc.GetAllDrawingObjects();
                        console.log('📄 文档级绘图对象数量:', docDrawingObjects ? docDrawingObjects.length : 0);

                        if (docDrawingObjects && docDrawingObjects.length > 0) {
                            for (let j = 0; j < docDrawingObjects.length; j++) {
                                const drawingObj = docDrawingObjects[j];
                                console.log('📊 检查绘图对象 ' + j);

                                let drawingType = 'unknown';
                                try {
                                    if (typeof drawingObj.GetClassType === 'function') {
                                        drawingType = drawingObj.GetClassType();
                                    }
                                } catch (typeError) {
                                    console.log('🚨 获取类型失败，使用默认类型');
                                    drawingType = 'unknown';
                                }

                                console.log('📊 绘图对象类型:', drawingType);

                                // 检查是否是图表类型
                                if (drawingType === 'chart' ||
                                    drawingType.includes('Chart') ||
                                    drawingType.includes('Drawing') ||
                                    drawingType.includes('Shape') ||
                                    drawingType.includes('Image')) {

                                    console.log('✅ 发现图表/图形:', drawingType);

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

                console.log('🔍 扫描结果 - 找到图表数量:', documentLevelCharts.length);

                const chartDetectionResults = [];

                // 处理找到的图表 - 完全安全模式
                documentLevelCharts.forEach(function(chartInfo) {
                    console.log('🔒 安全处理图表:', chartInfo.index);

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

                        console.log('📊 超级安全图表类型:', detailedChartType);

                        // 生成安全的唯一标识符
                        const uniqueId = 'chart_safe_' + chartInfo.drawingIndex + '_' + Date.now();

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

                        // 🔒 超级安全的数据检查 - 只访问内存存储
                        console.log('🔒 安全检查内存中的绑定数据...');
                        console.log('🔍 当前window.chartDataStorage状态:', window.chartDataStorage);

                        if (window.chartDataStorage) {
                            console.log('📦 内存存储中的所有键:', Object.keys(window.chartDataStorage));
                            console.log('🎯 当前图表drawingIndex:', chartInfo.drawingIndex);

                            for (const storageKey in window.chartDataStorage) {
                                const storedData = window.chartDataStorage[storageKey];
                                console.log('🔍 检查存储键:', storageKey, '数据drawingIndex:', storedData?.drawingIndex);

                                if (storedData && storedData.drawingIndex === chartInfo.drawingIndex) {
                                    chartResult.boundData = storedData;
                                    chartResult.hasBindingData = true;
                                    chartResult.bindingMethod = 'memory-storage';
                                    console.log('✅ 从内存找到绑定数据:', {
                                        uniqueId: storedData.uniqueId,
                                        title: storedData.boundData?.title
                                    });
                                    break;
                                }
                            }
                        } else {
                            console.log('❌ window.chartDataStorage 不存在!');
                        }

                        if (!chartResult.hasBindingData) {
                            console.log('⚠️ 图表 ' + chartInfo.index + ' 未找到绑定数据');
                        }

                        chartDetectionResults.push(chartResult);

                    } catch (processError) {
                        console.log('🚨 处理图表失败，添加基础记录:', processError.message);

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
                            uniqueId: 'chart_error_' + chartInfo.drawingIndex + '_' + Date.now(),
                            boundData: null,
                            hasBindingData: false,
                            isDocumentLevel: true,
                            drawingIndex: chartInfo.drawingIndex,
                            errorSafeMode: true
                        });
                    }
                });

                console.log('🔍 图表检测结果汇总:', {
                    total: chartDetectionResults.length,
                    withData: chartDetectionResults.filter(c => c.hasBindingData).length
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
                            bindingMetadata: {
                                bindingId: targetChart.boundData.bindingId,
                                boundAt: targetChart.boundData.boundAt,
                                bindingMethod: targetChart.bindingMethod
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

                    console.log('⚠️ 检测到图表但无绑定数据');
                    console.log('📊 图表类型:', chartTypeInfo.chartType);

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
                    // 没有检测到图表
                    console.log('⚠️ 文档中未检测到图表');
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
                console.log('❌ 超级安全模式图表点击检测失败:', error);
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
        return this.editor.runInDoc(dynamicFunction);
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

                // 生成图表指纹
                function buildChartFingerprint(sel, doc) {
                    var stableId = null;
                    try { if (sel && typeof sel.GetId === 'function') stableId = sel.GetId(); } catch(e){}
                    if (stableId) return 'id:' + stableId;

                    var parts = [];

                    // 获取图表类型
                    try {
                        var ch = (sel && typeof sel.GetChart === 'function') ? sel.GetChart() : null;
                        if (ch && typeof ch.GetChartType === 'function') {
                            var type = ch.GetChartType();
                            if (type) parts.push('type:' + type);
                        }
                    } catch (e) {}

                    // 祖先链
                    try {
                        var cur = sel, chain = [], hop = 0;
                        while (cur && hop < 8) {
                            chain.push(cur.GetClassType ? cur.GetClassType() : '?');
                            if (!cur.GetParent) break;
                            try { cur = cur.GetParent(); } catch (e) { break; }
                            hop++;
                        }
                        if (chain.length) parts.push('chain:' + chain.slice(0,3).join('-'));
                    } catch(_){}

                    if (!parts.length) parts.push('rand:' + Date.now().toString(36));
                    return parts.join('|');
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