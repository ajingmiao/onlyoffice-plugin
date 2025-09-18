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
     * @param {Object} chartData - 图表数据
     */
    async bindDataToChart(chartData) {
        console.log('开始图表数据绑定流程');
        console.log('从宿主页面接收到的图表数据:', chartData);

        // 检查是否从宿主页面传递了有效数据（在runInDoc外部检查）
        if (!chartData || !chartData.data || !chartData.metadata) {
            console.log('❌ 未收到有效的图表数据，无法进行绑定');
            return {
                success: false,
                error: '未从宿主页面接收到有效的图表数据，请确保调用时传递了完整的图表数据结构',
                data: {
                    timestamp: new Date().toLocaleString()
                }
            };
        }

        console.log('✅ 使用宿主页面传递的图表数据');

        // 由于runInDoc是隔离环境，将数据直接嵌入到函数体中
        const chartDataJson = JSON.stringify(chartData);
        console.log('🔍 准备嵌入的数据JSON长度:', chartDataJson.length);

        // 创建包含数据的函数字符串，然后用new Function执行
        const funcStr = `
            const doc = Api.GetDocument();
            console.log('=== 图表数据绑定开始 ===');

            try {
                // 直接使用嵌入的数据
                const bindingData = ${chartDataJson};
                console.log('📊 使用嵌入的绑定数据:', bindingData);

                if (!bindingData || !bindingData.data || !bindingData.metadata) {
                    console.log('❌ 嵌入的数据无效');
                    return {
                        success: false,
                        error: '嵌入的图表数据无效',
                        data: {
                            timestamp: new Date().toLocaleString()
                        }
                    };
                }

                console.log('📊 最终使用的绑定数据:', bindingData);

                // 直接在内部实现图表扫描逻辑
                console.log('开始扫描文档中的图表...');
                const documentLevelCharts = [];

                try {
                    if (typeof doc.GetAllDrawingObjects === 'function') {
                        const docDrawingObjects = doc.GetAllDrawingObjects();
                        console.log('📄 文档级绘图对象:', docDrawingObjects);

                        if (docDrawingObjects && docDrawingObjects.length > 0) {
                            console.log('🎯 找到 ' + docDrawingObjects.length + ' 个文档级绘图对象！');

                            for (let j = 0; j < docDrawingObjects.length; j++) {
                                const drawingObj = docDrawingObjects[j];
                                console.log('📊 绘图对象 ' + j + ':', drawingObj);

                                let drawingType = 'unknown';
                                if (typeof drawingObj.GetClassType === 'function') {
                                    drawingType = drawingObj.GetClassType();
                                }
                                console.log('📊 绘图对象类型:', drawingType);

                                // 检查是否是图表类型
                                if (drawingType === 'chart' ||
                                    drawingType.includes('Chart') ||
                                    drawingType.includes('Drawing') ||
                                    drawingType.includes('Shape') ||
                                    drawingType.includes('Image')) {

                                    console.log('✅ 发现文档级图表/图形:', drawingType);

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
                } catch (docError) {
                    console.log('检查文档级绘图对象失败:', docError);
                }

                console.log('扫描结果 - 文档级图表数:', documentLevelCharts.length);

                const boundCharts = [];
                const bindingResults = [];

                // 处理找到的图表
                documentLevelCharts.forEach(function(chartInfo) {
                    console.log('处理图表:', chartInfo.index, chartInfo.elementType);

                    // 简化的图表类型识别（直接内联实现）
                    let detailedChartType = {
                        category: 'chart',
                        specificType: 'unknown',
                        description: '图表',
                        confidence: 0.8
                    };

                    // 尝试获取准确的图表类型
                    try {
                        // 方法1: 尝试GetPrevChart
                        if (typeof chartInfo.element.GetPrevChart === 'function') {
                            console.log('🔍 优先尝试GetPrevChart方法');

                            let prevChart;
                            try {
                                prevChart = chartInfo.element.GetPrevChart();
                                console.log('📊 GetPrevChart返回:', prevChart);
                            } catch (sdkError) {
                                console.log('🚨 GetPrevChart触发SDK错误:', sdkError.message);
                                prevChart = null;
                            }

                            if (prevChart && typeof prevChart.GetChartType === 'function') {
                                try {
                                    const chartType = prevChart.GetChartType();
                                    console.log('📊 GetChartType返回:', chartType);
                                    if (chartType) {
                                        console.log('✅ 通过GetPrevChart获得准确图表类型:', chartType);
                                        detailedChartType.specificType = chartType;
                                        detailedChartType.description = '图表 (' + chartType + ')';
                                        detailedChartType.confidence = 1.0;
                                    } else {
                                        console.log('⚠️ GetChartType返回空值');
                                    }
                                } catch (chartTypeError) {
                                    console.log('🚨 GetChartType调用失败:', chartTypeError.message);
                                }
                            } else {
                                console.log('⚠️ prevChart为空，尝试其他方法');
                            }
                        }

                        // 方法2: 如果GetPrevChart失败，尝试直接调用GetChart
                        if (detailedChartType.specificType === 'unknown' && typeof chartInfo.element.GetChart === 'function') {
                            console.log('🔍 尝试GetChart方法');
                            try {
                                const chart = chartInfo.element.GetChart();
                                console.log('📊 GetChart返回:', chart);
                                if (chart && typeof chart.GetChartType === 'function') {
                                    const chartType = chart.GetChartType();
                                    console.log('📊 Chart.GetChartType返回:', chartType);
                                    if (chartType) {
                                        console.log('✅ 通过GetChart获得准确图表类型:', chartType);
                                        detailedChartType.specificType = chartType;
                                        detailedChartType.description = '图表 (' + chartType + ')';
                                        detailedChartType.confidence = 1.0;
                                    }
                                }
                            } catch (chartError) {
                                console.log('🚨 GetChart调用失败:', chartError.message);
                            }
                        }

                        // 方法2.5: 尝试直接调用图表元素的GetChartType方法
                        if (detailedChartType.specificType === 'unknown' && typeof chartInfo.element.GetChartType === 'function') {
                            console.log('🔍 尝试直接调用图表元素的GetChartType方法');
                            try {
                                const chartType = chartInfo.element.GetChartType();
                                console.log('📊 直接GetChartType返回:', chartType);
                                if (chartType && chartType !== 'chart') {
                                    console.log('✅ 通过直接GetChartType获得准确图表类型:', chartType);
                                    detailedChartType.specificType = chartType;
                                    detailedChartType.description = '图表 (' + chartType + ')';
                                    detailedChartType.confidence = 1.0;
                                }
                            } catch (chartTypeError) {
                                console.log('🚨 直接GetChartType调用失败:', chartTypeError.message);
                            }
                        }

                    } catch (chartTypeError) {
                        console.log('🚨 图表类型识别失败:', chartTypeError.message);
                    }

                    console.log('📊 图表类型识别结果:', detailedChartType);

                    // 生成唯一标识符
                    const uniqueId = 'chart_' + chartInfo.drawingIndex + '_' + Date.now();

                    // 创建绑定信息
                    const bindingInfo = {
                        chartIndex: chartInfo.index,
                        chartType: chartInfo.elementType,
                        detailedChartType: detailedChartType,
                        uniqueId: uniqueId,
                        boundData: bindingData.data || {},
                        bindingId: 'doc_chart_' + chartInfo.drawingIndex + '_' + Date.now(),
                        boundAt: new Date().toISOString(),
                        metadata: bindingData.metadata || {},
                        isDocumentLevel: true,
                        drawingIndex: chartInfo.drawingIndex
                    };

                    // 简化的数据绑定（使用内存存储）
                    let bindingResult = {
                        directBinding: false,
                        bindingMethod: 'memory-storage',
                        storageKey: null
                    };

                    try {
                        // 尝试直接绑定
                        if (typeof chartInfo.element.SetCustomProperty === 'function') {
                            chartInfo.element.SetCustomProperty('chartData', JSON.stringify(bindingInfo));
                            bindingResult.directBinding = true;
                            bindingResult.bindingMethod = 'custom-property';
                            console.log('✅ 直接在图表元素上绑定数据成功');
                        } else {
                            // 使用内存存储
                            if (!window.chartDataStorage) {
                                window.chartDataStorage = {};
                            }
                            const storageKey = 'doc_chart_' + chartInfo.drawingIndex + '_' + Date.now();
                            window.chartDataStorage[storageKey] = bindingInfo;
                            bindingResult.storageKey = storageKey;
                            console.log('✅ 使用内存存储绑定图表数据，存储键:', storageKey);
                        }
                    } catch (bindingError) {
                        console.log('数据绑定失败:', bindingError.message);
                    }

                    // 合并绑定结果
                    Object.assign(bindingInfo, bindingResult);

                    boundCharts.push(chartInfo);
                    bindingResults.push(bindingInfo);

                    console.log('🎯 图表绑定完成:', {
                        uniqueId: uniqueId,
                        chartType: detailedChartType.description,
                        specificType: detailedChartType.specificType,
                        bindingMethod: bindingResult.bindingMethod
                    });

                    // 验证绑定：立即获取刚绑定的数据并打印
                    console.log('🔍 验证数据绑定 - 开始获取绑定数据...');
                    try {
                        let retrievedData = null;
                        let retrievalMethod = null;

                        // 方法1：从自定义属性获取
                        if (bindingResult.bindingMethod === 'custom-property') {
                            if (typeof chartInfo.element.GetCustomProperty === 'function') {
                                const customData = chartInfo.element.GetCustomProperty('chartData');
                                if (customData) {
                                    retrievedData = JSON.parse(customData);
                                    retrievalMethod = 'custom-property';
                                    console.log('✅ 从自定义属性成功获取绑定数据');
                                }
                            }
                        }

                        // 方法2：从内存存储获取
                        if (!retrievedData && bindingResult.bindingMethod === 'memory-storage' && bindingResult.storageKey) {
                            if (window.chartDataStorage && window.chartDataStorage[bindingResult.storageKey]) {
                                retrievedData = window.chartDataStorage[bindingResult.storageKey];
                                retrievalMethod = 'memory-storage';
                                console.log('✅ 从内存存储成功获取绑定数据，存储键:', bindingResult.storageKey);
                            }
                        }

                        if (retrievedData) {
                            console.log('📋 验证成功！获取到的绑定数据:');
                            console.log('   - 获取方法:', retrievalMethod);
                            console.log('   - 图表标题:', retrievedData.boundData?.title || '未知');
                            console.log('   - 图表类型:', retrievedData.boundData?.type || '未知');
                            console.log('   - 数据源:', retrievedData.boundData?.dataSource || '未知');
                            console.log('   - 绑定ID:', retrievedData.bindingId);
                            console.log('   - 绑定时间:', retrievedData.boundAt);

                            // 打印关键指标
                            if (retrievedData.boundData?.metrics) {
                                console.log('   - 关键指标:');
                                const metrics = retrievedData.boundData.metrics;
                                if (metrics.totalSales) console.log('     * 总销售额: ¥' + metrics.totalSales.toLocaleString());
                                if (metrics.growthRate) console.log('     * 增长率: ' + metrics.growthRate + '%');
                                if (metrics.topProduct) console.log('     * 热门产品: ' + metrics.topProduct);
                                if (metrics.targetAchievement) console.log('     * 目标达成: ' + metrics.targetAchievement + '%');
                            }

                            // 打印数据系列预览
                            if (retrievedData.boundData?.series && retrievedData.boundData.series.length > 0) {
                                console.log('   - 数据系列:');
                                retrievedData.boundData.series.forEach(function(series) {
                                    console.log('     * ' + series.name + ': [' + series.data.slice(0, 3).join(', ') +
                                              (series.data.length > 3 ? '...' : '') + ']');
                                });
                            }

                            console.log('   - 完整数据对象:', retrievedData);
                        } else {
                            console.log('⚠️ 数据绑定验证失败：无法获取刚绑定的数据');
                            console.log('   - 尝试的方法:', bindingResult.bindingMethod);
                            console.log('   - 存储键:', bindingResult.storageKey || '无');
                        }
                    } catch (verifyError) {
                        console.log('❌ 数据绑定验证出错:', verifyError.message);
                    }
                });

                // 如果没有找到图表，提供指导
                if (boundCharts.length === 0) {
                    console.log('💡 提示: 当前文档中没有检测到图表元素');
                }

                return {
                    success: true,
                    message: boundCharts.length > 0 ?
                        ('成功绑定 ' + boundCharts.length + ' 个图表') :
                        ('文档扫描完成，但未发现图表元素。'),
                    data: {
                        chartsFound: boundCharts.length,
                        bindingResults: bindingResults,
                        boundData: bindingData,
                        bindingMethod: 'inline-processing',
                        timestamp: new Date().toLocaleString()
                    }
                };

            } catch (error) {
                console.log('❌ 图表数据绑定失败:', error);
                return {
                    success: false,
                    error: error.message,
                    data: {
                        timestamp: new Date().toLocaleString()
                    }
                };
            }
        `;

        // 使用new Function执行包含数据的代码
        const dynamicFunction = new Function(funcStr);
        return this.editor.runInDoc(dynamicFunction);
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
                            chartTypeInfo: chartTypeInfo,  // 新增图表类型信息
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
                            chartTypeInfo: chartTypeInfo,  // 新增图表类型信息
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
     * 获取文档中所有图表的绑定数据摘要
     */
    async getChartBindingSummary() {
        // 在runInDoc外部获取服务引用，避免作用域问题
        const chartDetector = this.chartDetector;
        const dataBinder = this.dataBinder;

        return new Promise((resolve) => {
            this.editor.runInDoc(() => {
                const doc = Api.GetDocument();
                console.log('=== 获取图表绑定摘要 ===');

                try {
                    const scanResults = chartDetector.scanDocument(doc);
                    const summary = {
                        totalCharts: 0,
                        chartsWithData: 0,
                        bindingSummary: []
                    };

                    // 统计文档级图表
                    scanResults.documentLevelCharts.forEach((chartInfo, index) => {
                        summary.totalCharts++;
                        const dataResult = dataBinder.getBoundData(chartInfo.element, chartInfo.drawingIndex);

                        if (dataResult.hasBindingData) {
                            summary.chartsWithData++;
                        }

                        summary.bindingSummary.push({
                            chartIndex: chartInfo.index,
                            chartType: chartInfo.elementType,
                            hasBindingData: dataResult.hasBindingData,
                            bindingPreview: dataResult.boundData?.bindingId || 'unknown',
                            source: 'document-level'
                        });
                    });

                    resolve({
                        success: true,
                        data: summary,
                        timestamp: new Date().toLocaleString('zh-CN')
                    });

                } catch (error) {
                    resolve({
                        success: false,
                        error: error.message
                    });
                }

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    /**
     * 清理临时数据
     */
    cleanupTempData() {
        this.dataBinder.cleanupTempData();
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

                // 生成图表指纹（复用之前的逻辑）
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
}