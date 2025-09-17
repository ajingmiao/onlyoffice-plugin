// services/chart-data-binder.js
export class ChartDataBinder {
    constructor() {
        this.storagePrefix = 'doc_chart_';
        this.fallbackPrefix = 'doc_chart_fallback_';
    }

    /**
     * 为图表绑定数据
     * @param {Object} element - 图表元素
     * @param {Object} bindingInfo - 绑定信息
     * @param {number} index - 元素索引
     * @returns {Object} 绑定结果
     */
    bindDataToChart(element, bindingInfo, index) {
        const result = {
            directBinding: false,
            markerCreated: false,
            bindingMethod: null,
            storageKey: null,
            errors: []
        };

        // 方法1: 尝试直接在元素上设置自定义属性
        this._tryDirectBinding(element, bindingInfo, result);

        // 方法2: 尝试创建内容控件标记（如果支持）
        if (typeof Api !== 'undefined') {
            this._tryContentControlBinding(bindingInfo, result);
        }

        return result;
    }

    /**
     * 从图表获取绑定数据
     * @param {Object} element - 图表元素
     * @param {number} index - 元素索引
     * @returns {Object} 绑定数据结果
     */
    getBoundData(element, index) {
        const result = {
            hasBindingData: false,
            boundData: null,
            bindingMethod: null,
            storageKey: null
        };

        // 方法1: 检查自定义属性
        this._checkCustomProperty(element, result);

        // 方法2: 检查内存存储
        if (!result.hasBindingData) {
            this._checkMemoryStorage(index, result);
        }

        return result;
    }

    /**
     * 清理临时数据
     */
    cleanupTempData() {
        if (window.tempChartData) {
            delete window.tempChartData;
        }
    }

    /**
     * 创建测试绑定数据
     */
    createTestBindingData() {
        return {
            data: {
                title: '测试图表数据',
                type: 'line-chart',
                dataSource: '测试系统',
                category: '测试分析',
                period: '2024年测试期间',
                metrics: {
                    totalSales: 1000000,
                    growthRate: 12.5,
                    topProduct: '测试产品A',
                    targetAchievement: 110
                },
                series: [
                    { name: '实际销售', data: [100, 120, 110, 140] },
                    { name: '目标销售', data: [110, 115, 120, 130] }
                ],
                categories: ['Q1', 'Q2', 'Q3', 'Q4']
            },
            metadata: {
                bindingId: 'test_chart_' + Date.now(),
                boundAt: new Date().toISOString(),
                dataVersion: '1.0',
                refreshInterval: 3600,
                lastUpdated: new Date().toLocaleString(),
                permissions: {
                    canEdit: true,
                    canRefresh: true,
                    canExport: true
                }
            }
        };
    }

    /**
     * 尝试直接绑定
     */
    _tryDirectBinding(element, bindingInfo, result) {
        try {
            if (typeof element.SetCustomProperty === 'function') {
                element.SetCustomProperty('chartData', JSON.stringify(bindingInfo));
                result.directBinding = true;
                result.bindingMethod = 'custom-property';
                console.log('✅ 直接在图表元素上绑定数据成功');
            } else {
                console.log('图表元素不支持SetCustomProperty，尝试其他方法');
                this._tryAlternativeBinding(element, bindingInfo, result);
            }
        } catch (directError) {
            console.log('直接绑定失败:', directError);
            result.errors.push({ method: 'SetCustomProperty', error: directError.message });
            this._tryAlternativeBinding(element, bindingInfo, result);
        }
    }

    /**
     * 尝试替代绑定方法
     */
    _tryAlternativeBinding(element, bindingInfo, result) {
        // 尝试其他可能的属性设置方法
        const methods = ['SetProperty', 'SetAttribute'];

        for (const method of methods) {
            if (typeof element[method] === 'function') {
                try {
                    element[method]('chartData', JSON.stringify(bindingInfo));
                    result.directBinding = true;
                    result.bindingMethod = method;
                    console.log(`✅ 使用${method}绑定数据成功`);
                    return;
                } catch (error) {
                    console.log(`${method}绑定失败:`, error.message);
                    result.errors.push({ method, error: error.message });
                }
            }
        }

        // 最后使用内存存储
        this._useMemoryStorage(bindingInfo, result);
    }

    /**
     * 使用内存存储作为备用方案
     */
    _useMemoryStorage(bindingInfo, result) {
        try {
            if (!window.chartDataStorage) {
                window.chartDataStorage = {};
            }

            const storageKey = this.storagePrefix + bindingInfo.drawingIndex + '_' + Date.now();
            window.chartDataStorage[storageKey] = bindingInfo;

            result.storageKey = storageKey;
            result.directBinding = true;
            result.bindingMethod = 'memory-storage';

            console.log('✅ 使用内存存储绑定图表数据，存储键:', storageKey);
        } catch (storageError) {
            console.log('内存存储也失败:', storageError);
            result.errors.push({ method: 'memory-storage', error: storageError.message });
        }
    }

    /**
     * 尝试创建内容控件绑定
     */
    _tryContentControlBinding(bindingInfo, result) {
        try {
            const doc = Api.GetDocument();
            if (typeof doc.AddTextContentControl === 'function') {
                const hiddenMarker = doc.AddTextContentControl();
                if (hiddenMarker) {
                    const tagData = 'doc-chart-data:' + JSON.stringify(bindingInfo);

                    if (typeof hiddenMarker.SetTag === 'function') {
                        hiddenMarker.SetTag(tagData);
                        console.log('✅ 图表数据已绑定到Tag:', tagData.substring(0, 100) + '...');
                    }

                    if (typeof hiddenMarker.SetAlias === 'function') {
                        hiddenMarker.SetAlias('图表数据: ' + bindingInfo.bindingId);
                    }

                    // 设置隐藏标记（不显示内容）
                    if (typeof hiddenMarker.SetPlaceholderText === 'function') {
                        hiddenMarker.SetPlaceholderText('');
                    }

                    if (typeof hiddenMarker.AddText === 'function') {
                        hiddenMarker.AddText(''); // 空文本，不显示
                    }

                    result.markerCreated = true;
                    result.markerMethod = 'content-control';
                }
            } else {
                console.log('文档不支持AddTextContentControl，跳过内容控件创建');
                result.markerMethod = 'not-supported';
            }
        } catch (markerError) {
            console.log('创建内容控件标记失败:', markerError);
            result.errors.push({ method: 'content-control', error: markerError.message });
        }
    }

    /**
     * 检查自定义属性
     */
    _checkCustomProperty(element, result) {
        try {
            if (typeof element.GetCustomProperty === 'function') {
                const customData = element.GetCustomProperty('chartData');
                if (customData) {
                    result.boundData = JSON.parse(customData);
                    result.hasBindingData = true;
                    result.bindingMethod = 'custom-property';
                    console.log('✅ 从自定义属性获取到图表数据:', result.boundData);
                }
            }
        } catch (customError) {
            console.log('读取自定义属性失败:', customError);
        }
    }

    /**
     * 检查内存存储
     */
    _checkMemoryStorage(index, result) {
        if (window.chartDataStorage) {
            for (const storageKey in window.chartDataStorage) {
                const storedData = window.chartDataStorage[storageKey];
                if (storedData && storedData.drawingIndex === index) {
                    result.boundData = storedData;
                    result.hasBindingData = true;
                    result.bindingMethod = 'memory-storage';
                    result.storageKey = storageKey;
                    console.log('✅ 从内存存储获取到图表数据:', result.boundData);
                    break;
                }
            }
        }
    }
}