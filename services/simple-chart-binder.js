// services/simple-chart-binder.js - 简化的图表绑定器
export class SimpleChartBinder {
    /**
     * 简单安全的图表数据绑定（纯内存模式）
     * @param {Object} chartData - 图表数据
     */
    static bindChartDataSafe(chartData) {
        console.log('🎯 === 开始简单安全图表数据绑定 ===');
        console.log('📊 接收图表数据:', chartData);

        try {
            // 检查数据有效性
            if (!chartData || !chartData.data || !chartData.metadata) {
                console.log('❌ 图表数据无效');
                return {
                    success: false,
                    error: '图表数据无效'
                };
            }

            // 初始化存储
            if (!window.chartDataStorage) {
                window.chartDataStorage = {};
            }

            // 生成存储键
            const timestamp = Date.now();
            const storageKey = `memory_chart_${timestamp}`;

            // 创建绑定信息
            const bindingInfo = {
                chartIndex: 'memory_chart_0',
                chartType: 'chart',
                detailedChartType: {
                    category: 'chart',
                    specificType: 'memory_bound',
                    description: '内存绑定图表',
                    confidence: 1.0
                },
                uniqueId: `chart_memory_${timestamp}`,
                boundData: chartData.data,
                bindingId: `memory_binding_${timestamp}`,
                boundAt: new Date().toISOString(),
                metadata: chartData.metadata,
                isDocumentLevel: true,
                drawingIndex: 0,
                bindingMethod: 'memory-storage',
                safeMode: true,
                createdAt: new Date().toLocaleString('zh-CN')
            };

            // 存储数据
            window.chartDataStorage[storageKey] = bindingInfo;

            console.log('✅ 图表数据绑定成功');
            console.log('📊 存储键:', storageKey);
            console.log('📊 图表标题:', bindingInfo.boundData?.title);

            return {
                success: true,
                message: '成功在内存中绑定图表数据（简单安全模式）',
                data: {
                    storageKey: storageKey,
                    bindingInfo: bindingInfo,
                    timestamp: new Date().toLocaleString('zh-CN')
                }
            };

        } catch (error) {
            console.log('❌ 图表数据绑定失败:', error.message);
            return {
                success: false,
                error: error.message
            };
        }
    }
}