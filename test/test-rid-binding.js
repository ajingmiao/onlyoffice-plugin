// 测试Word文档r:id绑定功能
export function testRidBinding() {
    console.log('=== 测试Word文档r:id绑定功能 ===');

    const testData = {
        chartId: 'test-chart-' + Date.now(),
        type: 'chart',
        rid: 'test-rid-123',
        chartType: 'barChart',
        tag: JSON.stringify({
            trackId: 'test-track-id',
            groupFields: ['field1', 'field2'],
            _t: Date.now()
        })
    };

    console.log('📊 测试数据:', testData);
    console.log('🎯 预期结果: 绑定数据应包含documentRId字段');
    console.log('🔖 新的指纹格式: 现在使用"rid:rId18"格式，而不是复杂的指纹算法');
    console.log('📋 使用方法:');
    console.log('1. 在OnlyOffice中插入一个图表');
    console.log('2. 选中图表');
    console.log('3. 调用BIND_CHART_DATA命令与上述测试数据');
    console.log('4. 检查返回的绑定数据是否包含documentRId');
    console.log('5. 检查自定义属性key是否为"chart-binding:rid:rId18"格式');

    return testData;
}

// 模拟绑定数据结构（包含新的documentRId字段和r:id指纹）
export const expectedBindingStructure = {
    chartId: 'chart-mfrvxntu-7jt6',
    createdAt: '2025-09-20T06:26:10.530Z',
    chartType: 'comboBarLine',
    fingerprint: 'rid:rId18', // 新的指纹格式：直接使用r:id
    documentRId: 'rId18', // 新增的Word文档r:id字段
    boundData: {
        type: 'chart',
        rid: 'test-rid-123',
        chartType: 'barChart',
        tag: '{"trackId":"test-track-id","groupFields":["field1","field2"],"_t":1726825570530}',
        tagData: {
            trackId: 'test-track-id',
            groupFields: ['field1', 'field2'],
            _t: 1726825570530
        },
        originalFormat: 'flat'
    },
    metadata: {
        sourceFormat: 'flat',
        convertedAt: '2025-09-20T06:26:10.530Z',
        trackId: 'test-track-id',
        groupFields: ['field1', 'field2'],
        timestamp: 1726825570530
    },
    lastUpdated: '2025-09-20T06:26:10.530Z',
    tag: '{"trackId":"test-track-id","groupFields":["field1","field2"],"_t":1726825570530}',
    tagData: {
        trackId: 'test-track-id',
        groupFields: ['field1', 'field2'],
        _t: 1726825570530
    },
    rid: 'test-rid-123'
};

// 对比新旧指纹方法的优势
export const fingerprintComparison = {
    oldMethod: {
        example: 'type:comboBarLine|chain:chart',
        issues: [
            '依赖图表类型，可能不稳定',
            '基于复杂算法，难以预测',
            '与Word文档结构无直接关联',
            '在POI读取时需要复杂匹配'
        ]
    },
    newMethod: {
        example: 'rid:rId18',
        advantages: [
            '直接使用Word文档关系ID',
            '与document.xml.rels中的r:id直接对应',
            '可以直接映射到charts/chartX.xml文件',
            'POI读取时可以精确匹配',
            '更简洁、更可靠的标识方法'
        ]
    }
};

console.log('📝 预期绑定数据结构:', expectedBindingStructure);
console.log('🔄 指纹方法对比:', fingerprintComparison);