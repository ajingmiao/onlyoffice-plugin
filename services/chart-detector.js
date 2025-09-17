// services/chart-detector.js
export class ChartDetector {
    constructor() {
        // 可能的图表元素类型
        this.chartElementTypes = [
            'Drawing', 'Chart', 'CDrawing', 'GraphicFrame',
            'Shape', 'Image', 'Ole', 'Inline', 'Float',
            'Anchor', 'DocContent', 'Run'
        ];
    }

    /**
     * 扫描文档中的所有图表
     * @param {Object} doc - OnlyOffice文档对象
     * @returns {Object} 扫描结果
     */
    scanDocument(doc) {
        console.log('=== 开始扫描文档图表 ===');

        const scanResults = {
            documentLevelCharts: [],
            elementLevelCharts: [],
            scanStatistics: {
                totalElementsScanned: 0,
                elementTypeCounts: {},
                hasCharts: false
            },
            allElementDetails: []
        };

        // 1. 扫描文档级绘图对象（主要方法）
        this._scanDocumentLevelDrawings(doc, scanResults);

        // 2. 扫描元素级图表（备用方法）
        this._scanElementLevelCharts(doc, scanResults);

        // 3. 更新统计信息
        scanResults.scanStatistics.hasCharts =
            scanResults.documentLevelCharts.length > 0 ||
            scanResults.elementLevelCharts.length > 0;

        console.log('📊 文档扫描完成统计:');
        console.log('- 文档级图表数:', scanResults.documentLevelCharts.length);
        console.log('- 元素级图表数:', scanResults.elementLevelCharts.length);
        console.log('- 总扫描元素数:', scanResults.scanStatistics.totalElementsScanned);
        console.log('- 元素类型分布:', scanResults.scanStatistics.elementTypeCounts);

        return scanResults;
    }

    /**
     * 获取当前选区信息
     * @param {Object} doc - OnlyOffice文档对象
     * @returns {Object} 选区信息
     */
    getCurrentSelection(doc) {
        try {
            const selection = doc.GetRangeBySelect();
            console.log('🎯 当前选区:', selection);

            const selectionInfo = {
                hasSelection: !!selection,
                selection: selection
            };

            if (selection) {
                try {
                    if (typeof selection.GetDrawingObjects === 'function') {
                        const selectionDrawings = selection.GetDrawingObjects();
                        selectionInfo.drawingObjects = selectionDrawings;
                        console.log('🎯 选区中的绘图对象:', selectionDrawings);
                    } else {
                        console.log('选区不支持GetDrawingObjects方法');
                    }
                } catch (selectionDrawingError) {
                    console.log('获取选区绘图对象失败:', selectionDrawingError.message);
                }
            } else {
                console.log('无法获取当前选区');
            }

            return selectionInfo;
        } catch (selectionError) {
            console.log('检查选区绘图对象失败:', selectionError.message);
            return { hasSelection: false, error: selectionError.message };
        }
    }

    /**
     * 扫描文档级绘图对象
     */
    _scanDocumentLevelDrawings(doc, scanResults) {
        try {
            if (typeof doc.GetAllDrawingObjects === 'function') {
                const docDrawingObjects = doc.GetAllDrawingObjects();
                console.log('📄 文档级绘图对象:', docDrawingObjects);

                if (docDrawingObjects && docDrawingObjects.length > 0) {
                    console.log('🎯 找到 ' + docDrawingObjects.length + ' 个文档级绘图对象！');

                    for (let j = 0; j < docDrawingObjects.length; j++) {
                        const drawingObj = docDrawingObjects[j];
                        console.log('📊 绘图对象 ' + j + ':', drawingObj);

                        const drawingType = this._getElementType(drawingObj);
                        console.log('📊 绘图对象类型:', drawingType);

                        // 检查是否是图表类型
                        if (this._isChartType(drawingType)) {
                            console.log('✅ 发现文档级图表/图形:', drawingType);

                            const chartInfo = {
                                element: drawingObj,
                                elementType: drawingType,
                                index: 'doc_drawing_' + j,
                                drawingIndex: j,
                                isDocumentLevel: true,
                                source: 'document-level'
                            };

                            scanResults.documentLevelCharts.push(chartInfo);
                        }
                    }
                }
            } else {
                console.log('文档不支持GetAllDrawingObjects方法');
            }
        } catch (docError) {
            console.log('检查文档级绘图对象失败:', docError);
        }
    }

    /**
     * 扫描元素级图表
     */
    _scanElementLevelCharts(doc, scanResults) {
        for (let i = 0; i < 100; i++) {
            try {
                const element = doc.GetElement(i);
                if (!element) {
                    console.log('文档元素 ' + i + ': null，结束遍历');
                    break;
                }

                scanResults.scanStatistics.totalElementsScanned++;
                const elementType = this._getElementType(element);

                // 统计元素类型
                scanResults.scanStatistics.elementTypeCounts[elementType] =
                    (scanResults.scanStatistics.elementTypeCounts[elementType] || 0) + 1;

                // 记录详细信息用于调试
                const elementDetail = {
                    index: i,
                    type: elementType,
                    methods: this._getAvailableMethods(element, 10)
                };

                scanResults.allElementDetails.push(elementDetail);
                console.log('文档元素 ' + i + ': ' + elementType);

                // 检查是否是图表元素
                if (this._isChartElement(elementType)) {
                    console.log('✅ 发现潜在图表/图形在位置 ' + i + ': ' + elementType);
                    console.log('🔍 元素详细信息:', elementDetail);

                    // 深度检查元素
                    this._performDeepElementCheck(element);

                    const chartInfo = {
                        element: element,
                        elementType: elementType,
                        index: i,
                        isDocumentLevel: false,
                        source: 'element-level',
                        debugInfo: elementDetail
                    };

                    scanResults.elementLevelCharts.push(chartInfo);
                }

            } catch (elementError) {
                console.log('读取文档元素 ' + i + ' 出错:', elementError);
            }
        }
    }

    /**
     * 获取元素类型
     */
    _getElementType(element) {
        if (typeof element.GetClassType === 'function') {
            return element.GetClassType();
        }
        return 'unknown';
    }

    /**
     * 检查是否是图表类型（严格检查）
     */
    _isChartType(elementType) {
        return elementType === 'chart' ||
               elementType.includes('Chart') ||
               elementType.includes('Drawing') ||
               elementType.includes('Shape') ||
               elementType.includes('Image');
    }

    /**
     * 检查是否是图表元素（宽泛检查）
     */
    _isChartElement(elementType) {
        return this.chartElementTypes.some(type =>
            elementType.includes(type) ||
            elementType.toLowerCase().includes(type.toLowerCase())
        ) || elementType.toLowerCase().includes('chart') ||
           elementType.toLowerCase().includes('graphic');
    }

    /**
     * 获取可用方法列表
     */
    _getAvailableMethods(element, limit = 10) {
        try {
            const methods = [];
            for (const prop in element) {
                if (typeof element[prop] === 'function') {
                    methods.push(prop);
                }
            }
            return methods.slice(0, limit);
        } catch (methodError) {
            return [];
        }
    }

    /**
     * 执行深度元素检查
     */
    _performDeepElementCheck(element) {
        try {
            // 检查是否有GetDrawingObjects方法
            if (typeof element.GetDrawingObjects === 'function') {
                const drawingObjects = element.GetDrawingObjects();
                console.log('🎨 绘图对象:', drawingObjects);
            }

            // 检查是否有GetAllDrawingObjects方法
            if (typeof element.GetAllDrawingObjects === 'function') {
                const allDrawingObjects = element.GetAllDrawingObjects();
                console.log('🎨 所有绘图对象:', allDrawingObjects);
            }

            // 检查是否有内容
            if (typeof element.GetElementsCount === 'function') {
                const elementsCount = element.GetElementsCount();
                console.log('📊 子元素数量:', elementsCount);
            }

            // 检查内容
            if (typeof element.GetContent === 'function') {
                const content = element.GetContent();
                console.log('📄 元素内容:', content);
            }

            // 检查文档内容
            if (typeof element.GetDocContent === 'function') {
                const docContent = element.GetDocContent();
                console.log('📄 文档内容:', docContent);
            }
        } catch (deepError) {
            console.log('深度检查失败:', deepError);
        }
    }

    /**
     * 提供检测指导
     */
    provideDetectionGuidance() {
        console.log('💡 提示: 当前文档中没有检测到图表元素');
        console.log('💡 可能原因:');
        console.log('  1. 图表可能存储在不同的位置（如段落内的内联对象）');
        console.log('  2. OnlyOffice的图表可能需要不同的API访问方式');
        console.log('  3. 图表可能需要保存后才能通过API访问');
        console.log('  4. 图表可能在文档的不同层级结构中');
        console.log('💡 建议:');
        console.log('  - 尝试保存文档后重新检测');
        console.log('  - 插入图表后点击图表区域，然后运行检测');
        console.log('  - 检查控制台是否有其他类型的元素出现');
    }
}