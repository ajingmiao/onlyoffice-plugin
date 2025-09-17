// services/element-detection-service.js
export class ElementDetectionService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 通用元素检测方法
    async detectClickedElement() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();

                console.log('=== 通用元素检测 ===');
                console.log('当前选区:', range);

                if (!range) {
                    console.log('未检测到选区');
                    resolve({
                        success: false,
                        error: 'No selection found',
                        data: {
                            clickType: 'document',
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    });
                    return;
                }

                // 获取选区信息
                var rangeText = '';
                if (typeof range.GetText === 'function') {
                    rangeText = range.GetText();
                }

                console.log('选区文本内容:', rangeText);

                // 尝试检测各种元素类型
                var detectionResults = {
                    tables: [],
                    paragraphs: [],
                    shapes: [],
                    images: [],
                    textRuns: []
                };

                // 扫描文档元素
                console.log('开始扫描文档元素...');
                for (var i = 0; i < 100; i++) {
                    try {
                        var element = doc.GetElement(i);
                        if (!element) {
                            console.log(`文档元素 ${i}: null，结束遍历`);
                            break;
                        }

                        var elementType = 'unknown';
                        if (typeof element.GetClassType === 'function') {
                            elementType = element.GetClassType();
                        }

                        console.log(`文档元素 ${i}: ${elementType}`);

                        // 检测表格
                        if (elementType === 'CTable') {
                            var tableInfo = {
                                index: i,
                                type: 'table',
                                rows: 0,
                                columns: 0,
                                content: [],
                                clickedCell: null
                            };

                            if (typeof element.GetRowsCount === 'function') {
                                tableInfo.rows = element.GetRowsCount();

                                // 尝试检测点击的具体单元格
                                for (var r = 0; r < Math.min(tableInfo.rows, 20); r++) {
                                    try {
                                        var row = element.GetRow(r);
                                        if (row && typeof row.GetCellsCount === 'function') {
                                            var cellCount = row.GetCellsCount();
                                            tableInfo.columns = Math.max(tableInfo.columns, cellCount);

                                            var rowData = [];
                                            for (var c = 0; c < Math.min(cellCount, 20); c++) {
                                                try {
                                                    var cell = row.GetCell(c);
                                                    var cellText = '';

                                                    if (cell && typeof cell.GetContent === 'function') {
                                                        var cellContent = cell.GetContent();
                                                        if (cellContent) {
                                                            // 提取单元格文本内容
                                                            if (typeof cellContent.GetElementsCount === 'function') {
                                                                var elemCount = cellContent.GetElementsCount();
                                                                for (var e = 0; e < elemCount; e++) {
                                                                    var cellElem = cellContent.GetElement(e);
                                                                    if (cellElem && typeof cellElem.GetText === 'function') {
                                                                        cellText += cellElem.GetText();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    rowData.push(cellText.trim());
                                                } catch (cellError) {
                                                    console.log(`读取单元格 [${r}][${c}] 出错:`, cellError);
                                                    rowData.push('ERROR');
                                                }
                                            }
                                            if (r < 5) { // 只保存前5行以避免过多数据
                                                tableInfo.content.push(rowData);
                                            }
                                        }
                                    } catch (rowError) {
                                        console.log(`读取表格第 ${r} 行出错:`, rowError);
                                    }
                                }
                            }

                            detectionResults.tables.push(tableInfo);
                        }

                        // 检测段落
                        else if (elementType === 'CDocumentParagraph') {
                            var paraInfo = {
                                index: i,
                                type: 'paragraph',
                                text: '',
                                alignment: 'unknown',
                                hasElements: false
                            };

                            try {
                                if (typeof element.GetText === 'function') {
                                    paraInfo.text = element.GetText();
                                }
                                if (typeof element.GetJc === 'function') {
                                    paraInfo.alignment = element.GetJc() || 'left';
                                }
                                if (typeof element.GetElementsCount === 'function') {
                                    paraInfo.hasElements = element.GetElementsCount() > 0;
                                }
                            } catch (paraError) {
                                console.log(`读取段落 ${i} 出错:`, paraError);
                            }

                            detectionResults.paragraphs.push(paraInfo);
                        }

                        // 检测图形/形状
                        else if (elementType.includes('Shape') || elementType.includes('Drawing')) {
                            var shapeInfo = {
                                index: i,
                                type: 'shape',
                                elementType: elementType,
                                hasContent: false
                            };

                            try {
                                if (typeof element.GetDocContent === 'function') {
                                    var shapeContent = element.GetDocContent();
                                    shapeInfo.hasContent = !!shapeContent;
                                }
                            } catch (shapeError) {
                                console.log(`读取图形 ${i} 出错:`, shapeError);
                            }

                            detectionResults.shapes.push(shapeInfo);
                        }

                    } catch (elementError) {
                        console.log(`读取文档元素 ${i} 出错:`, elementError);
                    }
                }

                // 确定主要的检测结果
                var primaryDetection = null;
                var clickType = 'unknown';
                var elementInfo = {};

                // 优先处理表格点击
                if (detectionResults.tables.length > 0) {
                    var mainTable = detectionResults.tables[detectionResults.tables.length - 1]; // 取最后一个表格
                    clickType = 'table';
                    primaryDetection = mainTable;

                    // 增强表格信息
                    elementInfo = {
                        elementType: 'table',
                        tableIndex: mainTable.index,
                        totalRows: mainTable.rows,
                        totalColumns: mainTable.columns,
                        clickedCell: mainTable.clickedCell,
                        sampleContent: mainTable.content.slice(0, 3), // 显示前3行内容
                        metadata: {
                            detectedAt: new Date().toLocaleString('zh-CN'),
                            detectionMethod: 'element-scan',
                            tablesInDocument: detectionResults.tables.length
                        }
                    };

                    if (mainTable.clickedCell) {
                        elementInfo.clickedCellInfo = {
                            row: mainTable.clickedCell.rowIndex,
                            column: mainTable.clickedCell.columnIndex,
                            content: mainTable.content[mainTable.clickedCell.row] ?
                                   mainTable.content[mainTable.clickedCell.row][mainTable.clickedCell.column] : 'N/A'
                        };
                    }
                }
                // 处理段落点击
                else if (detectionResults.paragraphs.length > 0) {
                    var mainParagraph = detectionResults.paragraphs[detectionResults.paragraphs.length - 1];
                    clickType = 'paragraph';
                    primaryDetection = mainParagraph;

                    elementInfo = {
                        elementType: 'paragraph',
                        paragraphIndex: mainParagraph.index,
                        text: mainParagraph.text.substring(0, 100) + (mainParagraph.text.length > 100 ? '...' : ''),
                        alignment: mainParagraph.alignment,
                        hasElements: mainParagraph.hasElements,
                        metadata: {
                            detectedAt: new Date().toLocaleString('zh-CN'),
                            detectionMethod: 'element-scan',
                            paragraphsInDocument: detectionResults.paragraphs.length,
                            fullTextLength: mainParagraph.text.length
                        }
                    };
                }
                // 处理图形点击
                else if (detectionResults.shapes.length > 0) {
                    var mainShape = detectionResults.shapes[detectionResults.shapes.length - 1];
                    clickType = 'shape';
                    primaryDetection = mainShape;

                    elementInfo = {
                        elementType: 'shape',
                        shapeIndex: mainShape.index,
                        shapeType: mainShape.elementType,
                        hasContent: mainShape.hasContent,
                        metadata: {
                            detectedAt: new Date().toLocaleString('zh-CN'),
                            detectionMethod: 'element-scan',
                            shapesInDocument: detectionResults.shapes.length
                        }
                    };
                }
                // 通用文档点击
                else {
                    clickType = 'document';
                    elementInfo = {
                        elementType: 'document',
                        hasSelection: !!range,
                        selectionInfo: range ? {
                            text: rangeText,
                            hasText: rangeText.length > 0
                        } : null,
                        metadata: {
                            detectedAt: new Date().toLocaleString('zh-CN'),
                            detectionMethod: 'element-scan',
                            totalElementsScanned: detectionResults.tables.length +
                                                 detectionResults.paragraphs.length +
                                                 detectionResults.shapes.length
                        }
                    };
                }

                var result = {
                    success: true,
                    message: `Detected ${clickType} click`,
                    data: {
                        clickType: clickType,
                        elementInfo: elementInfo,
                        primaryDetection: primaryDetection,
                        fullScanResults: {
                            tablesFound: detectionResults.tables.length,
                            paragraphsFound: detectionResults.paragraphs.length,
                            shapesFound: detectionResults.shapes.length
                        },
                        timestamp: new Date().toLocaleString('zh-CN')
                    }
                };

                console.log('✅ 元素检测完成:', result);
                resolve(result);

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    // 专门的表格行列检测方法（改进版）
    async detectTableCellClick() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();

                console.log('=== 精确表格单元格检测 ===');

                if (!range) {
                    resolve({
                        success: false,
                        error: 'No selection found for table cell detection'
                    });
                    return;
                }

                // 尝试更精确的单元格检测
                try {
                    // 方法1: 直接从选区获取父元素
                    var selection = doc.GetSelection();
                    if (selection && typeof selection.GetParent === 'function') {
                        var parent = selection.GetParent();
                        console.log('选区父元素:', parent);

                        // 检查父元素是否是单元格
                        if (parent && typeof parent.GetClassType === 'function') {
                            var parentType = parent.GetClassType();
                            console.log('父元素类型:', parentType);

                            if (parentType === 'CTableCell') {
                                // 找到单元格，现在需要确定其在表格中的位置
                                console.log('✅ 直接检测到单元格点击');

                                // 获取单元格所在的行
                                var cellRow = parent.GetParent();
                                if (cellRow && typeof cellRow.GetClassType === 'function' &&
                                    cellRow.GetClassType() === 'CTableRow') {

                                    // 获取行所在的表格
                                    var table = cellRow.GetParent();
                                    if (table && typeof table.GetClassType === 'function' &&
                                        table.GetClassType() === 'CTable') {

                                        // 计算行列位置
                                        var rowIndex = -1;
                                        var columnIndex = -1;

                                        // 查找行位置
                                        for (var r = 0; r < table.GetRowsCount(); r++) {
                                            if (table.GetRow(r) === cellRow) {
                                                rowIndex = r;
                                                break;
                                            }
                                        }

                                        // 查找列位置
                                        if (rowIndex >= 0) {
                                            for (var c = 0; c < cellRow.GetCellsCount(); c++) {
                                                if (cellRow.GetCell(c) === parent) {
                                                    columnIndex = c;
                                                    break;
                                                }
                                            }
                                        }

                                        // 获取单元格内容
                                        var cellContent = '';
                                        if (typeof parent.GetContent === 'function') {
                                            var content = parent.GetContent();
                                            if (content && typeof content.GetText === 'function') {
                                                cellContent = content.GetText();
                                            }
                                        }

                                        resolve({
                                            success: true,
                                            message: 'Precise table cell click detected',
                                            data: {
                                                clickType: 'table-cell',
                                                cellPosition: {
                                                    row: rowIndex + 1, // 用户友好的1基索引
                                                    column: columnIndex + 1,
                                                    rawRow: rowIndex,
                                                    rawColumn: columnIndex
                                                },
                                                cellContent: cellContent.trim(),
                                                tableInfo: {
                                                    totalRows: table.GetRowsCount(),
                                                    totalColumns: cellRow.GetCellsCount()
                                                },
                                                detectionMethod: 'direct-cell-detection',
                                                timestamp: new Date().toLocaleString('zh-CN')
                                            }
                                        });
                                        return;
                                    }
                                }
                            }
                        }
                    }
                } catch (directError) {
                    console.log('直接检测方法失败:', directError);
                }

                // 方法2: 回退到通用检测
                console.log('回退到通用表格检测方法...');
                resolve({
                    success: false,
                    error: 'Direct cell detection failed, falling back to general detection'
                });

            }, { async: false, cb: (res) => resolve(res) });
        });
    }
}