// services/table-service.js
export class TableService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 在光标位置插入表格
    async insertTable(opts) {
        const options = opts || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            rows: options.rows || 3,
            columns: options.columns || 3,
            widthType: options.widthType || 'percent',
            width: options.width || 100,
            headers: options.headers || [],
            data: options.data || [],
            style: options.style || 'default'
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            var rows = (Asc.scope && Asc.scope.rows) || 3;
            var columns = (Asc.scope && Asc.scope.columns) || 3;
            var widthType = (Asc.scope && Asc.scope.widthType) || 'percent';
            var width = (Asc.scope && Asc.scope.width) || 100;
            var headers = (Asc.scope && Asc.scope.headers) || [];
            var data = (Asc.scope && Asc.scope.data) || [];
            var style = (Asc.scope && Asc.scope.style) || 'default';

            console.log('=== 表格插入 ===');
            console.log('参数:', { rows, columns, widthType, width, headers, data, style });

            try {
                // 方法1: 尝试使用InsertContent插入表格
                console.log('尝试方法1: InsertContent插入表格...');
                var table = Api.CreateTable(columns, rows);

                if (!table) {
                    console.log('创建表格失败');
                    return { success: false, error: 'Failed to create table' };
                }

                console.log('表格创建成功:', table);

                // 设置表格宽度
                if (typeof table.SetWidth === 'function') {
                    table.SetWidth(widthType, width);
                    console.log('设置表格宽度:', widthType, width);
                }

                // 填充表头
                if (headers.length > 0) {
                    console.log('填充表头数据...');
                    for (var j = 0; j < Math.min(headers.length, columns); j++) {
                        if (typeof table.GetRow === 'function') {
                            var row = table.GetRow(0);
                            if (row && typeof row.GetCell === 'function') {
                                var cell = row.GetCell(j);
                                if (cell && typeof cell.GetContent === 'function') {
                                    var cellContent = cell.GetContent();
                                    if (cellContent && typeof cellContent.GetElement === 'function') {
                                        var para = cellContent.GetElement(0);
                                        if (para && typeof para.AddText === 'function') {
                                            para.AddText(headers[j]);
                                            console.log(`表头[${j}]: ${headers[j]}`);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 填充数据
                if (data.length > 0) {
                    console.log('填充表格数据...');
                    var startRow = headers.length > 0 ? 1 : 0;
                    for (var i = 0; i < Math.min(data.length, rows - startRow); i++) {
                        var rowData = data[i];
                        if (Array.isArray(rowData)) {
                            for (var j = 0; j < Math.min(rowData.length, columns); j++) {
                                if (typeof table.GetRow === 'function') {
                                    var row = table.GetRow(startRow + i);
                                    if (row && typeof row.GetCell === 'function') {
                                        var cell = row.GetCell(j);
                                        if (cell && typeof cell.GetContent === 'function') {
                                            var cellContent = cell.GetContent();
                                            if (cellContent && typeof cellContent.GetElement === 'function') {
                                                var para = cellContent.GetElement(0);
                                                if (para && typeof para.AddText === 'function') {
                                                    para.AddText(String(rowData[j]));
                                                    console.log(`数据[${i}][${j}]: ${rowData[j]}`);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // === 插入真正的表格（无标记控件）===
                console.log('插入真正的表格...');
                var insertSuccess = false;

                // 先尝试插入表格
                if (typeof doc.InsertContent === 'function') {
                    console.log('尝试使用InsertContent插入表格...');
                    var insertResult = doc.InsertContent([table], false);
                    console.log('InsertContent结果:', insertResult);

                    if (insertResult) {
                        insertSuccess = true;
                        console.log('✅ 表格通过InsertContent插入成功');
                    }
                }

                if (!insertSuccess && typeof doc.Push === 'function') {
                    console.log('尝试使用Push插入表格...');
                    doc.Push(table);
                    insertSuccess = true;
                    console.log('✅ 表格通过Document.Push插入成功');
                }

                if (insertSuccess) {
                    return {
                        success: true,
                        message: 'Table inserted successfully',
                        method: 'Clean-Table-Insert',
                        parameters: { rows, columns, width: `${width}${widthType}` }
                    };
                }

                // 如果上述方法都失败了
                return { success: false, error: 'All table insertion methods failed' };

            } catch (e) {
                console.log('表格插入失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 插入动态数据表格（从宿主系统绑定）
    async insertDynamicTable(dynamicData) {
        if (!dynamicData) {
            return { success: false, error: 'No dynamic data provided' };
        }

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            title: dynamicData.title || '数据表格',
            headers: dynamicData.headers || [],
            data: dynamicData.data || [],
            metadata: dynamicData.metadata || {},
            styling: dynamicData.styling || {},
            rows: (dynamicData.data || []).length + (dynamicData.headers ? 1 : 0),
            columns: Math.max(
                (dynamicData.headers || []).length,
                Math.max(...(dynamicData.data || []).map(row => Array.isArray(row) ? row.length : 0), 0)
            )
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();
            var scope = Asc.scope;

            console.log('=== 动态数据表格插入 ===');
            console.log('标题:', scope.title);
            console.log('元数据:', scope.metadata);
            console.log('样式:', scope.styling);
            console.log('数据结构:', { rows: scope.rows, columns: scope.columns });

            try {
                // 插入标题（如果有）
                var titlePara = null;
                if (scope.title) {
                    titlePara = Api.CreateParagraph();
                    titlePara.SetSpacingAfter(200);
                    titlePara.SetJc('center');

                    var titleRun = Api.CreateRun();
                    titleRun.SetBold(true);
                    titleRun.SetFontSize(14);
                    titleRun.AddText(scope.title);
                    titlePara.AddElement(titleRun);

                    // 添加元数据信息
                    if (scope.metadata.generatedAt) {
                        titleRun.AddLineBreak();
                        var metaRun = Api.CreateRun();
                        metaRun.SetFontSize(10);
                        metaRun.SetColor(128, 128, 128);
                        metaRun.AddText('生成时间: ' + scope.metadata.generatedAt);
                        titlePara.AddElement(metaRun);
                    }

                    if (scope.metadata.dataSource) {
                        var sourceRun = Api.CreateRun();
                        sourceRun.SetFontSize(10);
                        sourceRun.SetColor(128, 128, 128);
                        sourceRun.AddText(' | 数据源: ' + scope.metadata.dataSource);
                        titlePara.AddElement(sourceRun);
                    }
                }

                // 创建表格
                var table = Api.CreateTable(scope.columns, scope.rows);
                if (!table) {
                    console.log('创建动态表格失败');
                    return { success: false, error: 'Failed to create dynamic table' };
                }

                console.log('动态表格创建成功');

                // 设置表格样式
                if (typeof table.SetWidth === 'function') {
                    table.SetWidth('percent', 100);
                }

                // 设置表格边框
                if (typeof table.SetTableBorderTop === 'function') {
                    table.SetTableBorderTop('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderBottom('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderLeft('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderRight('single', 4, 0, 0, 0, 0);
                    table.SetTableBorderInsideH('single', 2, 0, 0, 0, 0);
                    table.SetTableBorderInsideV('single', 2, 0, 0, 0, 0);
                }

                var currentRow = 0;

                // 填充表头
                if (scope.headers.length > 0) {
                    console.log('填充动态表头...');
                    var headerRow = table.GetRow(0);
                    if (headerRow) {
                        for (var j = 0; j < Math.min(scope.headers.length, scope.columns); j++) {
                            var cell = headerRow.GetCell(j);
                            if (cell) {
                                var cellContent = cell.GetContent();
                                if (cellContent) {
                                    var para = cellContent.GetElement(0);
                                    if (para) {
                                        para.AddText(String(scope.headers[j]));

                                        // 表头样式
                                        if (scope.styling.headerStyle === 'bold') {
                                            var textPr = Api.CreateTextPr();
                                            textPr.SetBold(true);
                                            para.SetTextPr(textPr);
                                        }

                                        // 表头居中
                                        para.SetJc('center');

                                        // 表头背景色
                                        if (typeof cell.SetShd === 'function') {
                                            cell.SetShd('clear', 220, 220, 220);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    currentRow = 1;
                }

                // 填充数据
                if (scope.data.length > 0) {
                    console.log('填充动态数据...');
                    for (var i = 0; i < Math.min(scope.data.length, scope.rows - currentRow); i++) {
                        var rowData = scope.data[i];
                        if (Array.isArray(rowData)) {
                            var row = table.GetRow(currentRow + i);
                            if (row) {
                                for (var j = 0; j < Math.min(rowData.length, scope.columns); j++) {
                                    var cell = row.GetCell(j);
                                    if (cell) {
                                        var cellContent = cell.GetContent();
                                        if (cellContent) {
                                            var para = cellContent.GetElement(0);
                                            if (para) {
                                                var cellText = String(rowData[j] || '');
                                                para.AddText(cellText);

                                                // 数字居右对齐
                                                if (!isNaN(parseFloat(cellText.replace(/[^\d.-]/g, '')))) {
                                                    para.SetJc('right');
                                                }

                                                // 合计行样式
                                                if (i === scope.data.length - 1 && scope.styling.totalRowStyle === 'bold') {
                                                    var textPr = Api.CreateTextPr();
                                                    textPr.SetBold(true);
                                                    para.SetTextPr(textPr);

                                                    if (typeof cell.SetShd === 'function') {
                                                        cell.SetShd('clear', 240, 240, 240);
                                                    }
                                                }

                                                // 交替行颜色
                                                else if (scope.styling.alternateRowColors && i % 2 === 1) {
                                                    if (typeof cell.SetShd === 'function') {
                                                        cell.SetShd('clear', 250, 250, 250);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 插入表格到文档
                var insertSuccess = false;

                // 先插入标题
                if (titlePara) {
                    if (typeof doc.Push === 'function') {
                        doc.Push(titlePara);
                    }
                }

                // === 插入真正的动态表格（无标记控件）===
                console.log('插入真正的动态表格...');
                var insertSuccess = false;

                // 先插入标题（如果有）
                if (titlePara) {
                    if (typeof doc.Push === 'function') {
                        doc.Push(titlePara);
                        console.log('✅ 标题段落插入成功');
                    }
                }

                // 插入表格
                if (typeof doc.InsertContent === 'function') {
                    console.log('尝试使用InsertContent插入动态表格...');
                    var insertResult = doc.InsertContent([table], false);
                    if (insertResult) {
                        insertSuccess = true;
                        console.log('✅ 动态表格通过InsertContent插入成功');
                    }
                }

                if (!insertSuccess && typeof doc.Push === 'function') {
                    doc.Push(table);
                    insertSuccess = true;
                    console.log('✅ 动态表格通过Document.Push插入成功');
                }

                if (insertSuccess) {
                    return {
                        success: true,
                        message: 'Dynamic table inserted successfully',
                        method: 'Dynamic-Data-Binding',
                        parameters: {
                            title: scope.title,
                            rows: scope.rows,
                            columns: scope.columns,
                            dataSource: scope.metadata.dataSource,
                            recordCount: scope.data.length
                        }
                    };
                } else {
                    return { success: false, error: 'Dynamic table insertion failed' };
                }

            } catch (e) {
                console.log('动态表格插入失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 插入预设样式的表格
    async insertPresetTable(preset, customData) {
        const presets = {
            'simple': {
                rows: 3,
                columns: 3,
                headers: ['列1', '列2', '列3'],
                data: [
                    ['数据1', '数据2', '数据3'],
                    ['数据4', '数据5', '数据6']
                ]
            },
            'schedule': {
                rows: 6,
                columns: 7,
                headers: ['时间', '周一', '周二', '周三', '周四', '周五', '周六'],
                data: [
                    ['9:00', '', '', '', '', '', ''],
                    ['10:00', '', '', '', '', '', ''],
                    ['11:00', '', '', '', '', '', ''],
                    ['14:00', '', '', '', '', '', ''],
                    ['15:00', '', '', '', '', '', '']
                ]
            },
            'comparison': {
                rows: 5,
                columns: 3,
                headers: ['功能', '方案A', '方案B'],
                data: [
                    ['性能', '优秀', '良好'],
                    ['成本', '高', '中'],
                    ['维护', '复杂', '简单'],
                    ['扩展性', '强', '一般']
                ]
            }
        };

        const options = { ...presets[preset] || presets['simple'], ...customData };
        return this.insertTable(options);
    }

    // 处理表格点击事件（直接检测表格，不依赖Content Control）
    async handleTableClick() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();

                console.log('=== 表格点击检测（直接检测方式）===');
                console.log('当前选区:', range);

                if (!range) {
                    console.log('未检测到选区');
                    resolve({ success: false, error: 'No selection found' });
                    return;
                }

                // 检测文档中的所有表格
                console.log('扫描文档中的所有表格...');
                var foundTables = [];

                for (var i = 0; i < 50; i++) {
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

                        if (elementType === 'CTable') {
                            console.log(`✅ 发现表格在位置 ${i}!`);

                            // 提取表格数据
                            var tableData = {
                                index: i,
                                rows: 0,
                                columns: 0,
                                content: []
                            };

                            if (typeof element.GetRowsCount === 'function') {
                                tableData.rows = element.GetRowsCount();
                                console.log(`表格 ${i} 行数:`, tableData.rows);

                                // 提取表格内容（前5行）
                                for (var r = 0; r < Math.min(tableData.rows, 5); r++) {
                                    try {
                                        if (typeof element.GetRow === 'function') {
                                            var row = element.GetRow(r);
                                            if (row && typeof row.GetCellsCount === 'function') {
                                                var cellCount = row.GetCellsCount();
                                                tableData.columns = Math.max(tableData.columns, cellCount);

                                                var rowData = [];
                                                for (var c = 0; c < Math.min(cellCount, 10); c++) {
                                                    try {
                                                        if (typeof row.GetCell === 'function') {
                                                            var cell = row.GetCell(c);
                                                            var cellText = '';

                                                            if (cell && typeof cell.GetContent === 'function') {
                                                                var cellContent = cell.GetContent();
                                                                if (cellContent && typeof cellContent.GetElementsCount === 'function') {
                                                                    var elemCount = cellContent.GetElementsCount();
                                                                    for (var e = 0; e < elemCount; e++) {
                                                                        var cellElem = cellContent.GetElement(e);
                                                                        if (cellElem && typeof cellElem.GetText === 'function') {
                                                                            cellText += cellElem.GetText();
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            rowData.push(cellText.trim());
                                                        }
                                                    } catch (cellError) {
                                                        console.log(`读取单元格 [${r}][${c}] 出错:`, cellError);
                                                        rowData.push('ERROR');
                                                    }
                                                }
                                                tableData.content.push(rowData);
                                                console.log(`表格 ${i} 第 ${r} 行:`, rowData);
                                            }
                                        }
                                    } catch (rowError) {
                                        console.log(`读取表格第 ${r} 行出错:`, rowError);
                                    }
                                }
                            }

                            foundTables.push({ index: i, table: element, data: tableData });
                        }
                    } catch (elementError) {
                        console.log(`读取文档元素 ${i} 出错:`, elementError);
                    }
                }

                console.log(`总共发现 ${foundTables.length} 个表格`);

                if (foundTables.length > 0) {
                    // 简单策略：返回最后一个表格的信息（假设用户点击的是最近的表格）
                    var lastTable = foundTables[foundTables.length - 1];

                    // 尝试检测表格的类型（通过分析表格内容）
                    var detectedType = 'unknown-table';
                    var metadata = {
                        detectedAt: new Date().toLocaleString('zh-CN'),
                        method: 'content-analysis'
                    };

                    // 简单的表格类型识别
                    if (lastTable.data.content && lastTable.data.content.length > 0) {
                        var firstRow = lastTable.data.content[0];
                        if (firstRow && firstRow.length > 0) {
                            var headerText = firstRow.join('|').toLowerCase();

                            if (headerText.includes('产品') && headerText.includes('月')) {
                                detectedType = 'sales-report';
                                metadata.title = '销售报表';
                                metadata.category = '业务数据';
                            } else if (headerText.includes('时间') && headerText.includes('周')) {
                                detectedType = 'schedule';
                                metadata.title = '时间表';
                                metadata.category = '日程安排';
                            } else if (firstRow.length >= 3) {
                                detectedType = 'data-table';
                                metadata.title = '数据表格';
                                metadata.category = '通用表格';
                            }
                        }
                    }

                    var result = {
                        success: true,
                        message: 'Table detected in document',
                        data: {
                            clickType: 'table',
                            tableIndex: lastTable.index,
                            tableData: lastTable.data,
                            detectedType: detectedType,
                            metadata: metadata,
                            detectionMethod: 'direct-table-scan',
                            tablesFound: foundTables.length,
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    };

                    console.log('✅ 表格检测成功:', result);
                    resolve(result);
                    return;
                } else {
                    console.log('❌ 未发现任何表格');
                    resolve({
                        success: false,
                        error: 'No tables found in document',
                        data: {
                            clickType: 'document',
                            elementsScanned: 50,
                            hasRange: !!range,
                            timestamp: new Date().toLocaleString('zh-CN')
                        }
                    });
                }

            }, { async: false, cb: (res) => resolve(res) });
        });
    }
}