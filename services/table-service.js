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

    // 处理表格点击事件：精确定位当前单元格并返回稳定标识
    async handleTableClick() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                function safeCall(target, methodName, args) {
                    if (!target || typeof target[methodName] !== 'function') {
                        return { ok: false };
                    }
                    try {
                        return { ok: true, value: target[methodName].apply(target, args || []) };
                    } catch (err) {
                        console.log('[TableService] 调用', methodName, '出错:', err);
                        return { ok: false, error: err };
                    }
                }

                function parsePrefixedJsonTag(tag, prefix) {
                    if (typeof tag !== 'string' || tag.indexOf(prefix) !== 0) {
                        return null;
                    }
                    try {
                        return JSON.parse(tag.slice(prefix.length));
                    } catch (e) {
                        console.log('[TableService] 解析标签失败:', tag, e);
                        return null;
                    }
                }

                function generateUid(prefix) {
                    return (prefix || 'uid') + '-' + (new Date().getTime().toString(36)) + '-' + Math.floor(Math.random() * 1e8).toString(36);
                }

                function findCellFromRange(range) {
                    if (!range) {
                        return null;
                    }

                    var directCell = safeCall(range, 'GetParentTableCell');
                    if (directCell.ok && directCell.value) {
                        return directCell.value;
                    }

                    var paragraph = null;
                    var paraResult = safeCall(range, 'GetParagraph');
                    if (paraResult.ok) {
                        paragraph = paraResult.value;
                    }

                    if (!paragraph && typeof range.GetElementsCount === 'function' && typeof range.GetElement === 'function') {
                        try {
                            var elemCount = range.GetElementsCount();
                            for (var i = 0; i < elemCount; i++) {
                                var elem = range.GetElement(i);
                                if (!elem) {
                                    continue;
                                }
                                if (typeof elem.GetParentParagraph === 'function') {
                                    try {
                                        paragraph = elem.GetParentParagraph();
                                        if (paragraph) {
                                            break;
                                        }
                                    } catch (errPara) {
                                        console.log('[TableService] GetParentParagraph 失败:', errPara);
                                    }
                                }
                            }
                        } catch (errElements) {
                            console.log('[TableService] 遍历range元素失败:', errElements);
                        }
                    }

                    if (paragraph && typeof paragraph.GetParentTableCell === 'function') {
                        try {
                            var cellFromPara = paragraph.GetParentTableCell();
                            if (cellFromPara) {
                                return cellFromPara;
                            }
                        } catch (errCell) {
                            console.log('[TableService] 从段落获取单元格失败:', errCell);
                        }
                    }

                    if (typeof range.GetElement === 'function') {
                        try {
                            var firstElement = range.GetElement(0);
                            if (firstElement && typeof firstElement.GetParentTableCell === 'function') {
                                return firstElement.GetParentTableCell();
                            }
                        } catch (errFirstElem) {
                            console.log('[TableService] 从首元素查找单元格失败:', errFirstElem);
                        }
                    }

                    return null;
                }

                function ensureTableMeta(table) {
                    var info = {
                        table: table,
                        control: null,
                        createdControl: false,
                        controlId: '',
                        internalId: '',
                        tag: '',
                        metaObj: null,
                        tableUid: '',
                        descriptionMeta: '',
                        tableIndex: -1
                    };

                    if (!table) {
                        return info;
                    }

                    var parentControl = safeCall(table, 'GetParentContentControl');
                    if (parentControl.ok && parentControl.value) {
                        info.control = parentControl.value;
                    }

                    if (!info.control) {
                        var insertResult = safeCall(table, 'InsertInContentControl', [1]);
                        if (insertResult.ok && insertResult.value) {
                            info.control = insertResult.value;
                            info.createdControl = true;
                        }
                    }

                    if (info.control) {
                        var controlIdRes = safeCall(info.control, 'GetId');
                        if (controlIdRes.ok && controlIdRes.value) {
                            info.controlId = controlIdRes.value;
                        }

                        var internalIdRes = safeCall(info.control, 'GetInternalId');
                        if (internalIdRes.ok && internalIdRes.value) {
                            info.internalId = internalIdRes.value;
                        }

                        var tagRes = safeCall(info.control, 'GetTag');
                        if (tagRes.ok && typeof tagRes.value === 'string') {
                            info.tag = tagRes.value;
                            var parsed = parsePrefixedJsonTag(info.tag, 'table-meta:');
                            if (parsed && parsed.tableUid) {
                                info.tableUid = parsed.tableUid;
                            }
                            if (parsed) {
                                info.metaObj = parsed;
                            }
                        }

                        if (!info.tableUid) {
                            var newMeta = {
                                tableUid: generateUid('table'),
                                createdAt: new Date().toISOString()
                            };
                            info.tableUid = newMeta.tableUid;
                            info.tag = 'table-meta:' + JSON.stringify(newMeta);
                            info.metaObj = newMeta;
                            if (typeof info.control.SetTag === 'function') {
                                try {
                                    info.control.SetTag(info.tag);
                                } catch (setTagErr) {
                                    console.log('[TableService] 设置表格控件Tag失败:', setTagErr);
                                }
                            }
                        }
                    }

                    if (!info.tableUid) {
                        var descRes = safeCall(table, 'GetTableDescription');
                        if (descRes.ok && typeof descRes.value === 'string') {
                            info.descriptionMeta = descRes.value;
                            var parsedDesc = parsePrefixedJsonTag(descRes.value, 'table-meta:');
                            if (parsedDesc && parsedDesc.tableUid) {
                                info.tableUid = parsedDesc.tableUid;
                            }
                            if (parsedDesc) {
                                info.metaObj = parsedDesc;
                            }
                        }

                        if (!info.tableUid) {
                            var descMeta = 'table-meta:' + JSON.stringify({ tableUid: generateUid('table'), createdAt: new Date().toISOString(), source: 'table-description' });
                            info.descriptionMeta = descMeta;
                            var parsedMeta = parsePrefixedJsonTag(descMeta, 'table-meta:');
                            if (parsedMeta && parsedMeta.tableUid) {
                                info.tableUid = parsedMeta.tableUid;
                                info.metaObj = parsedMeta;
                            }
                            if (typeof table.SetTableDescription === 'function') {
                                try {
                                    table.SetTableDescription(descMeta);
                                } catch (setDescErr) {
                                    console.log('[TableService] 设置表格描述失败:', setDescErr);
                                }
                            }
                        }
                    }

                    if (!info.tag && info.descriptionMeta) {
                        info.tag = info.descriptionMeta;
                    }

                    if (!info.tableUid) {
                        info.tableUid = generateUid('table');
                    }

                    if (!info.metaObj) {
                        info.metaObj = parsePrefixedJsonTag(info.tag, 'table-meta:');
                    }

                    return info;
                }

                function computeTableIndex(doc, table) {
                    if (!doc || !table || typeof doc.GetElementsCount !== 'function' || typeof doc.GetElement !== 'function') {
                        return -1;
                    }
                    try {
                        var total = doc.GetElementsCount();
                        for (var idx = 0; idx < total; idx++) {
                            var element = doc.GetElement(idx);
                            if (element === table) {
                                return idx;
                            }
                        }
                    } catch (err) {
                        console.log('[TableService] 获取表格索引失败:', err);
                    }
                    return -1;
                }

                function findRowRange(cell) {
                    if (!cell) {
                        return { row: null, range: null };
                    }

                    var parentRowRes = safeCall(cell, 'GetParentRow');
                    var row = parentRowRes.ok ? parentRowRes.value : null;
                    var rowRange = null;

                    if (row && typeof row.GetRange === 'function') {
                        try {
                            rowRange = row.GetRange();
                        } catch (errRowRange) {
                            console.log('[TableService] 获取行Range失败:', errRowRange);
                        }
                    }

                    if (!rowRange && row && typeof row.GetCell === 'function') {
                        try {
                            var firstCell = row.GetCell(0);
                            if (firstCell && typeof firstCell.GetRange === 'function') {
                                rowRange = firstCell.GetRange();
                            }
                        } catch (errFirstCellRange) {
                            console.log('[TableService] 从首个单元格获取行Range失败:', errFirstCellRange);
                        }
                    }

                    return { row: row, range: rowRange };
                }

                function findExistingRowControl(doc, rowRange, tableInfo, rowIndex) {
                    if (!doc || typeof doc.GetAllContentControls !== 'function') {
                        return null;
                    }

                    var controls = [];
                    try {
                        controls = doc.GetAllContentControls();
                    } catch (errControls) {
                        console.log('[TableService] 获取所有内容控件失败:', errControls);
                        return null;
                    }

                    for (var i = 0; i < controls.length; i++) {
                        var control = controls[i];
                        if (!control) {
                            continue;
                        }

                        var tag = '';
                        var alias = '';
                        var controlType = null;
                        var internalId = '';

                        var tagRes = safeCall(control, 'GetTag');
                        if (tagRes.ok && typeof tagRes.value === 'string') {
                            tag = tagRes.value;
                        }

                        var aliasRes = safeCall(control, 'GetAlias');
                        if (aliasRes.ok && typeof aliasRes.value === 'string') {
                            alias = aliasRes.value;
                        }

                        var typeRes = safeCall(control, 'GetType');
                        if (typeRes.ok) {
                            controlType = typeRes.value;
                        }

                        var internalIdRes = safeCall(control, 'GetInternalId');
                        if (internalIdRes.ok && internalIdRes.value) {
                            internalId = internalIdRes.value;
                        }

                        var parsedRowMeta = parsePrefixedJsonTag(tag, 'row-meta:');
                        var matchesByTag = false;
                        if (parsedRowMeta) {
                            if (tableInfo && tableInfo.tableUid && parsedRowMeta.tableUid === tableInfo.tableUid) {
                                if (typeof parsedRowMeta.rowIndex === 'number') {
                                    matchesByTag = parsedRowMeta.rowIndex === rowIndex;
                                } else {
                                    matchesByTag = true;
                                }
                            } else if (tableInfo && tableInfo.internalId && parsedRowMeta.tableInternalId === tableInfo.internalId) {
                                matchesByTag = typeof parsedRowMeta.rowIndex === 'number' ? parsedRowMeta.rowIndex === rowIndex : true;
                            }
                        }

                        var matchesByRange = false;
                        if (!matchesByTag && rowRange && typeof control.IsRangeIn === 'function') {
                            try {
                                matchesByRange = control.IsRangeIn(rowRange);
                            } catch (errIsRange) {
                                console.log('[TableService] 判断控件是否在行范围内失败:', errIsRange);
                            }
                        }

                        if (matchesByTag || matchesByRange) {
                            return {
                                control: control,
                                tag: tag,
                                alias: alias,
                                type: controlType,
                                internalId: internalId,
                                meta: parsedRowMeta || null
                            };
                        }
                    }

                    return null;
                }

                var doc = null;
                try {
                    if (typeof Api !== 'undefined' && typeof Api.GetDocument === 'function') {
                        doc = Api.GetDocument();
                    }
                } catch (docErr) {
                    console.log('[TableService] 获取文档对象失败:', docErr);
                }

                if (!doc) {
                    return { success: false, error: 'Document object unavailable' };
                }

                var range = null;
                if (typeof doc.GetRangeBySelect === 'function') {
                    try {
                        range = doc.GetRangeBySelect();
                    } catch (rangeErr) {
                        console.log('[TableService] 获取选区失败:', rangeErr);
                    }
                }

                if (!range) {
                    return { success: false, error: 'No selection found' };
                }

                var cell = findCellFromRange(range);
                if (!cell) {
                    return { success: false, error: 'Selection is not inside a table cell' };
                }

                var tableResult = safeCall(cell, 'GetParentTable');
                var table = tableResult.ok ? tableResult.value : null;
                if (!table) {
                    return { success: false, error: 'Parent table not found' };
                }

                var rowIndexRes = safeCall(cell, 'GetRowIndex');
                var columnIndexRes = safeCall(cell, 'GetIndex');
                var rowIndex = rowIndexRes.ok ? rowIndexRes.value : null;
                var columnIndex = columnIndexRes.ok ? columnIndexRes.value : null;

                var tableMeta = ensureTableMeta(table);
                tableMeta.tableIndex = computeTableIndex(doc, table);

                var rowContext = findRowRange(cell);
                var existingRowControl = findExistingRowControl(doc, rowContext.range, tableMeta, rowIndex);

                var rowMeta = existingRowControl && existingRowControl.meta ? existingRowControl.meta : null;
                var resultData = {
                    clickType: 'table',
                    tableIndex: tableMeta.tableIndex,
                    tableUid: tableMeta.tableUid,
                    tableControlId: tableMeta.controlId,
                    tableInternalId: tableMeta.internalId,
                    tableTag: tableMeta.tag,
                    tableDescriptionTag: tableMeta.descriptionMeta,
                    tableMeta: tableMeta.metaObj || null,
                    tableControlCreated: tableMeta.createdControl,
                    rowIndex: rowIndex,
                    columnIndex: columnIndex,
                    rowControlId: existingRowControl ? existingRowControl.internalId : '',
                    rowTag: existingRowControl ? existingRowControl.tag : '',
                    rowAlias: existingRowControl ? existingRowControl.alias : '',
                    rowMeta: rowMeta,
                    hasRowControl: !!existingRowControl,
                    timestamp: new Date().toISOString()
                };

                return {
                    success: true,
                    message: 'Table cell detected',
                    data: resultData
                };

            }, { async: false, cb: (res) => resolve(res) });
        });
    }
}
