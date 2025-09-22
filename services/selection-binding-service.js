// services/selection-binding-service.js
export class SelectionBindingService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 分析当前选中的内容
    async analyzeSelection() {
        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();
            var range = doc.GetRangeBySelect();

            console.log('=== 选中内容分析 ===');
            console.log('当前选区:', range);

            if (!range) {
                console.log('未检测到选区');
                return { success: false, error: 'No selection found' };
            }

            try {
                var selectionInfo = {
                    hasSelection: true,
                    selectionType: 'unknown',
                    content: '',
                    elements: [],
                    bindable: false,
                    suggestedBindings: []
                };

                // 获取选中的文本
                if (typeof range.GetText === 'function') {
                    selectionInfo.content = range.GetText();
                    console.log('选中文本:', selectionInfo.content);
                }

                // 检测选区内容的类型
                console.log('分析选区内容类型...');

                // 方法1：检查是否选中了表格
                var tableDetected = false;
                for (var i = 0; i < 20; i++) {
                    try {
                        var element = doc.GetElement(i);
                        if (!element) break;

                        if (typeof element.GetClassType === 'function') {
                            var elementType = element.GetClassType();
                            if (elementType === 'CTable') {
                                console.log(`发现表格在位置 ${i}`);
                                // 简单判断：如果有选区且有表格，可能选中了表格
                                tableDetected = true;
                                selectionInfo.selectionType = 'table';
                                selectionInfo.bindable = true;
                                selectionInfo.suggestedBindings.push({
                                    type: 'table-data-source',
                                    description: '绑定为数据源表格',
                                    category: 'data-binding'
                                });
                                break;
                            }
                        }
                    } catch (e) {
                        console.log(`检查元素 ${i} 出错:`, e);
                    }
                }

                // 方法2：检查是否选中了文本内容
                if (!tableDetected && selectionInfo.content && selectionInfo.content.trim().length > 0) {
                    var text = selectionInfo.content.trim();
                    selectionInfo.selectionType = 'text';
                    selectionInfo.bindable = true;

                    // 分析文本内容，提供智能绑定建议
                    console.log('分析文本内容:', text);

                    // 检测数字
                    if (/^\d+(\.\d+)?$/.test(text)) {
                        selectionInfo.suggestedBindings.push({
                            type: 'number-field',
                            description: '绑定为数字字段',
                            category: 'data-field',
                            dataType: 'number',
                            value: parseFloat(text)
                        });
                    }

                    // 检测日期
                    if (/\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(text)) {
                        selectionInfo.suggestedBindings.push({
                            type: 'date-field',
                            description: '绑定为日期字段',
                            category: 'data-field',
                            dataType: 'date',
                            value: text
                        });
                    }

                    // 检测邮箱
                    if (/\w+@\w+\.\w+/.test(text)) {
                        selectionInfo.suggestedBindings.push({
                            type: 'email-field',
                            description: '绑定为邮箱字段',
                            category: 'data-field',
                            dataType: 'email',
                            value: text
                        });
                    }

                    // 检测姓名模式
                    if (/^[\u4e00-\u9fa5]{2,4}$/.test(text) || /^[A-Za-z\s]{2,20}$/.test(text)) {
                        selectionInfo.suggestedBindings.push({
                            type: 'name-field',
                            description: '绑定为姓名字段',
                            category: 'data-field',
                            dataType: 'name',
                            value: text
                        });
                    }

                    // 通用文本绑定
                    selectionInfo.suggestedBindings.push({
                        type: 'text-field',
                        description: '绑定为文本字段',
                        category: 'data-field',
                        dataType: 'text',
                        value: text
                    });

                    // 模板变量绑定
                    selectionInfo.suggestedBindings.push({
                        type: 'template-variable',
                        description: '绑定为模板变量',
                        category: 'template',
                        variable: text.replace(/\s+/g, '_').toLowerCase()
                    });
                }

                // 方法3：检查是否选中了段落
                try {
                    if (typeof range.GetParagraph === 'function') {
                        var para = range.GetParagraph();
                        if (para) {
                            console.log('选中了段落内容');
                            if (selectionInfo.selectionType === 'unknown') {
                                selectionInfo.selectionType = 'paragraph';
                                selectionInfo.bindable = true;
                                selectionInfo.suggestedBindings.push({
                                    type: 'paragraph-template',
                                    description: '绑定为段落模板',
                                    category: 'template'
                                });
                            }
                        }
                    }
                } catch (e) {
                    console.log('段落检测出错:', e);
                }

                // 如果没有建议的绑定，提供默认选项
                if (selectionInfo.suggestedBindings.length === 0) {
                    selectionInfo.suggestedBindings.push({
                        type: 'custom-binding',
                        description: '自定义数据绑定',
                        category: 'custom'
                    });
                }

                console.log('选区分析结果:', selectionInfo);

                return {
                    success: true,
                    message: 'Selection analyzed successfully',
                    data: {
                        ...selectionInfo,
                        timestamp: new Date().toLocaleString('zh-CN')
                    }
                };

            } catch (e) {
                console.log('选区分析失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 对选中内容进行数据绑定
    async bindSelection(bindingData) {
        const bindingOptions = bindingData || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            bindingType: bindingOptions.type || 'text-field',
            fieldName: bindingOptions.fieldName || 'custom_field',
            dataType: bindingOptions.dataType || 'text',
            category: bindingOptions.category || 'data-field',
            metadata: bindingOptions.metadata || {}
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();
            var range = doc.GetRangeBySelect();
            var scope = Asc.scope;

            console.log('=== 选中内容数据绑定 ===');
            console.log('绑定参数:', scope);
            console.log('当前选区:', range);

            if (!range) {
                console.log('未检测到选区');
                return { success: false, error: 'No selection to bind' };
            }

            try {
                var originalText = '';
                if (typeof range.GetText === 'function') {
                    originalText = range.GetText();
                }

                console.log('原始选中文本:', originalText);

                // 绑定文本字段 - 内联函数定义
                function bindTextFieldSelection(range, scope, originalText, doc) {
                    try {
                        console.log('执行文本字段绑定...');

                        // 创建Content Control替换选中文本
                        var sdt = doc.AddComboBoxContentControl();
                        if (!sdt) {
                            console.log('创建Content Control失败');
                            return { success: false, error: 'Failed to create Content Control' };
                        }

                        console.log('Content Control创建成功，开始设置内容...');

                        // 先设置Tag和Alias
                        if (typeof sdt.SetTag === 'function') {
                            // 优先使用 metadata.tag，如果没有则使用 fieldName
                            var tagValue = scope.metadata && scope.metadata.tag ? scope.metadata.tag : scope.fieldName;
                            sdt.SetTag(tagValue);
                            console.log('设置SDT Tag:', tagValue);
                        }

                        if (typeof sdt.SetAlias === 'function') {
                            sdt.SetAlias(scope.fieldName);
                        }

                        // 使用更简洁的方式设置内容和样式
                        console.log('设置ContentControl文本和样式...');

                        // 直接添加文本
                        if (typeof sdt.AddText === 'function') {
                            sdt.AddText(scope.fieldName);
                            console.log('已添加文本: ' + scope.fieldName);
                        }

                        // 创建文本属性并设置样式
                        if (typeof Api.CreateTextPr === 'function') {
                            var textProps = Api.CreateTextPr();
                            if (textProps) {
                                if (typeof textProps.SetColor === 'function') {
                                    textProps.SetColor(0, 100, 200); // 蓝色 RGB
                                    console.log('设置文本颜色为蓝色');
                                }
                                if (typeof textProps.SetBold === 'function') {
                                    textProps.SetBold(true);
                                    console.log('设置文本为粗体');
                                }

                                // 应用样式到整个ContentControl
                                if (typeof sdt.SetTextPr === 'function') {
                                    sdt.SetTextPr(textProps);
                                    console.log('已应用样式到ContentControl');
                                }
                            }
                        }

                        // 备选方案：如果上述方法不可用，尝试占位符
                        if (typeof sdt.AddText !== 'function' && typeof sdt.SetPlaceholderText === 'function') {
                            sdt.SetPlaceholderText('{{' + scope.fieldName + '}}');
                            console.log('使用占位符方式设置文本');
                        }

                        // 创建绑定元数据
                        var bindingMetadata = {
                            type: 'text-field-binding',
                            fieldName: scope.fieldName,
                            dataType: scope.dataType,
                            originalText: originalText,
                            boundAt: new Date().toISOString()
                        };
                        console.log('✅ 文本字段绑定完成');
                        return {
                            success: true,
                            message: 'Text field bound successfully',
                            method: 'text-field-binding',
                            binding: bindingMetadata
                        };

                    } catch (e) {
                        console.log('文本字段绑定出错:', e);
                        return { success: false, error: e.message };
                    }
                }

                // 绑定模板变量 - 内联函数定义
                function bindTemplateVariableSelection(range, scope, originalText, doc) {
                    try {
                        console.log('执行模板变量绑定...');

                        var sdt = doc.AddComboBoxContentControl();
                        if (!sdt) {
                            return { success: false, error: 'Failed to create template variable Control' };
                        }

                        var variableName = scope.metadata.variable || scope.fieldName;
                        var bindingMetadata = {
                            variableName: variableName
                        };

                        // 优先使用 metadata.tag，如果没有则使用 JSON 格式的绑定数据
                        var tagData = scope.metadata && scope.metadata.tag ? scope.metadata.tag : JSON.stringify(bindingMetadata);

                        if (typeof sdt.SetTag === 'function') {
                            sdt.SetTag(tagData);
                            console.log('设置模板变量 SDT Tag:', tagData);
                        }

                        if (typeof sdt.SetAlias === 'function') {
                            sdt.SetAlias(variableName);
                        }

                        // 直接添加文本
                        if (typeof sdt.AddText === 'function') {
                            sdt.AddText('{' + variableName + '}');
                            console.log('已添加模板变量文本: {' + variableName + '}');
                        }

                        // 创建文本属性并设置样式 - 新增
                        if (typeof Api.CreateTextPr === 'function') {
                            var textProps = Api.CreateTextPr();
                            if (textProps) {
                                if (typeof textProps.SetColor === 'function') {
                                    textProps.SetColor(150, 0, 150); // 紫色 RGB
                                    console.log('设置模板变量文本颜色为紫色');
                                }
                                if (typeof textProps.SetItalic === 'function') {
                                    textProps.SetItalic(true);
                                    console.log('设置模板变量文本为斜体');
                                }
                                if (typeof textProps.SetBold === 'function') {
                                    textProps.SetBold(true);
                                    console.log('设置模板变量文本为粗体');
                                }

                                // 应用样式到整个ContentControl
                                if (typeof sdt.SetTextPr === 'function') {
                                    sdt.SetTextPr(textProps);
                                    console.log('已应用紫色样式到模板变量ContentControl');
                                }
                            }
                        }

                        // 检查GetContent方法是否存在（备选方案）
                        if (typeof sdt.GetContent !== 'function' && typeof sdt.AddText !== 'function') {
                            console.log('sdt.GetContent不存在，使用简化方法');
                            if (typeof sdt.SetPlaceholderText === 'function') {
                                sdt.SetPlaceholderText('{' + variableName + '}');
                            }
                            return {
                                success: true,
                                message: 'Template variable bound successfully (simplified)',
                                method: 'template-variable-binding-simple',
                                binding: bindingMetadata
                            };
                        }

                        // 如果需要，使用GetContent方法进行更细致的控制
                        if (typeof sdt.GetContent === 'function' && typeof sdt.AddText !== 'function') {
                            var content = sdt.GetContent();
                            if (content) {
                                content.RemoveAllElements();
                                var para = content.GetElement(0);
                                if (para && typeof para.AddText === 'function') {
                                    para.AddText('{' + variableName + '}');
                                } else if (para) {
                                    var run = Api.CreateRun();
                                    run.AddText('{' + variableName + '}');

                                    if (typeof run.SetColor === 'function') {
                                        run.SetColor(150, 0, 150); // 紫色
                                    }
                                    if (typeof run.SetItalic === 'function') {
                                        run.SetItalic(true);
                                    }

                                    para.AddElement(run);
                                }
                            }
                        }

                        console.log('✅ 模板变量绑定完成');
                        return {
                            success: true,
                            message: 'Template variable bound successfully',
                            method: 'template-variable-binding',
                            binding: bindingMetadata
                        };

                    } catch (e) {
                        console.log('模板变量绑定出错:', e);
                        return { success: false, error: e.message };
                    }
                }

                // 绑定表格数据源 - 内联函数定义
                function bindTableDataSource(range, scope, doc) {
                    try {
                        console.log('执行表格数据源绑定...');

                        // 创建表格绑定标记
                        var marker = doc.AddComboBoxContentControl();
                        if (!marker) {
                            return { success: false, error: 'Failed to create table binding marker' };
                        }

                        var bindingMetadata = {
                            type: 'table-data-binding',
                            tableName: scope.fieldName,
                            bindingMode: 'data-source',
                            boundAt: new Date().toISOString()
                        };

                        // 优先使用 metadata.tag，如果没有则使用默认格式
                        var tagData = scope.metadata && scope.metadata.tag ? scope.metadata.tag : 'table-binding:' + JSON.stringify(bindingMetadata);

                        if (typeof marker.SetTag === 'function') {
                            marker.SetTag(tagData);
                            console.log('设置表格绑定 Tag:', tagData);
                        }

                        if (typeof marker.SetAlias === 'function') {
                            marker.SetAlias('表格数据绑定: ' + scope.fieldName);
                        }

                        // 直接添加文本
                        if (typeof marker.AddText === 'function') {
                            marker.AddText('📊 ' + scope.fieldName);
                            console.log('已添加表格数据源文本: 📊 ' + scope.fieldName);
                        }

                        // 创建文本属性并设置样式 - 新增
                        if (typeof Api.CreateTextPr === 'function') {
                            var textProps = Api.CreateTextPr();
                            if (textProps) {
                                if (typeof textProps.SetColor === 'function') {
                                    textProps.SetColor(0, 150, 0); // 绿色 RGB
                                    console.log('设置表格数据源文本颜色为绿色');
                                }
                                if (typeof textProps.SetBold === 'function') {
                                    textProps.SetBold(true);
                                    console.log('设置表格数据源文本为粗体');
                                }
                                if (typeof textProps.SetUnderline === 'function') {
                                    textProps.SetUnderline(true);
                                    console.log('设置表格数据源文本为下划线');
                                }

                                // 应用样式到整个ContentControl
                                if (typeof marker.SetTextPr === 'function') {
                                    marker.SetTextPr(textProps);
                                    console.log('已应用绿色样式到表格数据源ContentControl');
                                }
                            }
                        }

                        // 检查GetContent方法是否存在（备选方案）
                        if (typeof marker.GetContent !== 'function' && typeof marker.AddText !== 'function') {
                            console.log('marker.GetContent不存在，使用简化方法');
                            if (typeof marker.SetPlaceholderText === 'function') {
                                marker.SetPlaceholderText('📊 ' + scope.fieldName);
                            }
                            return {
                                success: true,
                                message: 'Table data source bound successfully (simplified)',
                                method: 'table-data-binding-simple',
                                binding: bindingMetadata
                            };
                        }

                        // 如果需要，使用GetContent方法进行更细致的控制
                        if (typeof marker.GetContent === 'function' && typeof marker.AddText !== 'function') {
                            var content = marker.GetContent();
                            if (content) {
                                content.RemoveAllElements();
                                var para = content.GetElement(0);
                                if (para && typeof para.AddText === 'function') {
                                    para.AddText('📊 ' + scope.fieldName);
                                } else if (para) {
                                    var run = Api.CreateRun();
                                    run.AddText('📊 ' + scope.fieldName);

                                    if (typeof run.SetColor === 'function') {
                                        run.SetColor(0, 150, 0); // 绿色
                                    }
                                    if (typeof run.SetBold === 'function') {
                                        run.SetBold(true);
                                    }

                                    para.AddElement(run);
                                }
                            }
                        }

                        console.log('✅ 表格数据源绑定完成');
                        return {
                            success: true,
                            message: 'Table data source bound successfully',
                            method: 'table-data-binding',
                            binding: bindingMetadata
                        };

                    } catch (e) {
                        console.log('表格数据源绑定出错:', e);
                        return { success: false, error: e.message };
                    }
                }

                // 绑定段落模板 - 内联函数定义
                function bindParagraphTemplate(range, scope, originalText, doc) {
                    try {
                        console.log('执行段落模板绑定...');

                        var sdt = doc.AddComboBoxContentControl();
                        if (!sdt) {
                            return { success: false, error: 'Failed to create paragraph template Control' };
                        }

                        var bindingMetadata = {
                            type: 'paragraph-template',
                            templateName: scope.fieldName,
                            originalContent: originalText,
                            boundAt: new Date().toISOString()
                        };

                        // 优先使用 metadata.tag，如果没有则使用默认格式
                        var tagData = scope.metadata && scope.metadata.tag ? scope.metadata.tag : 'paragraph-template:' + JSON.stringify(bindingMetadata);

                        if (typeof sdt.SetTag === 'function') {
                            sdt.SetTag(tagData);
                            console.log('设置段落模板 SDT Tag:', tagData);
                        }

                        if (typeof sdt.SetAlias === 'function') {
                            sdt.SetAlias('段落模板: ' + scope.fieldName);
                        }

                        // 检查GetContent方法是否存在
                        if (typeof sdt.GetContent !== 'function') {
                            console.log('sdt.GetContent不存在，使用简化方法');
                            var previewText = originalText.substring(0, 20) + (originalText.length > 20 ? '...' : '');
                            if (typeof sdt.SetPlaceholderText === 'function') {
                                sdt.SetPlaceholderText('📝 ' + scope.fieldName + ': ' + previewText);
                            }
                            return {
                                success: true,
                                message: 'Paragraph template bound successfully (simplified)',
                                method: 'paragraph-template-binding-simple',
                                binding: bindingMetadata
                            };
                        }

                        var content = sdt.GetContent();
                        if (content) {
                            content.RemoveAllElements();
                            var para = content.GetElement(0);
                            var previewText = originalText.substring(0, 20) + (originalText.length > 20 ? '...' : '');
                            if (para && typeof para.AddText === 'function') {
                                para.AddText('📝 ' + scope.fieldName + ': ' + previewText);
                            } else if (para) {
                                var run = Api.CreateRun();
                                run.AddText('📝 ' + scope.fieldName + ': ' + previewText);

                                if (typeof run.SetColor === 'function') {
                                    run.SetColor(200, 100, 0); // 橙色
                                }

                                para.AddElement(run);
                            }
                        }

                        return {
                            success: true,
                            message: 'Paragraph template bound successfully',
                            method: 'paragraph-template-binding',
                            binding: bindingMetadata
                        };

                    } catch (e) {
                        console.log('段落模板绑定出错:', e);
                        return { success: false, error: e.message };
                    }
                }

                // 自定义绑定 - 内联函数定义
                function bindCustomSelection(range, scope, originalText, doc) {
                    try {
                        console.log('执行自定义绑定...');

                        var sdt = doc.AddComboBoxContentControl();
                        if (!sdt) {
                            return { success: false, error: 'Failed to create custom binding Control' };
                        }

                        var bindingMetadata = {
                            type: 'custom-binding',
                            customType: scope.bindingType,
                            fieldName: scope.fieldName,
                            originalValue: originalText,
                            metadata: scope.metadata,
                            boundAt: new Date().toISOString()
                        };

                        // 优先使用 metadata.tag，如果没有则使用默认格式
                        var tagData = scope.metadata && scope.metadata.tag ? scope.metadata.tag : 'custom-binding:' + JSON.stringify(bindingMetadata);

                        if (typeof sdt.SetTag === 'function') {
                            sdt.SetTag(tagData);
                            console.log('设置自定义绑定 SDT Tag:', tagData);
                        }

                        if (typeof sdt.SetAlias === 'function') {
                            sdt.SetAlias('自定义绑定: ' + scope.fieldName);
                        }

                        // 检查GetContent方法是否存在
                        if (typeof sdt.GetContent !== 'function') {
                            console.log('sdt.GetContent不存在，使用简化方法');
                            if (typeof sdt.SetPlaceholderText === 'function') {
                                sdt.SetPlaceholderText('🔗 ' + scope.fieldName);
                            }
                            return {
                                success: true,
                                message: 'Custom binding created successfully (simplified)',
                                method: 'custom-binding-simple',
                                binding: bindingMetadata
                            };
                        }

                        var content = sdt.GetContent();
                        if (content) {
                            content.RemoveAllElements();
                            var para = content.GetElement(0);
                            if (para && typeof para.AddText === 'function') {
                                para.AddText('🔗 ' + scope.fieldName);
                            } else if (para) {
                                var run = Api.CreateRun();
                                run.AddText('🔗 ' + scope.fieldName);

                                if (typeof run.SetColor === 'function') {
                                    run.SetColor(100, 100, 100); // 灰色
                                }

                                para.AddElement(run);
                            }
                        }

                        return {
                            success: true,
                            message: 'Custom binding created successfully',
                            method: 'custom-binding',
                            binding: bindingMetadata
                        };

                    } catch (e) {
                        console.log('自定义绑定出错:', e);
                        return { success: false, error: e.message };
                    }
                }

                // 根据绑定类型执行不同的绑定策略
                var bindingResult = null;

                switch (scope.bindingType) {
                    case 'text-field':
                    case 'name-field':
                    case 'email-field':
                    case 'number-field':
                    case 'date-field':
                        // 文本字段绑定：将选中文本替换为Content Control
                        bindingResult = bindTextFieldSelection(range, scope, originalText, doc);
                        break;

                    case 'template-variable':
                        // 模板变量绑定：创建带变量标识的Content Control
                        bindingResult = bindTemplateVariableSelection(range, scope, originalText, doc);
                        break;

                    case 'table-data-source':
                        // 表格数据源绑定：为表格添加数据绑定标记
                        bindingResult = bindTableDataSource(range, scope, doc);
                        break;

                    case 'paragraph-template':
                        // 段落模板绑定：将段落设置为模板区块
                        bindingResult = bindParagraphTemplate(range, scope, originalText, doc);
                        break;

                    default:
                        // 自定义绑定
                        bindingResult = bindCustomSelection(range, scope, originalText, doc);
                        break;
                }

                if (bindingResult && bindingResult.success) {
                    console.log('✅ 数据绑定成功:', bindingResult);
                    return bindingResult;
                } else {
                    console.log('❌ 数据绑定失败:', bindingResult);
                    return bindingResult || { success: false, error: 'Binding failed' };
                }

            } catch (e) {
                console.log('数据绑定出错:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 处理绑定内容控件的点击事件
    async handleBindingClick() {
        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();
            var range = doc.GetRangeBySelect();

            console.log('=== 检测绑定控件点击 ===');
            console.log('当前选区:', range);

            if (!range) {
                console.log('未检测到选区');
                return { success: false, error: 'No selection found' };
            }

            try {
                // 获取所有Content Control
                var allControls = doc.GetAllContentControls();
                console.log('文档中的Content Controls数量:', allControls.length);

                // 查找当前选区内或选中的Content Control
                for (var i = 0; i < allControls.length; i++) {
                    var control = allControls[i];

                    // 检查是否命中当前控件
                    var isHit = false;
                    try {
                        if (typeof range.IsInContentControl === 'function' && range.IsInContentControl(control)) {
                            isHit = true;
                        } else if (typeof control.IsRangeIn === 'function' && control.IsRangeIn(range)) {
                            isHit = true;
                        } else if (typeof control.IsSelected === 'function' && control.IsSelected()) {
                            isHit = true;
                        }
                    } catch (e) {
                        console.log('检查控件命中出错:', e);
                    }

                    if (isHit) {
                        console.log('找到命中的Content Control:', i);

                        // 获取控件信息
                        var tag = '';
                        var alias = '';
                        var content = '';

                        try {
                            if (typeof control.GetTag === 'function') {
                                tag = control.GetTag();
                            }
                            if (typeof control.GetAlias === 'function') {
                                alias = control.GetAlias();
                            }

                            // 尝试获取控件内容
                            if (typeof control.GetContent === 'function') {
                                var controlContent = control.GetContent();
                                if (controlContent && typeof controlContent.GetText === 'function') {
                                    content = controlContent.GetText();
                                }
                            }
                        } catch (e) {
                            console.log('获取控件信息出错:', e);
                        }

                        console.log('控件Tag:', tag);
                        console.log('控件Alias:', alias);
                        console.log('控件内容:', content);

                        // 解析绑定数据
                        var bindingData = null;
                        var bindingType = 'unknown';

                        if (tag) {
                            try {
                                if (tag.startsWith('binding-data:')) {
                                    bindingType = 'data-binding';
                                    var jsonData = tag.substring('binding-data:'.length);
                                    bindingData = JSON.parse(jsonData);
                                } else if (tag.startsWith('template-var:')) {
                                    bindingType = 'template-variable';
                                    var jsonData = tag.substring('template-var:'.length);
                                    bindingData = JSON.parse(jsonData);
                                } else if (tag.startsWith('table-binding:')) {
                                    bindingType = 'table-data-binding';
                                    var jsonData = tag.substring('table-binding:'.length);
                                    bindingData = JSON.parse(jsonData);
                                } else if (tag.startsWith('paragraph-template:')) {
                                    bindingType = 'paragraph-template';
                                    var jsonData = tag.substring('paragraph-template:'.length);
                                    bindingData = JSON.parse(jsonData);
                                } else if (tag.startsWith('custom-binding:')) {
                                    bindingType = 'custom-binding';
                                    var jsonData = tag.substring('custom-binding:'.length);
                                    bindingData = JSON.parse(jsonData);
                                }
                            } catch (e) {
                                console.log('解析绑定数据出错:', e);
                            }
                        }

                        var clickResult = {
                            success: true,
                            message: 'Binding control clicked',
                            data: {
                                clickType: 'binding-control',
                                controlIndex: i,
                                controlTag: tag,
                                controlAlias: alias,
                                controlContent: content,
                                bindingType: bindingType,
                                bindingData: bindingData,
                                timestamp: new Date().toLocaleString('zh-CN')
                            }
                        };

                        console.log('绑定控件点击结果:', clickResult);
                        return clickResult;
                    }
                }

                // 没有找到绑定控件
                console.log('当前位置没有找到绑定控件');
                return {
                    success: false,
                    error: 'No binding control found at current position',
                    data: {
                        clickType: 'non-binding',
                        timestamp: new Date().toLocaleString('zh-CN')
                    }
                };

            } catch (e) {
                console.log('检测绑定控件点击失败:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }

    // 更新绑定数据
    async updateBinding(updateData) {
        const options = updateData || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            rid: options.rid,
            tag: options.tag || options.fieldName,
            fieldName: options.fieldName,
            displayName: options.displayName,
            metadata: options.metadata || {}
        };

        const updateResult = await this.editor.runInDoc(function () {
            function parsePrefixedJsonTag(tag, prefix) {
                if (typeof tag !== 'string' || tag.indexOf(prefix) !== 0) {
                    return null;
                }
                try {
                    return JSON.parse(tag.slice(prefix.length));
                } catch (err) {
                    console.log('[SelectionBinding] 解析标签失败:', tag, err);
                    return null;
                }
            }

            var doc = Api.GetDocument();
            var scope = Asc.scope;

            console.log('=== 更新绑定数据 ===');
            console.log('更新参数:', scope);

            try {
                // 获取所有Content Control
                var allControls = doc.GetAllContentControls();
                console.log('文档中的Content Controls数量:', allControls.length);

                // 查找匹配rid的控件
                for (var i = 0; i < allControls.length; i++) {
                    var control = allControls[i];

                    try {
                        var tag = '';
                        if (typeof control.GetTag === 'function') {
                            tag = control.GetTag();
                        }

                        var alias = '';
                        if (typeof control.GetAlias === 'function') {
                            alias = control.GetAlias();
                        }

                        // 调试：显示实际的控件信息
                        console.log("控件 " + i + ": Tag=\"" + tag + "\", Alias=\"" + alias + "\"");
                        console.log("查找目标: rid=\"" + scope.rid + "\", tag=\"" + scope.tag + "\"");

                        // 检查是否匹配rid或现有tag
                        var isMatch = false;

                        // 首先尝试直接获取控件的internalId
                        var controlInternalId = '';
                        try {
                            if (typeof control.GetInternalId === 'function') {
                                controlInternalId = control.GetInternalId();
                            }
                        } catch (e) {
                            console.log('获取控件internalId失败:', e);
                        }

                        console.log("控件internalId:", controlInternalId, "目标rid:", scope.rid);

                        // 优先使用internalId匹配（最可靠）
                        if (scope.rid && controlInternalId && controlInternalId === scope.rid) {
                            isMatch = true;
                            console.log("匹配方式: internalId精确匹配");
                        }
                        // 备选：检查tag和alias中是否包含rid（用于向后兼容）
                        else if (scope.rid && (tag.includes(scope.rid) || alias.includes(scope.rid))) {
                            isMatch = true;
                            console.log("匹配方式: tag/alias包含匹配");
                        }
                        // 尝试解析JSON格式的tag中的id字段
                        else if (scope.rid) {
                            try {
                                if (tag.startsWith("{") || tag.includes("id")) {
                                    var tagObj = JSON.parse(tag);
                                    if (tagObj && (tagObj.internalId === scope.rid || tagObj.id === scope.rid)) {
                                        isMatch = true;
                                        console.log("匹配方式: JSON tag解析匹配");
                                    }
                                }
                            } catch (jsonErr) {
                                // 忽略JSON解析错误
                                console.log("JSON解析失败:", jsonErr);
                            }
                        }

                        if (isMatch) {
                            console.log('找到匹配的控件:', i, 'Tag:', tag, 'Alias:', alias);

                            // 更新Tag
                            if (typeof control.SetTag === 'function' && scope.tag) {
                                control.SetTag(scope.tag);
                                console.log('更新Tag为:', scope.tag);
                            }

                            // 更新Alias
                            if (typeof control.SetAlias === 'function' && scope.displayName) {
                                control.SetAlias(scope.displayName);
                                console.log('更新Alias为:', scope.displayName);
                            }

                            return {
                                success: true,
                                message: 'Binding updated successfully',
                                data: {
                                    controlIndex: i,
                                    oldTag: tag,
                                    newTag: scope.tag,
                                    displayName: scope.displayName,
                                    updatedAt: new Date().toISOString()
                                }
                            };
                        }
                    } catch (e) {
                        console.log('检查控件出错:', e);
                    }
                }

                // 没有找到匹配的控件
                console.log('未找到匹配的控件');
                return {
                    success: false,
                    error: 'No matching control found for rid: ' + scope.rid
                };

            } catch (e) {
                console.log('更新绑定数据出错:', e);
                return { success: false, error: e.message };
            }
        });
    }

    // 为当前光标所在行绑定隐藏数据
    async bindRowData(bindingPayload) {
        const payload = bindingPayload || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = Object.assign({}, window.Asc.scope || {}, {
            rowBindingPayload: {
                data: payload.data || {},
                extraMeta: payload.extraMeta || {},
                alias: payload.alias || '',
                lock: (payload.lock && typeof payload.lock === 'object') ? payload.lock : null,
                requestedRowUid: payload.rowUid || '',
                preserveExisting: !!payload.preserveExisting
            }
        });

        const initialResult = await new Promise((resolve) => {
            this.editor.runInDoc(function () {
                function safeCall(target, methodName, args) {
                    if (!target || typeof target[methodName] !== 'function') {
                        return { ok: false };
                    }
                    try {
                        return { ok: true, value: target[methodName].apply(target, args || []) };
                    } catch (err) {
                        console.log('[SelectionBinding] 调用', methodName, '失败:', err);
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
                        console.log('[SelectionBinding] 解析标签失败:', tag, e);
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
                    var paraRes = safeCall(range, 'GetParagraph');
                    if (paraRes.ok) {
                        paragraph = paraRes.value;
                    }

                    if (!paragraph && typeof range.GetElementsCount === 'function' && typeof range.GetElement === 'function') {
                        try {
                            var elementsCount = range.GetElementsCount();
                            for (var i = 0; i < elementsCount; i++) {
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
                                        console.log('[SelectionBinding] 获取父段落失败:', errPara);
                                    }
                                }
                            }
                        } catch (errElements) {
                            console.log('[SelectionBinding] 遍历选区元素失败:', errElements);
                        }
                    }

                    if (paragraph && typeof paragraph.GetParentTableCell === 'function') {
                        try {
                            var cellFromParagraph = paragraph.GetParentTableCell();
                            if (cellFromParagraph) {
                                return cellFromParagraph;
                            }
                        } catch (errCell) {
                            console.log('[SelectionBinding] 从段落取单元格失败:', errCell);
                        }
                    }

                    if (typeof range.GetElement === 'function') {
                        try {
                            var firstElement = range.GetElement(0);
                            if (firstElement && typeof firstElement.GetParentTableCell === 'function') {
                                return firstElement.GetParentTableCell();
                            }
                        } catch (errFirstElement) {
                            console.log('[SelectionBinding] 从首元素获取单元格失败:', errFirstElement);
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
                        descriptionMeta: ''
                    };

                    if (!table) {
                        return info;
                    }

                    var parentControl = safeCall(table, 'GetParentContentControl');
                    if (parentControl.ok && parentControl.value) {
                        info.control = parentControl.value;
                    }

                    if (!info.control) {
                        var insertRes = safeCall(table, 'InsertInContentControl', [1]);
                        if (insertRes.ok && insertRes.value) {
                            info.control = insertRes.value;
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
                                    console.log('[SelectionBinding] 设置表格控件标签失败:', setTagErr);
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
                            }
                            if (parsedMeta) {
                                info.metaObj = parsedMeta;
                            }
                            if (typeof table.SetTableDescription === 'function') {
                                try {
                                    table.SetTableDescription(descMeta);
                                } catch (setDescErr) {
                                    console.log('[SelectionBinding] 设置表格描述失败:', setDescErr);
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
                        console.log('[SelectionBinding] 获取表格索引失败:', err);
                    }
                    return -1;
                }

                function findRowRange(cell) {
                    if (!cell) {
                        return { row: null, range: null };
                    }

                    var rowRes = safeCall(cell, 'GetParentRow');
                    var row = rowRes.ok ? rowRes.value : null;
                    var rowRange = null;

                    if (row && typeof row.GetRange === 'function') {
                        try {
                            rowRange = row.GetRange();
                        } catch (rowRangeErr) {
                            console.log('[SelectionBinding] 获取行Range失败:', rowRangeErr);
                        }
                    }

                    if (!rowRange && row && typeof row.GetCell === 'function') {
                        try {
                            var firstCell = row.GetCell(0);
                            if (firstCell && typeof firstCell.GetRange === 'function') {
                                rowRange = firstCell.GetRange();
                            }
                        } catch (firstCellErr) {
                            console.log('[SelectionBinding] 从首个单元格取Range失败:', firstCellErr);
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
                    } catch (controlsErr) {
                        console.log('[SelectionBinding] 获取内容控件失败:', controlsErr);
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
                            } catch (isRangeErr) {
                                console.log('[SelectionBinding] 判断控件是否在行范围失败:', isRangeErr);
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
                    console.log('[SelectionBinding] 获取文档对象失败:', docErr);
                }

                if (!doc) {
                    return { success: false, error: 'Document object unavailable' };
                }

                var payload = (Asc.scope && Asc.scope.rowBindingPayload) ? Asc.scope.rowBindingPayload : {};

                var range = null;
                if (typeof doc.GetRangeBySelect === 'function') {
                    try {
                        range = doc.GetRangeBySelect();
                    } catch (rangeErr) {
                        console.log('[SelectionBinding] 获取选区失败:', rangeErr);
                    }
                }

                if (!range) {
                    return { success: false, error: 'No selection found' };
                }

                var cell = findCellFromRange(range);
                if (!cell) {
                    return { success: false, error: 'Selection is not inside a table cell' };
                }

                var tableRes = safeCall(cell, 'GetParentTable');
                var table = tableRes.ok ? tableRes.value : null;
                if (!table) {
                    return { success: false, error: 'Parent table not found' };
                }

                var tableMeta = ensureTableMeta(table);
                tableMeta.tableIndex = computeTableIndex(doc, table);

                var rowIndexRes = safeCall(cell, 'GetRowIndex');
                var columnIndexRes = safeCall(cell, 'GetIndex');
                var rowIndex = rowIndexRes.ok ? rowIndexRes.value : null;
                var columnIndex = columnIndexRes.ok ? columnIndexRes.value : null;

                var rowContext = findRowRange(cell);
                var existingRowControl = findExistingRowControl(doc, rowContext.range, tableMeta, rowIndex);

                var nowIso = new Date().toISOString();
                var rowMeta = existingRowControl && existingRowControl.meta ? existingRowControl.meta : null;

                if (!rowMeta) {
                    rowMeta = {
                        tableUid: tableMeta.tableUid,
                        tableInternalId: tableMeta.internalId || '',
                        rowIndex: rowIndex,
                        rowUid: payload.requestedRowUid || generateUid('row'),
                        createdAt: nowIso
                    };
                }

                if (!rowMeta.rowUid) {
                    rowMeta.rowUid = payload.requestedRowUid || generateUid('row');
                }

                if (!rowMeta.createdAt) {
                    rowMeta.createdAt = nowIso;
                }

                rowMeta.updatedAt = nowIso;
                rowMeta.data = payload.data || {};
                if (payload.extraMeta && typeof payload.extraMeta === 'object') {
                    rowMeta.extra = payload.extraMeta;
                }

                if (!rowMeta.tableUid) {
                    rowMeta.tableUid = tableMeta.tableUid;
                }
                if (!rowMeta.tableInternalId && tableMeta.internalId) {
                    rowMeta.tableInternalId = tableMeta.internalId;
                }

                var rowTag = 'row-meta:' + JSON.stringify(rowMeta);
                var aliasText = payload.alias || ('表格行绑定 - 第' + (typeof rowIndex === 'number' ? (rowIndex + 1) : '?') + '行');

                if (existingRowControl && existingRowControl.control) {
                    if (!payload.preserveExisting) {
                        if (typeof existingRowControl.control.SetTag === 'function') {
                            try {
                                existingRowControl.control.SetTag(rowTag);
                            } catch (setTagErr) {
                                console.log('[SelectionBinding] 更新行标签失败:', setTagErr);
                            }
                        }
                    }

                    if (typeof existingRowControl.control.SetAlias === 'function') {
                        try {
                            existingRowControl.control.SetAlias(aliasText);
                        } catch (setAliasErr) {
                            console.log('[SelectionBinding] 更新行别名失败:', setAliasErr);
                        }
                    }

                    return {
                        success: true,
                        message: 'Row binding updated',
                        data: {
                            needCreation: false,
                            action: 'updated',
                            tableUid: tableMeta.tableUid,
                            tableInternalId: tableMeta.internalId,
                            tableControlId: tableMeta.controlId,
                            tableTag: tableMeta.tag,
                            tableDescriptionTag: tableMeta.descriptionMeta,
                            tableMeta: tableMeta.metaObj || null,
                            tableIndex: tableMeta.tableIndex,
                            rowIndex: rowIndex,
                            columnIndex: columnIndex,
                            rowControlId: existingRowControl.internalId,
                            rowTag: rowTag,
                            rowAlias: aliasText,
                            rowMeta: rowMeta,
                            timestamp: nowIso
                        }
                    };
                }

                var selectionPrepared = false;
                if (rowContext.range) {
                    if (typeof rowContext.range.Select === 'function') {
                        try {
                            rowContext.range.Select();
                            selectionPrepared = true;
                        } catch (selectErr) {
                            console.log('[SelectionBinding] 选中行范围失败:', selectErr);
                        }
                    }

                    if (!selectionPrepared && typeof doc.SetSelection === 'function') {
                        try {
                            doc.SetSelection(rowContext.range);
                            selectionPrepared = true;
                        } catch (setSelectionErr) {
                            console.log('[SelectionBinding] 设置选区失败:', setSelectionErr);
                        }
                    }
                }

                return {
                    success: true,
                    message: 'Row binding requires creation',
                        data: {
                            needCreation: true,
                            action: 'create',
                            tableUid: tableMeta.tableUid,
                            tableInternalId: tableMeta.internalId,
                            tableControlId: tableMeta.controlId,
                            tableTag: tableMeta.tag,
                            tableDescriptionTag: tableMeta.descriptionMeta,
                            tableMeta: tableMeta.metaObj || null,
                            tableIndex: tableMeta.tableIndex,
                            rowIndex: rowIndex,
                            columnIndex: columnIndex,
                            tag: rowTag,
                            alias: aliasText,
                            rowMeta: rowMeta,
                        lock: payload.lock,
                        selectionPrepared: selectionPrepared,
                        timestamp: nowIso
                    }
                };

            }, { async: false, cb: (res) => resolve(res) });
        });

        if (!initialResult) {
            return { success: false, error: 'Row binding command failed to execute' };
        }

        if (!initialResult.success) {
            return initialResult;
        }

        const docData = initialResult.data || {};

        if (!docData.needCreation) {
            const persistResult = await this.persistRowMetaToTable({
                tableUid: docData.tableUid,
                tableInternalId: docData.tableInternalId,
                tableControlId: docData.tableControlId,
                tableTag: docData.tableTag,
                tableDescriptionTag: docData.tableDescriptionTag,
                tableIndex: docData.tableIndex,
                existingMeta: docData.tableMeta || null,
                rowMeta: docData.rowMeta || null,
                rowIndex: docData.rowIndex,
                rowControlId: docData.rowControlId,
                action: docData.action || 'updated'
            });

            if (persistResult && persistResult.success && persistResult.data) {
                initialResult.data.tableTag = persistResult.data.tableTag || initialResult.data.tableTag;
                initialResult.data.tableDescriptionTag = persistResult.data.tableDescriptionTag || initialResult.data.tableDescriptionTag;
                initialResult.data.tableMeta = persistResult.data.tableMeta || initialResult.data.tableMeta;
                if (persistResult.data.persistenceLog) {
                    initialResult.data.persistenceLog = persistResult.data.persistenceLog;
                }
            } else if (persistResult && !persistResult.success) {
                initialResult.data.persistenceWarning = persistResult.error || 'Table metadata persistence failed';
            }

            return initialResult;
        }

        if (!window.Asc || !window.Asc.plugin || typeof window.Asc.plugin.executeMethod !== 'function') {
            return {
                success: false,
                error: 'AddContentControl executeMethod unavailable',
                data: docData
            };
        }

        const addOptions = {};
        if (docData.tag) {
            addOptions.Tag = docData.tag;
        }
        if (docData.lock && typeof docData.lock === 'object') {
            addOptions.Lock = docData.lock;
        }

        const creationResult = await new Promise((resolve) => {
            try {
                window.Asc.plugin.executeMethod('AddContentControl', [3, addOptions], function () {
                    try {
                        window.Asc.plugin.executeMethod('GetCurrentContentControlPr', [], function (ccPr) {
                            resolve({ success: true, ccPr });
                        });
                    } catch (getPrErr) {
                        console.log('[SelectionBinding] 获取新建内容控件属性失败:', getPrErr);
                        resolve({ success: false, error: getPrErr && getPrErr.message ? getPrErr.message : 'GetCurrentContentControlPr failed' });
                    }
                });
            } catch (err) {
                resolve({ success: false, error: err && err.message ? err.message : String(err) });
            }
        });

        if (!creationResult.success) {
            return {
                success: false,
                error: creationResult.error || 'Row content control creation failed',
                data: docData
            };
        }

        const finalizeScope = {
            rowBindingFinalize: {
                internalId: (creationResult.ccPr && creationResult.ccPr.InternalId) ? creationResult.ccPr.InternalId : '',
                tag: docData.tag,
                alias: docData.alias,
                tableUid: docData.tableUid,
                tableInternalId: docData.tableInternalId,
                tableControlId: docData.tableControlId,
                tableTag: docData.tableTag,
                tableDescriptionTag: docData.tableDescriptionTag,
                rowIndex: docData.rowIndex,
                columnIndex: docData.columnIndex,
                rowMeta: docData.rowMeta
            }
        };

        const finalizeResult = await new Promise((resolve) => {
            this.editor.runInDoc(function () {
                function safeCall(target, methodName, args) {
                    if (!target || typeof target[methodName] !== 'function') {
                        return { ok: false };
                    }
                    try {
                        return { ok: true, value: target[methodName].apply(target, args || []) };
                    } catch (err) {
                        console.log('[SelectionBinding] finalize 调用', methodName, '失败:', err);
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
                        console.log('[SelectionBinding] finalize 解析标签失败:', tag, e);
                        return null;
                    }
                }

                var doc = null;
                try {
                    if (typeof Api !== 'undefined' && typeof Api.GetDocument === 'function') {
                        doc = Api.GetDocument();
                    }
                } catch (docErr) {
                    console.log('[SelectionBinding] finalize 获取文档失败:', docErr);
                }

                if (!doc) {
                    return { success: false, error: 'Document object unavailable' };
                }

                var payload = (Asc.scope && Asc.scope.rowBindingFinalize) ? Asc.scope.rowBindingFinalize : null;
                if (!payload || !payload.internalId) {
                    return { success: false, error: 'Row control internalId missing' };
                }

                var controls = [];
                if (typeof doc.GetAllContentControls === 'function') {
                    try {
                        controls = doc.GetAllContentControls();
                    } catch (controlsErr) {
                        console.log('[SelectionBinding] finalize 获取控件失败:', controlsErr);
                    }
                }

                for (var i = 0; i < controls.length; i++) {
                    var control = controls[i];
                    if (!control) {
                        continue;
                    }

                    var internalIdRes = safeCall(control, 'GetInternalId');
                    if (!internalIdRes.ok || internalIdRes.value !== payload.internalId) {
                        continue;
                    }

                    if (payload.tag && typeof control.SetTag === 'function') {
                        try {
                            control.SetTag(payload.tag);
                        } catch (setTagErr) {
                            console.log('[SelectionBinding] finalize 设置Tag失败:', setTagErr);
                        }
                    }

                    if (payload.alias && typeof control.SetAlias === 'function') {
                        try {
                            control.SetAlias(payload.alias);
                        } catch (setAliasErr) {
                            console.log('[SelectionBinding] finalize 设置Alias失败:', setAliasErr);
                        }
                    }

                    var finalTagRes = safeCall(control, 'GetTag');
                    var finalAliasRes = safeCall(control, 'GetAlias');

                    var finalTag = finalTagRes.ok && typeof finalTagRes.value === 'string' ? finalTagRes.value : (payload.tag || '');
                    var finalAlias = finalAliasRes.ok && typeof finalAliasRes.value === 'string' ? finalAliasRes.value : (payload.alias || '');
                    var parsedMeta = parsePrefixedJsonTag(finalTag, 'row-meta:');

                    return {
                        success: true,
                        message: 'Row binding created successfully',
                        data: {
                            action: 'created',
                            tableUid: payload.tableUid,
                            tableInternalId: payload.tableInternalId,
                            tableControlId: payload.tableControlId,
                            tableTag: payload.tableTag,
                            tableDescriptionTag: payload.tableDescriptionTag,
                            rowIndex: payload.rowIndex,
                            columnIndex: payload.columnIndex,
                            rowControlId: payload.internalId,
                            rowTag: finalTag,
                            rowAlias: finalAlias,
                            rowMeta: parsedMeta || payload.rowMeta || null,
                            timestamp: new Date().toISOString()
                        }
                    };
                }

                return { success: false, error: 'New row content control not found after creation' };

            }, { async: false, scope: finalizeScope, cb: (res) => resolve(res) });
        });

        if (!finalizeResult) {
            return { success: false, error: 'Row binding finalize failed' };
        }

        if (!finalizeResult.success) {
            return finalizeResult;
        }

        const finalData = Object.assign({}, docData, finalizeResult.data || {});
        finalData.action = 'created';
        finalData.needCreation = false;

        const finalizePersistResult = await this.persistRowMetaToTable({
            tableUid: finalData.tableUid,
            tableInternalId: finalData.tableInternalId,
            tableControlId: finalData.tableControlId,
            tableTag: finalData.tableTag,
            tableDescriptionTag: finalData.tableDescriptionTag,
            tableIndex: finalData.tableIndex,
            existingMeta: finalData.tableMeta || null,
            rowMeta: finalData.rowMeta || null,
            rowIndex: finalData.rowIndex,
            rowControlId: finalData.rowControlId,
            action: finalData.action
        });

        if (finalizePersistResult && finalizePersistResult.success && finalizePersistResult.data) {
            finalData.tableTag = finalizePersistResult.data.tableTag || finalData.tableTag;
            finalData.tableDescriptionTag = finalizePersistResult.data.tableDescriptionTag || finalData.tableDescriptionTag;
            finalData.tableMeta = finalizePersistResult.data.tableMeta || finalData.tableMeta;
            if (finalizePersistResult.data.persistenceLog) {
                finalData.persistenceLog = finalizePersistResult.data.persistenceLog;
            }
        } else if (finalizePersistResult && !finalizePersistResult.success) {
            finalData.persistenceWarning = finalizePersistResult.error || 'Table metadata persistence failed';
        }

        return {
            success: true,
            message: 'Row binding created successfully',
            data: finalData
        };
    }

    async persistRowMetaToTable(context) {
        const payload = Object.assign({
            tableUid: '',
            tableInternalId: '',
            tableControlId: '',
            tableTag: '',
            tableDescriptionTag: '',
            tableIndex: -1,
            existingMeta: null,
            rowMeta: null,
            rowIndex: null,
            rowControlId: '',
            action: ''
        }, context || {});

        window.Asc = window.Asc || {};
        window.Asc.scope = Object.assign({}, window.Asc.scope || {}, {
            tableRowPersistence: payload
        });

        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                function safeCall(target, methodName, args) {
                    if (!target || typeof target[methodName] !== 'function') {
                        return { ok: false };
                    }
                    try {
                        return { ok: true, value: target[methodName].apply(target, args || []) };
                    } catch (err) {
                        console.log('[SelectionBinding] 表格元数据持久化调用', methodName, '失败:', err);
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
                        console.log('[SelectionBinding] 解析标签失败:', tag, e);
                        return null;
                    }
                }

                function generateUid(prefix) {
                    return (prefix || 'uid') + '-' + (new Date().getTime().toString(36)) + '-' + Math.floor(Math.random() * 1e8).toString(36);
                }

                function cloneMeta(meta) {
                    if (!meta || typeof meta !== 'object') {
                        return {};
                    }
                    try {
                        return JSON.parse(JSON.stringify(meta));
                    } catch (err) {
                        console.log('[SelectionBinding] 克隆元数据失败:', err);
                        var copy = {};
                        for (var key in meta) {
                            if (Object.prototype.hasOwnProperty.call(meta, key)) {
                                copy[key] = meta[key];
                            }
                        }
                        return copy;
                    }
                }

                var doc = null;
                try {
                    if (typeof Api !== 'undefined' && typeof Api.GetDocument === 'function') {
                        doc = Api.GetDocument();
                    }
                } catch (docErr) {
                    console.log('[SelectionBinding] 表格元数据持久化获取文档失败:', docErr);
                }

                if (!doc) {
                    return { success: false, error: 'Document object unavailable' };
                }

                var persistencePayload = (Asc.scope && Asc.scope.tableRowPersistence) ? Asc.scope.tableRowPersistence : null;
                if (!persistencePayload) {
                    return { success: false, error: 'Persistence payload missing' };
                }

                var controls = [];
                if (typeof doc.GetAllContentControls === 'function') {
                    try {
                        controls = doc.GetAllContentControls();
                    } catch (controlsErr) {
                        console.log('[SelectionBinding] 表格元数据持久化获取控件失败:', controlsErr);
                    }
                }

                var targetControl = null;
                for (var i = 0; i < controls.length; i++) {
                    var control = controls[i];
                    if (!control) {
                        continue;
                    }

                    var idRes = safeCall(control, 'GetId');
                    var internalIdRes = safeCall(control, 'GetInternalId');
                    var tagRes = safeCall(control, 'GetTag');
                    var tag = tagRes.ok && typeof tagRes.value === 'string' ? tagRes.value : '';
                    var parsedTag = parsePrefixedJsonTag(tag, 'table-meta:');

                    var matches = false;
                    if (persistencePayload.tableInternalId && internalIdRes.ok && internalIdRes.value === persistencePayload.tableInternalId) {
                        matches = true;
                    } else if (!matches && persistencePayload.tableControlId && idRes.ok && idRes.value === persistencePayload.tableControlId) {
                        matches = true;
                    } else if (!matches && persistencePayload.tableUid && parsedTag && parsedTag.tableUid === persistencePayload.tableUid) {
                        matches = true;
                    }

                    if (matches) {
                        targetControl = control;
                        break;
                    }
                }

                var existingMeta = null;
                if (persistencePayload.existingMeta && typeof persistencePayload.existingMeta === 'object') {
                    existingMeta = cloneMeta(persistencePayload.existingMeta);
                }

                if (!existingMeta && typeof persistencePayload.tableTag === 'string') {
                    existingMeta = parsePrefixedJsonTag(persistencePayload.tableTag, 'table-meta:');
                }

                if (!existingMeta && typeof persistencePayload.tableDescriptionTag === 'string') {
                    existingMeta = parsePrefixedJsonTag(persistencePayload.tableDescriptionTag, 'table-meta:');
                }

                if (!existingMeta) {
                    existingMeta = {};
                }

                if (!existingMeta.tableUid && persistencePayload.tableUid) {
                    existingMeta.tableUid = persistencePayload.tableUid;
                }

                if (!existingMeta.createdAt) {
                    existingMeta.createdAt = new Date().toISOString();
                }

                var nowIso = new Date().toISOString();
                existingMeta.lastUpdatedAt = nowIso;
                if (persistencePayload.action) {
                    existingMeta.lastAction = persistencePayload.action;
                }

                existingMeta.rowBindings = existingMeta.rowBindings && typeof existingMeta.rowBindings === 'object' ? existingMeta.rowBindings : {};
                existingMeta.rowIndexMap = existingMeta.rowIndexMap && typeof existingMeta.rowIndexMap === 'object' ? existingMeta.rowIndexMap : {};
                existingMeta.rowControlMap = existingMeta.rowControlMap && typeof existingMeta.rowControlMap === 'object' ? existingMeta.rowControlMap : {};

                var storedRowUid = '';
                if (persistencePayload.rowMeta && typeof persistencePayload.rowMeta === 'object') {
                    var rowMeta = cloneMeta(persistencePayload.rowMeta);
                    if (typeof persistencePayload.rowIndex === 'number' && typeof rowMeta.rowIndex !== 'number') {
                        rowMeta.rowIndex = persistencePayload.rowIndex;
                    }
                    if (!rowMeta.rowUid) {
                        rowMeta.rowUid = generateUid('row');
                    }
                    if (!rowMeta.createdAt) {
                        rowMeta.createdAt = nowIso;
                    }
                    rowMeta.updatedAt = nowIso;
                    storedRowUid = rowMeta.rowUid;

                    existingMeta.rowBindings[rowMeta.rowUid] = rowMeta;

                    if (typeof rowMeta.rowIndex === 'number') {
                        var idxKey = String(rowMeta.rowIndex);
                        existingMeta.rowIndexMap[idxKey] = {
                            rowUid: rowMeta.rowUid,
                            rowControlId: persistencePayload.rowControlId || rowMeta.rowControlId || '',
                            updatedAt: nowIso
                        };
                    }

                    if (persistencePayload.rowControlId) {
                        existingMeta.rowControlMap[persistencePayload.rowControlId] = {
                            rowUid: rowMeta.rowUid,
                            rowIndex: rowMeta.rowIndex,
                            updatedAt: nowIso
                        };
                    }

                    existingMeta.lastRowUid = rowMeta.rowUid;
                    existingMeta.lastRowIndex = rowMeta.rowIndex;
                }

                existingMeta.rowCount = Object.keys(existingMeta.rowBindings).length;

                var newTag = 'table-meta:' + JSON.stringify(existingMeta);

                var controlTagAfter = newTag;
                var controlSetSuccess = false;
                if (targetControl && typeof targetControl.SetTag === 'function') {
                    try {
                        targetControl.SetTag(newTag);
                        var refreshedTag = safeCall(targetControl, 'GetTag');
                        if (refreshedTag.ok && typeof refreshedTag.value === 'string') {
                            controlTagAfter = refreshedTag.value;
                        }
                        controlSetSuccess = true;
                    } catch (setTagErr) {
                        console.log('[SelectionBinding] 设置表格内容控件Tag失败:', setTagErr);
                    }
                }

                var tableElement = null;
                if (targetControl) {
                    var rangeRes = safeCall(targetControl, 'GetRange');
                    if (rangeRes.ok && rangeRes.value && typeof rangeRes.value.GetElementsCount === 'function' && typeof rangeRes.value.GetElement === 'function') {
                        try {
                            var elemCount = rangeRes.value.GetElementsCount();
                            for (var elemIdx = 0; elemIdx < elemCount; elemIdx++) {
                                var element = rangeRes.value.GetElement(elemIdx);
                                if (element && typeof element.GetClassType === 'function') {
                                    var classType = element.GetClassType();
                                    if (classType === 'CTable') {
                                        tableElement = element;
                                        break;
                                    }
                                }
                            }
                        } catch (rangeErr) {
                            console.log('[SelectionBinding] 通过控件范围查找表格失败:', rangeErr);
                        }
                    }
                }

                if (!tableElement && typeof persistencePayload.tableIndex === 'number' && persistencePayload.tableIndex >= 0) {
                    if (typeof doc.GetElement === 'function' && typeof doc.GetElementsCount === 'function') {
                        try {
                            var totalElements = doc.GetElementsCount();
                            if (persistencePayload.tableIndex < totalElements) {
                                var candidate = doc.GetElement(persistencePayload.tableIndex);
                                if (candidate && typeof candidate.GetClassType === 'function' && candidate.GetClassType() === 'CTable') {
                                    tableElement = candidate;
                                }
                            }
                        } catch (idxErr) {
                            console.log('[SelectionBinding] 根据索引查找表格失败:', idxErr);
                        }
                    }
                }

                var descriptionTag = persistencePayload.tableDescriptionTag || '';
                if (tableElement && typeof tableElement.SetTableDescription === 'function') {
                    try {
                        tableElement.SetTableDescription(newTag);
                        var descRes = safeCall(tableElement, 'GetTableDescription');
                        if (descRes.ok && typeof descRes.value === 'string') {
                            descriptionTag = descRes.value;
                        } else {
                            descriptionTag = newTag;
                        }
                    } catch (setDescErr) {
                        console.log('[SelectionBinding] 设置表格描述失败:', setDescErr);
                        if (!descriptionTag) {
                            descriptionTag = newTag;
                        }
                    }
                } else if (!descriptionTag) {
                    descriptionTag = newTag;
                }

                return {
                    success: true,
                    data: {
                        tableTag: controlTagAfter,
                        tableDescriptionTag: descriptionTag,
                        tableMeta: existingMeta,
                        persistenceLog: {
                            updatedAt: nowIso,
                            controlMatched: !!targetControl,
                            controlTagUpdated: controlSetSuccess,
                            storedRowUid: storedRowUid,
                            rowCount: existingMeta.rowCount
                        }
                    }
                };

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    // 更新行绑定数据
    async updateRowBinding(updateData) {
        const options = updateData || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            rid: options.rid,
            tag: options.tag,
            children: options.children || [],
            metadata: options.metadata || {}
        };

        const updateResult = await this.editor.runInDoc(function () {
            function parsePrefixedJsonTag(tag, prefix) {
                if (typeof tag !== 'string' || tag.indexOf(prefix) !== 0) {
                    return null;
                }
                try {
                    return JSON.parse(tag.slice(prefix.length));
                } catch (err) {
                    console.log('[SelectionBinding] 解析标签失败:', tag, err);
                    return null;
                }
            }

            var doc = Api.GetDocument();
            var scope = Asc.scope;

            console.log('=== 更新行绑定数据 ===');
            console.log('更新参数:', scope);

            try {
                var allControls = doc.GetAllContentControls();
                console.log('文档中的Content Controls数量:', allControls.length);

                for (var i = 0; i < allControls.length; i++) {
                    var control = allControls[i];

                    try {
                        var tag = '';
                        if (typeof control.GetTag === 'function') {
                            tag = control.GetTag();
                        }

                        var alias = '';
                        if (typeof control.GetAlias === 'function') {
                            alias = control.GetAlias();
                        }

                        var isMatch = false;
                        var controlInternalId = '';
                        try {
                            if (typeof control.GetInternalId === 'function') {
                                controlInternalId = control.GetInternalId();
                            }
                        } catch (e) {
                            console.log('获取控件internalId失败:', e);
                        }

                        if (scope.rid && controlInternalId && controlInternalId === scope.rid) {
                            isMatch = true;
                        } else if (scope.rid && (tag.indexOf(scope.rid) !== -1 || alias.indexOf(scope.rid) !== -1)) {
                            isMatch = true;
                        }

                        if (isMatch) {
                            console.log('找到匹配的行控件:', i, 'Tag:', tag, 'Alias:', alias);

                            if (typeof control.SetTag === 'function' && scope.tag) {
                                control.SetTag(scope.tag);
                                console.log('更新行Tag为:', scope.tag);
                            }

                            if (typeof control.SetAlias === 'function') {
                                var childrenCount = scope.children ? scope.children.length : 0;
                                var newAlias = '表格行绑定: ' + scope.rid + ' (子项: ' + childrenCount + ')';
                                control.SetAlias(newAlias);
                                console.log('更新行Alias为:', newAlias);
                            }

                            var updatedTag = scope.tag || tag;
                            var parsedRowMeta = parsePrefixedJsonTag(updatedTag, 'row-meta:');
                            var tableUid = parsedRowMeta && parsedRowMeta.tableUid ? parsedRowMeta.tableUid : '';
                            var tableInternalId = parsedRowMeta && parsedRowMeta.tableInternalId ? parsedRowMeta.tableInternalId : '';
                            var rowIndex = parsedRowMeta && typeof parsedRowMeta.rowIndex === 'number' ? parsedRowMeta.rowIndex : null;

                            var tableControlId = '';
                            var tableTag = '';
                            var tableMeta = null;

                            if (tableUid) {
                                for (var j = 0; j < allControls.length; j++) {
                                    var tableCandidate = allControls[j];
                                    if (!tableCandidate || tableCandidate === control) {
                                        continue;
                                    }

                                    var candidateTag = '';
                                    try {
                                        if (typeof tableCandidate.GetTag === 'function') {
                                            candidateTag = tableCandidate.GetTag();
                                        }
                                    } catch (candidateTagErr) {
                                        console.log('[SelectionBinding] 获取表格控件Tag失败:', candidateTagErr);
                                    }

                                    var parsedTableMeta = parsePrefixedJsonTag(candidateTag, 'table-meta:');
                                    if (parsedTableMeta && parsedTableMeta.tableUid === tableUid) {
                                        tableTag = candidateTag;
                                        tableMeta = parsedTableMeta;
                                        if (!tableInternalId && typeof tableCandidate.GetInternalId === 'function') {
                                            try {
                                                tableInternalId = tableCandidate.GetInternalId();
                                            } catch (tableInternalErr) {
                                                console.log('[SelectionBinding] 获取表格控件InternalId失败:', tableInternalErr);
                                            }
                                        }
                                        if (typeof tableCandidate.GetId === 'function') {
                                            try {
                                                tableControlId = tableCandidate.GetId();
                                            } catch (tableIdErr) {
                                                console.log('[SelectionBinding] 获取表格控件Id失败:', tableIdErr);
                                            }
                                        }
                                        break;
                                    }
                                }
                            }

                            return {
                                success: true,
                                message: 'Row binding updated successfully',
                                data: {
                                    controlIndex: i,
                                    oldTag: tag,
                                    newTag: scope.tag,
                                    rid: scope.rid,
                                    childrenCount: scope.children ? scope.children.length : 0,
                                    updatedAt: new Date().toISOString(),
                                    rowMeta: parsedRowMeta,
                                    rowIndex: rowIndex,
                                    rowControlId: controlInternalId,
                                    tableUid: tableUid,
                                    tableInternalId: tableInternalId,
                                    tableControlId: tableControlId,
                                    tableTag: tableTag,
                                    tableMeta: tableMeta,
                                    tableDescriptionTag: ''
                                }
                            };
                        }
                    } catch (errCheck) {
                        console.log('检查行控件出错:', errCheck);
                    }
                }

                console.log('未找到匹配的行控件');
                return {
                    success: false,
                    error: 'No matching row control found for rid: ' + scope.rid
                };

            } catch (errUpdate) {
                console.log('更新行绑定数据出错:', errUpdate);
                return { success: false, error: errUpdate.message };
            }
        });

        if (!updateResult) {
            return { success: false, error: 'Row binding update failed to execute' };
        }

        if (!updateResult.success) {
            return updateResult;
        }

        const resultData = updateResult.data || {};
        const persistencePayload = {
            tableUid: resultData.tableUid || (resultData.rowMeta && resultData.rowMeta.tableUid) || '',
            tableInternalId: resultData.tableInternalId || (resultData.rowMeta && resultData.rowMeta.tableInternalId) || '',
            tableControlId: resultData.tableControlId || '',
            tableTag: resultData.tableTag || '',
            tableDescriptionTag: resultData.tableDescriptionTag || '',
            existingMeta: resultData.tableMeta || null,
            rowMeta: resultData.rowMeta || null,
            rowIndex: resultData.rowIndex,
            rowControlId: resultData.rowControlId || (options && options.rid ? options.rid : ''),
            action: 'updated'
        };

        if (persistencePayload.tableUid || persistencePayload.tableInternalId || persistencePayload.rowMeta) {
            const persistResult = await this.persistRowMetaToTable(persistencePayload);

            if (persistResult && persistResult.success && persistResult.data) {
                updateResult.data.tableTag = persistResult.data.tableTag || updateResult.data.tableTag;
                updateResult.data.tableDescriptionTag = persistResult.data.tableDescriptionTag || updateResult.data.tableDescriptionTag;
                updateResult.data.tableMeta = persistResult.data.tableMeta || updateResult.data.tableMeta;
                if (persistResult.data.persistenceLog) {
                    updateResult.data.persistenceLog = persistResult.data.persistenceLog;
                }
            } else if (persistResult && !persistResult.success) {
                updateResult.data = updateResult.data || {};
                updateResult.data.persistenceWarning = persistResult.error || 'Table metadata persistence failed';
            }
        }

        return updateResult;
    }
}
