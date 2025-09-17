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

        return this.editor.runInDoc(function () {
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

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();
            var scope = Asc.scope;

            console.log('=== 更新行绑定数据 ===');
            console.log('更新参数:', scope);

            try {
                // 获取所有Content Control
                var allControls = doc.GetAllContentControls();
                console.log('文档中的Content Controls数量:', allControls.length);

                // 查找匹配rid的行控件
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

                        // 检查是否匹配rid
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

                        // 优先使用internalId匹配
                        if (scope.rid && controlInternalId && controlInternalId === scope.rid) {
                            isMatch = true;
                        }
                        // 备选方案
                        else if (scope.rid && (tag.includes(scope.rid) || alias.includes(scope.rid))) {
                            isMatch = true;
                        }

                        if (isMatch) {
                            console.log('找到匹配的行控件:', i, 'Tag:', tag, 'Alias:', alias);

                            // 更新Tag
                            if (typeof control.SetTag === 'function' && scope.tag) {
                                control.SetTag(scope.tag);
                                console.log('更新行Tag为:', scope.tag);
                            }

                            // 更新Alias（包含子项信息）
                            if (typeof control.SetAlias === 'function') {
                                var childrenCount = scope.children ? scope.children.length : 0;
                                var newAlias = '表格行绑定: ' + scope.rid + ' (子项: ' + childrenCount + ')';
                                control.SetAlias(newAlias);
                                console.log('更新行Alias为:', newAlias);
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
                                    updatedAt: new Date().toISOString()
                                }
                            };
                        }
                    } catch (e) {
                        console.log('检查行控件出错:', e);
                    }
                }

                // 没有找到匹配的控件
                console.log('未找到匹配的行控件');
                return {
                    success: false,
                    error: 'No matching row control found for rid: ' + scope.rid
                };

            } catch (e) {
                console.log('更新行绑定数据出错:', e);
                return { success: false, error: e.message };
            }
        });
    }
}