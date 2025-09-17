// services/shape-service.js
export class ShapeService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 测试Shape内联插入可能性
    async testShapeInlineInsertion() {
        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            console.log('=== 测试Shape内联插入可能性 ===');

            // 测试方法1: 检查Run对象是否有AddDrawing方法
            var testRun = Api.CreateRun();
            console.log('Run对象的方法:', Object.getOwnPropertyNames(testRun).filter(name => name.includes('Add')));
            console.log('Run.AddDrawing 是否存在:', typeof testRun.AddDrawing === 'function');

            // 测试方法2: 检查Range对象是否有AddDrawing方法
            var range = doc.GetRangeBySelect();
            if (range) {
                console.log('Range对象的方法:', Object.getOwnPropertyNames(range).filter(name => name.includes('Add')));
                console.log('Range.AddDrawing 是否存在:', typeof range.AddDrawing === 'function');
            }

            // 测试方法3: 检查Shape是否可以作为DocumentElement用于InsertContent
            var testShape = Api.CreateShape("rect", 1000000, 500000,
                Api.CreateSolidFill(Api.CreateRGBColor(255, 0, 0)),
                Api.CreateStroke(0, Api.CreateNoFill()));

            console.log('创建的Shape对象类型:', typeof testShape);
            console.log('Shape对象的方法:', Object.getOwnPropertyNames(testShape).slice(0, 10));

            // 测试方法4: 尝试InsertContent插入Shape
            if (typeof doc.InsertContent === 'function') {
                try {
                    console.log('尝试使用InsertContent插入Shape...');
                    var insertResult = doc.InsertContent([testShape], false);
                    console.log('InsertContent插入Shape结果:', insertResult);
                    if (insertResult) {
                        return {
                            success: true,
                            message: 'Shape可以通过InsertContent插入',
                            method: 'InsertContent'
                        };
                    }
                } catch (e) {
                    console.log('InsertContent插入Shape失败:', e.message);
                }
            }

            // 测试方法5: 创建包含Shape的段落，然后尝试内联插入
            if (typeof doc.InsertContent === 'function') {
                try {
                    console.log('尝试插入包含Shape的段落...');
                    var para = Api.CreateParagraph();
                    para.AddDrawing(testShape);
                    var insertResult = doc.InsertContent([para], true); // isInline = true
                    console.log('插入包含Shape的段落结果:', insertResult);
                    if (insertResult) {
                        return {
                            success: true,
                            message: 'Shape可以在包含的段落中插入（可能不是真正内联）',
                            method: 'InsertContent-Paragraph'
                        };
                    }
                } catch (e) {
                    console.log('插入包含Shape的段落失败:', e.message);
                }
            }

            return {
                success: false,
                message: 'Shape不支持内联插入，只能添加到段落级别',
                recommendation: '不建议创建ShapeService，因为功能受限'
            };
        });
    }

    // 在光标位置的段落中插入Shape（段落级别，最可靠的方法）
    async insertShapeInParagraph(opts) {
        const options = opts || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            shapeType: options.shapeType || 'rect',
            width: options.width || 100,
            height: options.height || 50,
            fillColor: options.fillColor || [255, 111, 61],
            strokeColor: options.strokeColor || [0, 0, 0],
            strokeWidth: options.strokeWidth || 0,
            text: options.text || ''
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            var shapeType = (Asc.scope && Asc.scope.shapeType) || 'rect';
            var width = (Asc.scope && Asc.scope.width) || 100;
            var height = (Asc.scope && Asc.scope.height) || 50;
            var fillColor = (Asc.scope && Asc.scope.fillColor) || [255, 111, 61];
            var strokeColor = (Asc.scope && Asc.scope.strokeColor) || [0, 0, 0];
            var strokeWidth = (Asc.scope && Asc.scope.strokeWidth) || 0;
            var text = (Asc.scope && Asc.scope.text) || '';

            console.log('=== Shape段落级别插入 ===');
            console.log('参数:', { shapeType, width, height, fillColor, strokeColor, text });

            try {
                // 获取光标所在的段落（增强调试版本）
                var range = doc.GetRangeBySelect();
                var targetParagraph = null;

                console.log('=== 详细调试光标位置获取 ===');
                console.log('Range对象:', range);
                console.log('Range类型:', typeof range);

                if (range) {
                    console.log('Range对象存在，检查可用方法...');
                    console.log('Range的所有方法:', Object.getOwnPropertyNames(range));

                    // 方法1: 尝试GetParagraph
                    if (typeof range.GetParagraph === 'function') {
                        console.log('尝试range.GetParagraph()...');
                        targetParagraph = range.GetParagraph();
                        console.log('GetParagraph结果:', targetParagraph);
                        console.log('GetParagraph结果类型:', typeof targetParagraph);

                        if (targetParagraph) {
                            console.log('✅ 通过Range.GetParagraph()成功获取到段落');
                        } else {
                            console.log('❌ Range.GetParagraph()返回null/undefined');
                        }
                    } else {
                        console.log('❌ Range对象没有GetParagraph方法');
                    }

                    // 方法2: 如果方法1失败，尝试其他Range方法
                    if (!targetParagraph) {
                        console.log('尝试其他Range方法...');

                        // 检查Range是否有其他获取段落的方法
                        var rangeMethods = Object.getOwnPropertyNames(range).filter(name =>
                            name.toLowerCase().includes('paragraph') ||
                            name.toLowerCase().includes('para') ||
                            name.toLowerCase().includes('element')
                        );
                        console.log('Range中包含paragraph/para/element的方法:', rangeMethods);

                        // 尝试一些可能的方法
                        if (typeof range.GetParent === 'function') {
                            console.log('尝试range.GetParent()...');
                            var parent = range.GetParent();
                            console.log('GetParent结果:', parent, typeof parent);
                        }

                        if (typeof range.GetElement === 'function') {
                            console.log('尝试range.GetElement()...');
                            var element = range.GetElement();
                            console.log('GetElement结果:', element, typeof element);
                            // 检查这个element是否是段落或包含段落
                            if (element && typeof element.GetClassType === 'function') {
                                console.log('Element类型:', element.GetClassType());
                            }
                        }
                    }

                    // 方法3: 尝试通过Document的当前选择获取
                    if (!targetParagraph) {
                        console.log('尝试通过文档选择获取段落...');
                        // 检查文档是否有获取当前元素的方法
                        var docMethods = Object.getOwnPropertyNames(doc).filter(name =>
                            name.toLowerCase().includes('current') ||
                            name.toLowerCase().includes('active') ||
                            name.toLowerCase().includes('select')
                        );
                        console.log('Document中包含current/active/select的方法:', docMethods);
                    }
                } else {
                    console.log('❌ 无法获取Range对象');
                }

                // 如果所有方法都失败，使用第一个段落作为后备
                if (!targetParagraph) {
                    console.log('所有获取当前段落的方法都失败，检查可用段落...');

                    // 检查文档中的段落数量
                    var paragraphCount = 0;
                    for (var i = 0; i < 20; i++) { // 检查前20个元素
                        var element = doc.GetElement(i);
                        if (element) {
                            console.log(`文档元素[${i}]:`, element, typeof element);
                            if (typeof element.GetClassType === 'function') {
                                console.log(`  类型: ${element.GetClassType()}`);
                            }
                            paragraphCount++;
                        } else {
                            break;
                        }
                    }
                    console.log(`文档总共有 ${paragraphCount} 个元素`);

                    if (typeof doc.GetElement === 'function') {
                        targetParagraph = doc.GetElement(0);
                        if (targetParagraph) {
                            console.log('使用第一个段落作为后备');
                            console.log('第一个段落对象:', targetParagraph);
                            if (typeof targetParagraph.GetClassType === 'function') {
                                console.log('第一个段落类型:', targetParagraph.GetClassType());
                            }
                        } else {
                            console.log('连第一个段落都无法获取');
                        }
                    }
                }

                if (!targetParagraph) {
                    return { success: false, error: 'No paragraph available for Shape insertion' };
                }

                // 创建填充和描边
                var fill = Api.CreateSolidFill(Api.CreateRGBColor(fillColor[0], fillColor[1], fillColor[2]));
                var stroke = strokeWidth > 0
                    ? Api.CreateStroke(strokeWidth * 36000, Api.CreateSolidFill(Api.CreateRGBColor(strokeColor[0], strokeColor[1], strokeColor[2])))
                    : Api.CreateStroke(0, Api.CreateNoFill());

                // 创建形状
                var shape = Api.CreateShape(shapeType, width * 36000, height * 36000, fill, stroke);

                if (!shape) {
                    console.log('创建Shape失败');
                    return { success: false, error: 'Failed to create shape' };
                }

                console.log('Shape创建成功');

                // 如果有文本，添加到形状中
                if (text) {
                    var docContent = shape.GetDocContent();
                    if (docContent) {
                        docContent.RemoveAllElements();
                        var paragraph = Api.CreateParagraph();
                        paragraph.AddText(text);
                        docContent.AddElement(0, paragraph);
                        console.log('添加文本到Shape:', text);
                    }
                }

                // 段落级别插入（最可靠的方法）
                if (typeof targetParagraph.AddDrawing === 'function') {
                    var addResult = targetParagraph.AddDrawing(shape);
                    console.log('段落AddDrawing执行结果:', addResult);
                    console.log('Shape已添加到段落');
                    return {
                        success: true,
                        message: 'Shape inserted in paragraph at cursor location',
                        method: 'Paragraph-AddDrawing',
                        parameters: { shapeType, width, height, text }
                    };
                } else {
                    return { success: false, error: 'AddDrawing method not available on paragraph' };
                }

            } catch (e) {
                console.log('Shape insertion failed:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }
    async insertShapeInline(opts) {
        const options = opts || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            shapeType: options.shapeType || 'rect',
            width: options.width || 100,
            height: options.height || 50,
            fillColor: options.fillColor || [255, 111, 61],
            strokeColor: options.strokeColor || [0, 0, 0],
            strokeWidth: options.strokeWidth || 0,
            text: options.text || ''
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            var shapeType = (Asc.scope && Asc.scope.shapeType) || 'rect';
            var width = (Asc.scope && Asc.scope.width) || 100;
            var height = (Asc.scope && Asc.scope.height) || 50;
            var fillColor = (Asc.scope && Asc.scope.fillColor) || [255, 111, 61];
            var strokeColor = (Asc.scope && Asc.scope.strokeColor) || [0, 0, 0];
            var strokeWidth = (Asc.scope && Asc.scope.strokeWidth) || 0;
            var text = (Asc.scope && Asc.scope.text) || '';

            console.log('=== Shape内联插入 ===');
            console.log('参数:', { shapeType, width, height, fillColor, strokeColor, text });

            try {
                // 创建填充和描边
                var fill = Api.CreateSolidFill(Api.CreateRGBColor(fillColor[0], fillColor[1], fillColor[2]));
                var stroke = strokeWidth > 0
                    ? Api.CreateStroke(strokeWidth * 36000, Api.CreateSolidFill(Api.CreateRGBColor(strokeColor[0], strokeColor[1], strokeColor[2])))
                    : Api.CreateStroke(0, Api.CreateNoFill());

                // 创建形状
                var shape = Api.CreateShape(shapeType, width * 36000, height * 36000, fill, stroke);

                if (!shape) {
                    console.log('创建Shape失败');
                    return { success: false, error: 'Failed to create shape' };
                }

                console.log('Shape创建成功');

                // 如果有文本，添加到形状中
                if (text) {
                    var docContent = shape.GetDocContent();
                    if (docContent) {
                        docContent.RemoveAllElements();
                        var paragraph = Api.CreateParagraph();
                        paragraph.AddText(text);
                        docContent.AddElement(0, paragraph);
                        console.log('添加文本到Shape:', text);
                    }
                }

                // 方法1: 优先尝试InsertContent直接插入（基于测试成功）
                if (typeof doc.InsertContent === 'function') {
                    console.log('尝试使用InsertContent内联插入Shape...');
                    try {
                        // 尝试不同的插入参数
                        console.log('尝试 isInline=false...');
                        var insertResult1 = doc.InsertContent([shape], false);
                        console.log('InsertContent(false)结果:', insertResult1);

                        console.log('尝试 isInline=true...');
                        var insertResult2 = doc.InsertContent([shape], true);
                        console.log('InsertContent(true)结果:', insertResult2);

                        // 检查文档是否真的有变化
                        console.log('检查文档元素数量...');
                        var elementCount = 0;
                        for (var i = 0; i < 10; i++) {
                            if (doc.GetElement(i)) {
                                elementCount++;
                            } else {
                                break;
                            }
                        }
                        console.log('文档元素总数:', elementCount);

                        if (insertResult1 || insertResult2) {
                            console.log('InsertContent报告成功，但可能未实际插入');
                            // 不直接返回，继续尝试其他方法
                        }
                    } catch (e) {
                        console.log('InsertContent失败:', e.message);
                    }
                }

                // 方法2: 尝试通过Run.AddDrawing内联插入
                console.log('尝试通过Run.AddDrawing内联插入...');
                var range = doc.GetRangeBySelect();
                if (range) {
                    try {
                        var run = Api.CreateRun();
                        console.log('创建Run成功，Run类型:', typeof run);

                        if (typeof run.AddDrawing === 'function') {
                            console.log('Run.AddDrawing方法存在，尝试添加Shape...');
                            var addDrawingResult = run.AddDrawing(shape);
                            console.log('Run.AddDrawing结果:', addDrawingResult);

                            // 尝试通过InsertContent插入这个包含Shape的Run
                            if (typeof doc.InsertContent === 'function') {
                                console.log('尝试插入包含Shape的Run...');
                                var insertRunResult = doc.InsertContent([run], true);
                                console.log('插入Run结果:', insertRunResult);

                                if (insertRunResult) {
                                    console.log('通过Run.AddDrawing + InsertContent插入可能成功');
                                    return {
                                        success: true,
                                        message: 'Shape inserted via Run.AddDrawing + InsertContent',
                                        method: 'Run-AddDrawing-InsertContent',
                                        parameters: { shapeType, width, height, text }
                                    };
                                }
                            }

                            // 尝试将Run添加到当前段落
                            console.log('尝试将含Shape的Run添加到段落...');
                            var para = range.GetParagraph ? range.GetParagraph() : null;
                            if (para && typeof para.AddElement === 'function') {
                                var addElementResult = para.AddElement(run);
                                console.log('段落AddElement结果:', addElementResult);

                                if (addElementResult !== false) {
                                    console.log('通过段落AddElement插入Run成功');
                                    return {
                                        success: true,
                                        message: 'Shape inserted via Run.AddDrawing + Paragraph.AddElement',
                                        method: 'Run-AddDrawing-Paragraph',
                                        parameters: { shapeType, width, height, text }
                                    };
                                }
                            }
                        }
                    } catch (e) {
                        console.log('Run.AddDrawing方法失败:', e.message);
                    }
                }

                // 方法3: 回退到段落级别插入
                console.log('内联插入失败，回退到段落级别插入...');
                var targetParagraph = null;

                if (range && typeof range.GetParagraph === 'function') {
                    targetParagraph = range.GetParagraph();
                }

                if (!targetParagraph && typeof doc.GetElement === 'function') {
                    targetParagraph = doc.GetElement(0);
                }

                if (targetParagraph && typeof targetParagraph.AddDrawing === 'function') {
                    targetParagraph.AddDrawing(shape);
                    console.log('Shape添加到段落成功');
                    return {
                        success: true,
                        message: 'Shape inserted at paragraph level (inline insertion failed)',
                        method: 'Paragraph-AddDrawing',
                        parameters: { shapeType, width, height, text }
                    };
                }

                return { success: false, error: 'All insertion methods failed' };

            } catch (e) {
                console.log('Shape insertion failed:', e);
                console.log('Error stack:', e.stack);
                return { success: false, error: e.message };
            }
        });
    }
    async insertShape(opts) {
        const options = opts || {};

        window.Asc = window.Asc || {};
        window.Asc.scope = {
            shapeType: options.shapeType || 'rect',
            width: options.width || 100,
            height: options.height || 50,
            fillColor: options.fillColor || [255, 111, 61],
            strokeColor: options.strokeColor || [0, 0, 0],
            strokeWidth: options.strokeWidth || 0,
            text: options.text || ''
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            var shapeType = (Asc.scope && Asc.scope.shapeType) || 'rect';
            var width = (Asc.scope && Asc.scope.width) || 100;
            var height = (Asc.scope && Asc.scope.height) || 50;
            var fillColor = (Asc.scope && Asc.scope.fillColor) || [255, 111, 61];
            var strokeColor = (Asc.scope && Asc.scope.strokeColor) || [0, 0, 0];
            var strokeWidth = (Asc.scope && Asc.scope.strokeWidth) || 0;
            var text = (Asc.scope && Asc.scope.text) || '';

            console.log('=== Shape插入 (仅段落级别) ===');
            console.log('参数:', { shapeType, width, height, fillColor, strokeColor, text });

            try {
                // 获取目标段落
                var range = doc.GetRangeBySelect();
                var targetParagraph = null;

                if (range && typeof range.GetParagraph === 'function') {
                    targetParagraph = range.GetParagraph();
                }

                if (!targetParagraph && typeof doc.GetElement === 'function') {
                    targetParagraph = doc.GetElement(0);
                }

                if (!targetParagraph) {
                    return { success: false, error: 'No paragraph available' };
                }

                // 创建填充和描边
                var fill = Api.CreateSolidFill(Api.CreateRGBColor(fillColor[0], fillColor[1], fillColor[2]));
                var stroke = strokeWidth > 0
                    ? Api.CreateStroke(strokeWidth * 36000, Api.CreateSolidFill(Api.CreateRGBColor(strokeColor[0], strokeColor[1], strokeColor[2])))
                    : Api.CreateStroke(0, Api.CreateNoFill());

                // 创建形状
                var shape = Api.CreateShape(shapeType, width * 36000, height * 36000, fill, stroke);

                if (!shape) {
                    return { success: false, error: 'Failed to create shape' };
                }

                // 如果有文本，添加到形状中
                if (text) {
                    var docContent = shape.GetDocContent();
                    if (docContent) {
                        docContent.RemoveAllElements();
                        var paragraph = Api.CreateParagraph();
                        paragraph.AddText(text);
                        docContent.AddElement(0, paragraph);
                    }
                }

                // 添加到段落
                if (typeof targetParagraph.AddDrawing === 'function') {
                    targetParagraph.AddDrawing(shape);
                    return {
                        success: true,
                        message: 'Shape inserted in paragraph',
                        parameters: { shapeType, width, height, text }
                    };
                } else {
                    return { success: false, error: 'AddDrawing method not available' };
                }

            } catch (e) {
                console.log('Shape insertion failed:', e);
                return { success: false, error: e.message };
            }
        });
    }
}