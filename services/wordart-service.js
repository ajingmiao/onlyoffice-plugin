// services/wordart-service.js
export class WordArtService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 在光标位置插入WordArt
    async insertWordArt(opts) {
        const options = opts || {};

        // 设置 Asc.scope 让文档内脚本能访问
        window.Asc = window.Asc || {};
        window.Asc.scope = {
            text: options.text || 'onlyoffice',
            fontSize: options.fontSize || 30,
            bold: options.bold !== undefined ? options.bold : true,
            caps: options.caps !== undefined ? options.caps : true,
            fontFamily: options.fontFamily || 'Comic Sans MS',
            textColor: options.textColor || [51, 51, 51],
            fillColor: options.fillColor || [255, 111, 61],
            strokeColor: options.strokeColor || [51, 51, 51],
            strokeWidth: options.strokeWidth || 1,
            transform: options.transform || 'textArchUp',
            width: options.width || 150,
            height: options.height || 50,
            rotation: options.rotation || 0
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            // 从 Asc.scope 读取参数
            var text = (Asc.scope && Asc.scope.text) || 'onlyoffice';
            var fontSize = (Asc.scope && Asc.scope.fontSize) || 30;
            var bold = (Asc.scope && Asc.scope.bold !== undefined) ? Asc.scope.bold : true;
            var caps = (Asc.scope && Asc.scope.caps !== undefined) ? Asc.scope.caps : true;
            var fontFamily = (Asc.scope && Asc.scope.fontFamily) || 'Comic Sans MS';
            var textColor = (Asc.scope && Asc.scope.textColor) || [51, 51, 51];
            var fillColor = (Asc.scope && Asc.scope.fillColor) || [255, 111, 61];
            var strokeColor = (Asc.scope && Asc.scope.strokeColor) || [51, 51, 51];
            var strokeWidth = (Asc.scope && Asc.scope.strokeWidth) || 1;
            var transform = (Asc.scope && Asc.scope.transform) || 'textArchUp';
            var width = (Asc.scope && Asc.scope.width) || 150;
            var height = (Asc.scope && Asc.scope.height) || 50;
            var rotation = (Asc.scope && Asc.scope.rotation) || 0;

            console.log('=== WordArt Insertion ===');
            console.log('Parameters:', {
                text: text,
                fontSize: fontSize,
                bold: bold,
                caps: caps,
                fontFamily: fontFamily,
                textColor: textColor,
                fillColor: fillColor,
                strokeColor: strokeColor,
                transform: transform,
                width: width,
                height: height
            });

            try {
                // 获取当前光标位置的段落
                var range = doc.GetRangeBySelect();
                var targetParagraph = null;

                console.log('=== 查找光标所在段落 ===');

                if (range) {
                    console.log('找到选择范围');

                    // 方法1: 通过Range获取段落
                    if (typeof range.GetParagraph === 'function') {
                        targetParagraph = range.GetParagraph();
                        if (targetParagraph) {
                            console.log('通过Range.GetParagraph()成功获取段落');
                        } else {
                            console.log('Range.GetParagraph()返回null');
                        }
                    } else {
                        console.log('Range对象没有GetParagraph方法');
                    }

                    // 方法2: 如果方法1失败，尝试通过Document.GetSelection
                    if (!targetParagraph && typeof doc.GetSelection === 'function') {
                        console.log('尝试通过Document.GetSelection获取段落...');
                        try {
                            var selection = doc.GetSelection();
                            if (selection) {
                                console.log('获取到Selection对象');

                                // 尝试获取起始位置
                                if (typeof selection.GetStart === 'function') {
                                    var start = selection.GetStart();
                                    if (start && typeof start.GetParagraph === 'function') {
                                        targetParagraph = start.GetParagraph();
                                        if (targetParagraph) {
                                            console.log('通过Selection.GetStart().GetParagraph()成功获取段落');
                                        }
                                    }
                                }
                            } else {
                                console.log('无法获取Selection对象');
                            }
                        } catch (e) {
                            console.log('GetSelection方法执行失败:', e.message);
                        }
                    } else if (!targetParagraph) {
                        console.log('Document对象没有GetSelection方法');
                    }

                    // 方法3: 尝试通过Range的内容获取段落
                    if (!targetParagraph && typeof range.GetStart === 'function') {
                        console.log('尝试通过Range.GetStart获取段落...');
                        try {
                            var rangeStart = range.GetStart();
                            if (rangeStart && typeof rangeStart.GetParagraph === 'function') {
                                targetParagraph = rangeStart.GetParagraph();
                                if (targetParagraph) {
                                    console.log('通过Range.GetStart().GetParagraph()成功获取段落');
                                }
                            }
                        } catch (e) {
                            console.log('Range.GetStart方法执行失败:', e.message);
                        }
                    } else if (!targetParagraph) {
                        console.log('Range对象没有GetStart方法');
                    }
                } else {
                    console.log('无法获取选择范围');
                }

                // 如果所有方法都失败，使用第一个段落作为后备
                if (!targetParagraph) {
                    console.log('所有获取当前段落的方法都失败，使用第一个段落作为后备');
                    if (typeof doc.GetElement === 'function') {
                        targetParagraph = doc.GetElement(0);
                        if (targetParagraph) {
                            console.log('成功获取第一个段落作为后备');
                        } else {
                            console.log('第一个段落为空');
                        }
                    } else {
                        console.log('Document对象没有GetElement方法');
                    }
                }

                if (!targetParagraph) {
                    console.log('错误：没有可用的段落来插入WordArt');
                    return { success: false, error: 'No paragraph available for WordArt insertion' };
                }

                // 创建文本属性
                var textPr = Api.CreateTextPr();
                if (textPr.SetFontSize) textPr.SetFontSize(fontSize);
                if (textPr.SetBold) textPr.SetBold(bold);
                if (textPr.SetCaps) textPr.SetCaps(caps);
                if (textPr.SetColor) textPr.SetColor(textColor[0], textColor[1], textColor[2], false);
                if (textPr.SetFontFamily) textPr.SetFontFamily(fontFamily);

                // 创建填充和描边
                var fill = Api.CreateSolidFill(Api.CreateRGBColor(fillColor[0], fillColor[1], fillColor[2]));
                var stroke = Api.CreateStroke(strokeWidth * 36000, Api.CreateSolidFill(Api.CreateRGBColor(strokeColor[0], strokeColor[1], strokeColor[2])));

                // 创建WordArt
                var textArt = Api.CreateWordArt(
                    textPr,
                    text,
                    transform,
                    fill,
                    stroke,
                    rotation,
                    width * 36000,
                    height * 36000
                );

                if (!textArt) {
                    console.log('创建WordArt对象失败');
                    return { success: false, error: 'Failed to create WordArt object' };
                }

                console.log('WordArt对象创建成功');

                // 添加到段落
                if (typeof targetParagraph.AddDrawing === 'function') {
                    var addResult = targetParagraph.AddDrawing(textArt);
                    console.log('AddDrawing执行结果:', addResult);
                    console.log('WordArt已添加到目标段落');
                    return {
                        success: true,
                        message: 'WordArt inserted in current paragraph',
                        parameters: { text, fontSize, fontFamily, transform }
                    };
                } else {
                    console.log('错误：段落对象没有AddDrawing方法');
                    return { success: false, error: 'AddDrawing method not available on paragraph' };
                }

            } catch (e) {
                console.log('WordArt insertion failed:', e);
                console.log('Error stack:', e.stack);
                return {
                    success: false,
                    error: e.message || 'Unknown error',
                    stack: e.stack
                };
            }
        });
    }

    // 预设的WordArt样式
    insertPresetWordArt(preset, text) {
        const presets = {
            'classic': {
                text: text || 'CLASSIC',
                fontSize: 36,
                bold: true,
                caps: true,
                fontFamily: 'Times New Roman',
                textColor: [0, 0, 0],
                fillColor: [255, 215, 0],
                strokeColor: [0, 0, 0],
                strokeWidth: 2,
                transform: 'textArchUp'
            },
            'modern': {
                text: text || 'MODERN',
                fontSize: 28,
                bold: true,
                caps: false,
                fontFamily: 'Arial',
                textColor: [255, 255, 255],
                fillColor: [33, 150, 243],
                strokeColor: [13, 71, 161],
                strokeWidth: 1,
                transform: 'textPlain'
            },
            'fun': {
                text: text || 'FUN!',
                fontSize: 32,
                bold: true,
                caps: true,
                fontFamily: 'Comic Sans MS',
                textColor: [255, 255, 255],
                fillColor: [255, 87, 34],
                strokeColor: [191, 54, 12],
                strokeWidth: 2,
                transform: 'textWave1'
            }
        };

        const options = presets[preset] || presets['classic'];
        return this.insertWordArt(options);
    }
}