// services/link-service.js
export class LinkService {
    constructor(editorService) {
        this.editor = editorService;
    }

    // 处理链接点击事件 - 使用sdt-service.js的exact模式
    async handleLinkClick() {
        return new Promise((resolve) => {
            this.editor.runInDoc(function () {
                var doc = Api.GetDocument();
                var range = doc.GetRangeBySelect();

                console.log('=== Handling link click (sdt-service pattern) ===');
                console.log('Current range:', range);

                if (!range) {
                    console.log('No range selected');
                    return null;
                }

                // 完全按照sdt-service.js的方式检测控件
                var ctrls = doc.GetAllContentControls();
                console.log('Found content controls:', ctrls.length);

                for (var i = 0; i < ctrls.length; i++) {
                    var ctrl = ctrls[i];
                    console.log('Checking control', i, 'with tag:', ctrl.GetTag ? ctrl.GetTag() : 'no tag');

                    // 使用与sdt-service.js完全相同的hit检测逻辑
                    var hit =
                        (typeof range.IsInContentControl === 'function' && range.IsInContentControl(ctrl)) ||
                        (typeof ctrl.IsRangeIn === 'function' && ctrl.IsRangeIn(range)) ||
                        (typeof ctrl.IsSelected === 'function' && ctrl.IsSelected());

                    console.log('Control hit status:', hit);

                    if (hit) {
                        // 检查是否是我们的链接控件
                        var tag = ctrl.GetTag ? ctrl.GetTag() : '';
                        if (tag && tag.startsWith('link-data:')) {
                            console.log('Found link control!');
                            try {
                                var jsonStr = tag.substring('link-data:'.length);
                                var jsonData = jsonStr ? JSON.parse(jsonStr) : {};

                                // 返回控件信息和数据
                                var result = {
                                    controlId: ctrl.GetId ? ctrl.GetId() : 'unknown',
                                    controlTitle: ctrl.GetAlias ? ctrl.GetAlias() : '',
                                    tag: tag,
                                    data: jsonData
                                };

                                console.log('Link click detected, returning data:', result);
                                return result;
                            } catch (e) {
                                console.log('Error parsing link data:', e);
                                return null;
                            }
                        }
                    }
                }

                console.log('No link control found at current position');
                return null;

            }, { async: false, cb: (res) => resolve(res) });
        });
    }

    insertLinkInline(opts) {
        const options = opts || {};

        // 设置 Asc.scope 让文档内脚本能访问
        window.Asc = window.Asc || {};
        window.Asc.scope = {
            text: options.text || '点击这里',
            url: options.url || '',
            style: options.style || {},
            json: options.json || ''
        };

        return this.editor.runInDoc(function () {
            var doc = Api.GetDocument();

            // —— 只读 Asc.scope —— //
            var text = (Asc.scope && Asc.scope.text) || '点击这里';
            var url = (Asc.scope && Asc.scope.url) || '';
            var style = (Asc.scope && Asc.scope.style) || {};
            var bold = (style.bold !== undefined) ? !!style.bold : true;
            var underline = (style.underline !== undefined) ? !!style.underline : true;
            var json = (Asc.scope && Asc.scope.json) || '';

            function toRgb(c) {
                if (!c) return [25, 118, 210];                       // 默认 #1976D2
                if (Array.isArray(c) && c.length === 3) return c;  // [r,g,b]
                if (typeof c === 'string' && /^#?[0-9a-fA-F]{6}$/.test(c)) {
                    var hex = c[0] === '#' ? c.slice(1) : c;
                    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
                }
                return [25, 118, 210];
            }
            var rgb = toRgb(style.color);

            console.log('=== Link insertion (using sdt-service pattern) ===');
            console.log('Parameters:', { text: text, url: url, json: json, style: style });
            console.log('RGB color:', rgb);
            console.log('Bold:', bold, 'Underline:', underline);

            var jsonData = json ? JSON.stringify(json) : '';

            try {
                if (!url || !String(url).trim()) {
                    // 保持Content Control以支持点击事件，但尝试不同的样式设置方式
                    console.log('=== Using Content Control with improved styling ===');

                    var sdt = doc.AddComboBoxContentControl();
                    sdt.SetTag('link-data:' + jsonData);
                    if (sdt.SetAlias) {
                        sdt.SetAlias('可点击链接');
                    }

                    var content = sdt.GetContent();
                    content.RemoveAllElements();
                    var p = content.GetElement(0);

                    if (p) {
                        // 方法1: 创建自定义字符样式
                        var linkStyle = null;
                        if (typeof doc.CreateStyle === 'function') {
                            linkStyle = doc.CreateStyle('LinkStyle_' + Date.now(), 'character');
                            if (linkStyle) {
                                console.log('Created custom link style');
                                var styleTextPr = linkStyle.GetTextPr();
                                if (styleTextPr) {
                                    if (typeof styleTextPr.SetColor === 'function') {
                                        styleTextPr.SetColor(rgb[0], rgb[1], rgb[2]);
                                        console.log('Set style color via TextPr');
                                    }
                                    if (typeof styleTextPr.SetBold === 'function') styleTextPr.SetBold(bold);
                                    if (typeof styleTextPr.SetUnderline === 'function') styleTextPr.SetUnderline(underline);
                                    console.log('Set style properties');
                                }
                            }
                        }

                        var run = Api.CreateRun();
                        run.AddText(text);

                        // 方法2: 优先尝试应用自定义样式
                        if (linkStyle && typeof run.SetStyle === 'function') {
                            var styleResult = run.SetStyle(linkStyle);
                            console.log('Applied custom style result:', styleResult);
                        }

                        // 方法3: 设置直接样式属性
                        console.log('Setting direct styles...');
                        console.log('Available run methods:', Object.getOwnPropertyNames(run).filter(name => name.includes('Set')));

                        if (typeof run.SetBold === 'function') {
                            var boldResult = run.SetBold(bold);
                            console.log('SetBold result:', boldResult);
                        }
                        if (typeof run.SetUnderline === 'function') {
                            var underlineResult = run.SetUnderline(underline);
                            console.log('SetUnderline result:', underlineResult);
                        }
                        if (typeof run.SetColor === 'function') {
                            var colorResult = run.SetColor(rgb[0], rgb[1], rgb[2]);
                            console.log('SetColor result:', colorResult, 'RGB:', rgb[0], rgb[1], rgb[2]);
                        }

                        // 方法4: 使用TextPr对象设置样式
                        if (typeof Api.CreateTextPr === 'function') {
                            var textPr = Api.CreateTextPr();
                            if (textPr) {
                                console.log('Created TextPr object');
                                console.log('TextPr methods:', Object.getOwnPropertyNames(textPr).filter(name => name.includes('Set')));

                                if (typeof textPr.SetColor === 'function') {
                                    textPr.SetColor(rgb[0], rgb[1], rgb[2]);
                                    console.log('Set TextPr color');
                                }
                                if (typeof textPr.SetBold === 'function') textPr.SetBold(bold);
                                if (typeof textPr.SetUnderline === 'function') textPr.SetUnderline(underline);

                                if (typeof run.SetTextPr === 'function') {
                                    var textPrResult = run.SetTextPr(textPr);
                                    console.log('SetTextPr result:', textPrResult);
                                }
                            }
                        }

                        // 添加到段落
                        p.AddElement(run);
                        console.log('Added run to paragraph');

                        // 方法5: 在添加后尝试重新设置样式
                        console.log('Re-applying styles after adding to paragraph...');
                        if (typeof run.SetColor === 'function') {
                            var reColorResult = run.SetColor(rgb[0], rgb[1], rgb[2]);
                            console.log('Re-set color result:', reColorResult);
                        }

                        // 方法6: 尝试对段落设置默认样式
                        if (typeof p.SetColor === 'function') {
                            var pColorResult = p.SetColor(rgb[0], rgb[1], rgb[2]);
                            console.log('Paragraph SetColor result:', pColorResult);
                        }

                        // 方法7: 尝试Content Control级别的样式设置
                        console.log('Trying Content Control level styling...');
                        if (typeof sdt.SetColor === 'function') {
                            var sdtColorResult = sdt.SetColor(rgb[0], rgb[1], rgb[2]);
                            console.log('Content Control SetColor result:', sdtColorResult);
                        }
                    }

                    console.log('Content control created with click support');
                } else {
                    // 有URL的情况保持不变
                    this.insertWithContentControl(text, url, rgb, bold, underline, jsonData, doc);
                }

                console.log('Link insertion completed successfully');

            } catch (e) {
                console.log('Link insertion failed:', e);
                console.log('Error stack:', e.stack);
            }
        });
    }

    // 辅助方法：使用Content Control插入
    insertWithContentControl(text, url, rgb, bold, underline, jsonData, doc) {
        if (!url || !String(url).trim()) {
            // 空URL - 使用Content Control方式
            var sdt = doc.AddComboBoxContentControl();
            sdt.SetTag('link-data:' + jsonData);
            if (sdt.SetAlias) {
                sdt.SetAlias('可点击链接');
            }

            var content = sdt.GetContent();
            content.RemoveAllElements();
            var p = content.GetElement(0);

            if (p) {
                var run = Api.CreateRun();
                run.AddText(text);

                // 尝试各种样式设置方法
                if (run.SetBold) run.SetBold(bold);
                if (run.SetUnderline) run.SetUnderline(underline);
                if (run.SetColor) run.SetColor(rgb[0], rgb[1], rgb[2]);

                var textPr = Api.CreateTextPr();
                if (textPr && run.SetTextPr) {
                    if (textPr.SetColor) textPr.SetColor(rgb[0], rgb[1], rgb[2]);
                    if (textPr.SetBold) textPr.SetBold(bold);
                    if (textPr.SetUnderline) textPr.SetUnderline(underline);
                    run.SetTextPr(textPr);
                }

                p.AddElement(run);
            }
            console.log('Added content control with JSON data');
        } else {
            // 有URL - 使用超链接
            var sdt = doc.AddComboBoxContentControl();
            sdt.SetTag('link-data:' + jsonData);
            if (sdt.SetAlias) {
                sdt.SetAlias('超链接: ' + url);
            }

            var content = sdt.GetContent();
            content.RemoveAllElements();
            var p = content.GetElement(0);

            if (p) {
                var run = Api.CreateRun();
                if (run.SetBold) run.SetBold(bold);
                if (run.SetUnderline) run.SetUnderline(underline);
                if (run.SetColor) run.SetColor(rgb[0], rgb[1], rgb[2]);

                var hyperlink = Api.CreateHyperlink(url, text);
                run.AddElement(hyperlink);
                p.AddElement(run);
            }
            console.log('Added hyperlink content control');
        }
    }
}
