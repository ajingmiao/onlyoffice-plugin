// import { logger } from '../core/logger.js';

// export class PluginBridge {
//   onInit(cb) {
//     window.Asc = window.Asc || {};
//     window.Asc.plugin = window.Asc.plugin || {};
//     window.Asc.plugin.init = () => {
//       logger.info('plugin init');
//       cb?.();
//     };
//   }

//   callCommand(fn, isAsync = false, callback) {
//     return window.Asc?.plugin?.callCommand?.(fn, isAsync, callback);
//   }

//   onSelectionChanged(cb) {
//     window.Asc = window.Asc || {};
//     window.Asc.plugin = window.Asc.plugin || {};

//     window.Asc.plugin.event_onClick = function(isSelectionUse) {
//       logger.info('🔥 onClick event triggered!', { isSelectionUse });

//       // 执行详细的点击内容分析
//       window.Asc.plugin.executeCommand(function() {
//         try {
//           var doc = Api.GetDocument();
//           var range = doc.GetRangeBySelect();

//           var clickInfo = {
//             hasSelection: !!range,
//             selectionText: '',
//             elementType: 'unknown',
//             contentControlInfo: null,
//             detectedElements: []
//           };

//           // 获取选中的文本
//           if (range && typeof range.GetText === 'function') {
//             clickInfo.selectionText = range.GetText();
//           }

//           // 检测点击位置的元素类型
//           logger.info('=== 点击内容分析 ===');

//           // 1. 扫描文档元素，找到可能被点击的元素
//           var maxElements = 200; // 设置最大扫描数量，避免无限循环
//           var foundTargetElements = 0; // 记录找到的目标元素数量

//           for (var i = 0; i < maxElements; i++) {
//             try {
//               var element = doc.GetElement(i);
//               if (!element) {
//                 logger.info('📋 文档元素扫描完成，共检测', i, '个元素，找到', foundTargetElements, '个相关元素');
//                 break; // 没有更多元素，停止扫描
//               }

//               var elementType = 'unknown';
//               if (typeof element.GetClassType === 'function') {
//                 elementType = element.GetClassType();
//               }

//               var elementInfo = {
//                 index: i,
//                 type: elementType,
//                 isClickTarget: false
//               };

//               // 检查不同类型的元素
//               switch(elementType) {
//                 case 'CTable':
//                   elementInfo.details = '表格元素';
//                   clickInfo.detectedElements.push(elementInfo);
//                   foundTargetElements++;
//                   // 如果有选区且可能在表格中，标记为可能的点击目标
//                   if (range) {
//                     elementInfo.isClickTarget = true;
//                     clickInfo.elementType = 'table';
//                   }
//                   break;

//                 case 'CDocumentParagraph':
//                   elementInfo.details = '段落元素';
//                   if (range) {
//                     // 检查是否点击了段落
//                     try {
//                       if (typeof range.GetParagraph === 'function') {
//                         var para = range.GetParagraph();
//                         if (para === element) {
//                           elementInfo.isClickTarget = true;
//                           clickInfo.elementType = 'paragraph';
//                           foundTargetElements++;
//                         }
//                       }
//                     } catch (e) {
//                       // 忽略检查错误
//                     }
//                   }
//                   clickInfo.detectedElements.push(elementInfo);
//                   break;

//                 case 'CDrawing':
//                 case 'CShape':
//                   elementInfo.details = '图形/图片元素';
//                   clickInfo.detectedElements.push(elementInfo);
//                   foundTargetElements++;
//                   if (range) {
//                     elementInfo.isClickTarget = true;
//                     clickInfo.elementType = 'image/shape';
//                   }
//                   break;

//                 default:
//                   if (elementType.includes('Drawing') || elementType.includes('Image')) {
//                     elementInfo.details = '图像元素';
//                     clickInfo.detectedElements.push(elementInfo);
//                     foundTargetElements++;
//                     if (range) {
//                       elementInfo.isClickTarget = true;
//                       clickInfo.elementType = 'image';
//                     }
//                   }
//                   break;
//               }
//             } catch (elementError) {
//               // 忽略元素读取错误
//             }
//           }

//           // 如果没有检测到特定元素类型，设为文档点击
//           if (clickInfo.elementType === 'unknown') {
//             if (clickInfo.hasSelection && clickInfo.selectionText) {
//               clickInfo.elementType = 'text-content';
//             } else {
//               clickInfo.elementType = 'document';
//             }
//           }

//           logger.info('📍 点击内容类型:', clickInfo.elementType);
//           logger.info('📄 选中文本:', clickInfo.selectionText);
//           logger.info('🔍 检测到的元素:', clickInfo.detectedElements.filter(e => e.isClickTarget));

//           return clickInfo;
//         } catch (error) {
//           logger.error('点击内容分析错误:', error);
//           return { elementType: 'error', error: error.message };
//         }
//       }, function(clickInfo) {
//         // 回调函数 - 处理分析结果
//         logger.info('✅ 点击内容分析完成:', clickInfo);

//         // 根据点击类型输出不同信息
//         switch(clickInfo.elementType) {
//           case 'table':
//             logger.info('🔲 用户点击了表格');
//             break;
//           case 'paragraph':
//             logger.info('📝 用户点击了段落文本');
//             break;
//           case 'image':
//           case 'image/shape':
//             logger.info('🖼️ 用户点击了图片/图形');
//             break;
//           case 'text-content':
//             logger.info('📄 用户选中了文本内容');
//             break;
//           case 'document':
//             logger.info('📋 用户点击了空白文档区域');
//             break;
//           default:
//             logger.info('❓ 用户点击了未知类型内容:', clickInfo.elementType);
//         }
//       });

//       // 获取当前内容控件属性
//       window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function(obj) {
//         logger.info('Current content control properties:', obj);

//         var hasContentControl = obj && obj.Tag;

//         if (hasContentControl) {
//           logger.info('🏷️ 检测到内容控件，Tag:', obj.Tag);

//           if (obj.Tag.startsWith('link-data:')) {
//             logger.info('🔗 Link content control clicked!');
//           } else if (obj.Tag.startsWith('table-binding:')) {
//             logger.info('🔲 Table binding content control clicked!');
//           } else if (obj.Tag.startsWith('paragraph-template:')) {
//             logger.info('📝 Paragraph template content control clicked!');
//           } else if (obj.Tag.startsWith('custom-binding:')) {
//             logger.info('⚙️ Custom binding content control clicked!');
//           } else {
//             logger.info('📋 Content control clicked with tag:', obj.Tag);
//           }

//           // 🔥 直接在这里构造内容控件点击数据并调用回调
//           logger.info('🚀 内容控件点击，直接触发处理...');

//           // 构造内容控件检测结果
//           var contentControlResult = {
//             tag: obj.Tag,
//             alias: obj.Alias || '',
//             internalId: obj.InternalId || '',
//             appearance: obj.Appearance || 0
//           };

//           // 直接调用回调，传递内容控件信息
//           if (cb) {
//             logger.info('📞 直接调用 cb 回调，传递内容控件数据...');
//             cb(contentControlResult);
//           }
//         } else {
//           logger.info('❌ 没有检测到内容控件');

//           // 执行原有的回调（用于非内容控件的点击）
//           cb?.();
//         }
//       });
//     };

//     // 使用 onSelectionChanged 事件（用于光标移动等）
//     window.Asc.plugin.event_onSelectionChanged = () => {
//       logger.info('Selection changed event triggered - running element detection');
//       cb?.();
//     };

//     // 如果有 connector API，也使用它
//     if (window.connector && typeof window.connector.attachEvent === 'function') {
//       logger.info('Using connector.attachEvent for enhanced selection tracking');

//       // 监听内容控件变化事件
//       window.connector.attachEvent("onChangeContentControl", (obj) => {
//         logger.info('Content control changed via connector:', obj);
//         cb?.();
//       });

//       // 监听目标位置变化（光标移动）
//       window.connector.attachEvent("onTargetPositionChanged", (obj) => {
//         logger.info('Target position changed via connector:', obj);
//         cb?.();
//       });
//     }

//     logger.info('✅ Selection change handlers configured');
//     logger.info('onClick handler set:', typeof window.Asc.plugin.event_onClick);
//     logger.info('Selection change handler set:', typeof window.Asc.plugin.event_onSelectionChanged);
//   }

//   onHyperLinkClick(cb) {
//     window.Asc = window.Asc || {};
//     window.Asc.plugin = window.Asc.plugin || {};
//     window.Asc.plugin.event_onHyperLinkClick = (data) => {
//       logger.info('hyperlink click detected:', data);
//       cb?.(data);
//     };
//   }
// }
