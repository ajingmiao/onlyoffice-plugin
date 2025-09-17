// services/click-type-detector.js
export class ClickTypeDetector {
  constructor(editorService) {
    this.editor = editorService;
  }

  /**
   * 检测当前点击/选择的类型
   * @returns {Promise<Object>} 点击类型检测结果
   */
  async detectClickType() {
    console.log('🎯 开始点击类型检测...');

    // 使用动态函数生成，避免作用域问题
    const funcStr = `
      const doc = Api.GetDocument();
      console.log('=== 点击类型检测（智能模式）===');

      try {
        const result = { type: 'unknown', data: null, confidence: 0 };

        // 1. 优先检查绘图对象（图表、图片、形状）
        console.log('🔍 检查绘图对象...');
        try {
          const drawings = doc.GetAllDrawingObjects();
          console.log('📄 找到绘图对象数量:', drawings ? drawings.length : 0);

          if (drawings && drawings.length > 0) {
            // 检查最近的绘图对象（通常是被点击的）
            for (let i = drawings.length - 1; i >= 0; i--) {
              const obj = drawings[i];
              console.log('🔍 检查绘图对象 ' + i);

              let objType = 'unknown';
              try {
                if (typeof obj.GetClassType === 'function') {
                  objType = obj.GetClassType();
                }
              } catch (typeError) {
                console.log('🚨 获取对象类型失败:', typeError.message);
                continue;
              }

              console.log('📊 绘图对象类型:', objType);

              // 根据对象类型返回相应结果
              if (objType === 'chart') {
                console.log('✅ 检测到图表点击');
                return {
                  type: 'chart',
                  data: {
                    index: i,
                    elementType: objType,
                    drawingIndex: i,
                    isDocumentLevel: true
                  },
                  confidence: 0.9
                };
              } else if (objType === 'shape') {
                console.log('✅ 检测到形状点击');
                return {
                  type: 'shape',
                  data: {
                    index: i,
                    elementType: objType,
                    drawingIndex: i
                  },
                  confidence: 0.8
                };
              } else if (objType === 'image' || objType.includes('Image')) {
                console.log('✅ 检测到图片点击');
                return {
                  type: 'image',
                  data: {
                    index: i,
                    elementType: objType,
                    drawingIndex: i
                  },
                  confidence: 0.8
                };
              } else if (objType.includes('Drawing') || objType.includes('Ole')) {
                console.log('✅ 检测到其他绘图对象');
                return {
                  type: 'drawing',
                  data: {
                    index: i,
                    elementType: objType,
                    drawingIndex: i
                  },
                  confidence: 0.7
                };
              }
            }
          }
        } catch (drawingError) {
          console.log('🚨 绘图对象检测失败:', drawingError.message);
        }

        // 2. 检查文本选择
        console.log('🔍 检查文本选择...');
        try {
          const range = doc.GetRangeBySelect();
          if (range) {
            console.log('✅ 检测到文本选择');

            // 进一步检查是否在特殊元素中
            try {
              // 检查是否在内容控件中
              const ctrls = doc.GetAllContentControls();
              if (ctrls && ctrls.length > 0) {
                for (let j = 0; j < ctrls.length; j++) {
                  const ctrl = ctrls[j];
                  try {
                    if (typeof range.IsInContentControl === 'function' &&
                        range.IsInContentControl(ctrl)) {
                      console.log('✅ 检测到内容控件选择');
                      return {
                        type: 'content-control',
                        data: {
                          controlIndex: j,
                          hasRange: true,
                          tag: ctrl.GetTag ? ctrl.GetTag() : '',
                          alias: ctrl.GetAlias ? ctrl.GetAlias() : ''
                        },
                        confidence: 0.9
                      };
                    }
                  } catch (ctrlError) {
                    // 忽略单个控件检查失败
                  }
                }
              }

              // 检查是否在表格中（简化检查）
              // 这里可以添加表格检测逻辑
              console.log('📋 基础文本选择');
              return {
                type: 'text',
                data: {
                  hasRange: true,
                  selectionType: 'text'
                },
                confidence: 0.8
              };

            } catch (detailError) {
              console.log('🚨 详细文本检测失败:', detailError.message);
              return {
                type: 'text',
                data: { hasRange: true },
                confidence: 0.6
              };
            }
          } else {
            console.log('⚠️ 没有文本选择');
          }
        } catch (rangeError) {
          console.log('⚠️ 文本选择检测失败 (可能是特殊元素):', rangeError.message);
          // 这通常意味着点击的是非文本元素
        }

        // 3. 如果都没检测到，返回未知类型
        console.log('❓ 未能确定点击类型');
        return {
          type: 'unknown',
          data: {
            reason: 'no-clear-selection'
          },
          confidence: 0.1
        };

      } catch (error) {
        console.log('❌ 点击类型检测失败:', error.message);
        return {
          type: 'error',
          data: {
            error: error.message
          },
          confidence: 0
        };
      }
    `;

    // 使用new Function执行检测代码
    const dynamicFunction = new Function(funcStr);
    const result = await this.editor.runInDoc(dynamicFunction);

    console.log('🎯 点击类型检测结果:', result);
    return result;
  }

  /**
   * 检测是否应该跳过某种类型的检测
   * @param {string} detectionType - 检测类型 ('sdt', 'chart', 'table', etc.)
   * @param {Object} clickTypeResult - 点击类型检测结果
   * @returns {boolean} 是否应该跳过
   */
  shouldSkipDetection(detectionType, clickTypeResult) {
    if (!clickTypeResult || !clickTypeResult.type) {
      return false; // 如果无法确定类型，不跳过
    }

    const clickType = clickTypeResult.type;

    switch (detectionType) {
      case 'sdt':
        // 如果点击的是图表、图片、形状，跳过SDT检测
        return ['chart', 'image', 'shape', 'drawing'].includes(clickType);

      case 'chart':
        // 如果点击的不是图表，跳过图表检测
        return clickType !== 'chart';

      case 'table':
        // 如果点击的是图表、图片，跳过表格检测
        return ['chart', 'image', 'shape'].includes(clickType);

      default:
        return false;
    }
  }
}