export class SelectionService {
    constructor(pluginBridge, sdtService) {
      this.plugin = pluginBridge;
      this.sdt = sdtService;
      console.log('🔧 SelectionService 构造完成');
    }

    bindSelectionChange(handler) {
      console.log('🔗 bindSelectionChange 被调用，handler 类型:', typeof handler);
      this.plugin.onSelectionChanged(async (contentControlData) => {
        console.log('📡 onSelectionChanged 回调被触发');
        console.log('📦 收到的数据:', contentControlData);

        try {
          var result;

          if (contentControlData && contentControlData.tag) {
            // 如果传递了内容控件数据，直接使用
            console.log('✅ 收到内容控件数据，直接使用');
            result = {
              tag: contentControlData.tag,
              alias: contentControlData.alias || '',
              internalId: contentControlData.internalId || '',
              appearance: contentControlData.appearance || 0,
              type: 'content-control'
            };
          } else {
            // 否则尝试检测活动的 SDT
            console.log('🔍 没有内容控件数据，尝试 detectActiveSdt...');
            result = await this.sdt.detectActiveSdt();
            console.log('🔍 detectActiveSdt 返回结果:', result);
            if (result) {
              result.type = 'sdt-detected';
              console.log('✅ detectActiveSdt 检测到内容控件');
            } else {
              console.log('⚠️ detectActiveSdt 未检测到内容控件，可能是图表或其他元素');
              result = null;
            }
          }

          console.log('🔍 最终检测结果:', result);
          console.log('📞 即将调用 handler，handler 类型:', typeof handler);

          if (handler) {
            handler(result);
            console.log('✅ handler 调用完成');
          } else {
            console.warn('⚠️ handler 为空');
          }
        } catch (error) {
          console.error('❌ bindSelectionChange 回调出错:', error);
        }
      });
      console.log('🔗 bindSelectionChange 设置完成');
    }
  }
  