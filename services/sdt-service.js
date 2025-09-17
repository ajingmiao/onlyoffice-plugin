import { TAG_PREFIX } from '../core/constants.js';

export class SdtService {
  constructor(editorService) {
    this.editor = editorService;
  }

  /**
   * 插入一个行内内容控件（示例用组合框；你也可切换 DropDownList）
   * tagKey: 业务标识（会被加上 TAG_PREFIX）
   * title:  标题
   * placeholder: 占位显示文本
   */
  insertInlineSdt({ tagKey, title, placeholder }) {
    return this.editor.runInDoc(function () {
      var doc = Api.GetDocument();

      // 组合框控件（也可用 doc.AddDropDownListContentControl()）
      var sdt = doc.AddComboBoxContentControl();

      // 显示内容（清空默认内容后，加入占位文字）
      var content = sdt.GetContent();
      content.RemoveAllElements();
      var p = content.GetElement(0);
      p.AddText(placeholder || '{{placeholder}}');

      // Tag / Title（Tag 用于数据绑定）
      sdt.SetTag((TAG_PREFIX + tagKey) || 'bind:unknown');
      if (title) sdt.SetTitle(title);
    }, { async: false });
  }

  /** 在当前 Range/选择内探测是否命中某个 SDT，并返回其 tag/alias */
  detectActiveSdt() {
    console.log('🔍 detectActiveSdt 开始执行...');
    return new Promise((resolve) => {
      this.editor.runInDoc(function () {
        console.log('📄 进入 runInDoc 上下文...');

        try {
          var doc = Api.GetDocument();
          if (!doc) {
            console.log('❌ 无法获取文档对象');
            return null;
          }

          var range;
          try {
            range = doc.GetRangeBySelect();
            console.log('📋 获取选区成功:', range);
          } catch (rangeError) {
            console.log('❌ 获取选区失败 (可能是图表等特殊元素):', rangeError.message);
            // 对于图表等特殊元素，GetRangeBySelect 可能失败，这是正常的
            return null;
          }

          if (!range) {
            console.log('❌ 没有选区，返回 null');
            return null;
          }

          console.log('🔍 开始扫描内容控件...');
          var ctrls;
          try {
            ctrls = doc.GetAllContentControls();
            console.log('📊 找到', ctrls ? ctrls.length : 0, '个内容控件');
          } catch (ctrlError) {
            console.log('❌ 获取内容控件失败:', ctrlError.message);
            return null;
          }

          if (!ctrls || ctrls.length === 0) {
            console.log('❌ 文档中没有内容控件');
            return null;
          }

          // 简化的内容控件检测，避免使用有问题的API
          for (var i = 0; i < ctrls.length; i++) {
            var c = ctrls[i];
            console.log('🔍 检查内容控件', i);

            try {
              // 只尝试获取基本信息，不做复杂的位置检测
              var tag = '';
              var alias = '';

              if (typeof c.GetTag === 'function') {
                tag = c.GetTag() || '';
              }

              if (typeof c.GetAlias === 'function') {
                alias = c.GetAlias() || '';
              }

              // 如果有有效的标签，就认为可能是当前选中的
              if (tag) {
                console.log('✅ 找到有标签的内容控件:', { tag, alias });
                return { tag: tag, alias: alias };
              }
            } catch (controlError) {
              console.log('❌ 检查内容控件', i, '失败:', controlError.message);
              continue;
            }
          }

          console.log('❌ 没有找到匹配的内容控件');
          return null;

        } catch (error) {
          console.log('❌ detectActiveSdt 执行失败:', error.message);
          return null;
        }
      }, { async: false, cb: (res) => {
        console.log('🔄 runInDoc 回调被调用，结果:', res);
        resolve(res);
      } });
    });
  }
}
