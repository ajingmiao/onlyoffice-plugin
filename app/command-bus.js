import { COMMANDS } from '../core/constants.js';

export class CommandBus {
  constructor({ sdtService, linkService, wordArtService, shapeService, tableService, selectionBindingService, elementDetectionService, chartBindingService, clickTypeDetector }) {
    this.sdt = sdtService;
    this.link = linkService;
    this.wordart = wordArtService;
    this.shape = shapeService;
    this.table = tableService;
    this.selectionBinding = selectionBindingService;
    this.elementDetection = elementDetectionService;
    this.chartBinding = chartBindingService;
    this.clickTypeDetector = clickTypeDetector;
  }

  async dispatch(cmd) {
    console.log('[CommandBus] 收到 INSERT_LINK 指令，参数 =', cmd);
    const { command, data } = cmd || {};
    switch (command) {
       
      case COMMANDS.INSERT_TEXT: {
        // 兼容字符串或 {text:'...'}，将其视为 tagKey
        const payload = data;
        const text = (typeof payload === 'string') ? payload : (payload && payload.text) || '';
        const tagKey = text || 'customer_name';

        await this.sdt.insertInlineSdt({
          tagKey,
          title: '客户姓名',
          placeholder: `{{${tagKey}}}`,
        });
        return { ok: true };
      }

      case COMMANDS.REPORT_ACTIVE_SDT: {
        const info = await this.sdt.detectActiveSdt();
        return { ok: true, data: info };
      }


      case COMMANDS.INSERT_LINK: {
        // 允许 data 是字符串（显示文本），或对象 { text, url, color, bold, underline, json }
        const text = (typeof data === 'string') ? data : (data?.text || '点击这里');
        const url  = (typeof data === 'object' && data?.url !== undefined) ? data.url : ''; // 允许空
        const json = (typeof data === 'object' && data?.json !== undefined) ? data.json : null;

        await this.link.insertLinkInline({
          text,
          url,
          json,
          style: {
            bold: (typeof data === 'object' && data?.bold !== undefined) ? !!data.bold : true,
            underline: (typeof data === 'object' && data?.underline !== undefined) ? !!data.underline : true,
            // 标准蓝：#1976D2（亦可传入 data.color）
            color: (typeof data === 'object' && data?.color) || '#1976D2'
          }
        });
        return { ok: true };
      }

      case COMMANDS.LINK_CLICKED: {
        const linkData = await this.link.handleLinkClick();
        if (linkData) {
          // 发送链接数据回宿主页面
          return { ok: true, data: linkData };
        } else {
          return { ok: false, error: 'No link data found at current position' };
        }
      }

      case COMMANDS.INSERT_WORDART: {
        // 允许 data 包含 WordArt 的各种参数
        const result = await this.wordart.insertWordArt(data);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'WordArt insertion failed' };
        }
      }

      case COMMANDS.INSERT_PRESET_WORDART: {
        // 预设样式的 WordArt: { preset: 'classic|modern|fun', text: '...' }
        const preset = data?.preset || 'classic';
        const text = data?.text;
        const result = await this.wordart.insertPresetWordArt(preset, text);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'Preset WordArt insertion failed' };
        }
      }

      case COMMANDS.TEST_SHAPE_INLINE: {
        // 测试Shape内联插入可能性
        const result = await this.shape.testShapeInlineInsertion();
        return { ok: true, data: result };
      }

      case COMMANDS.INSERT_SHAPE_INLINE: {
        // 内联插入Shape
        const result = await this.shape.insertShapeInline(data);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'Shape inline insertion failed' };
        }
      }

      case COMMANDS.INSERT_SHAPE_PARAGRAPH: {
        // 段落级别插入Shape（可靠方法）
        const result = await this.shape.insertShapeInParagraph(data);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'Shape paragraph insertion failed' };
        }
      }

      case COMMANDS.INSERT_TABLE: {
        // 插入表格
        const result = await this.table.insertTable(data);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'Table insertion failed' };
        }
      }

      case COMMANDS.INSERT_PRESET_TABLE: {
        // 插入预设表格
        const preset = data?.preset || 'simple';
        const customData = data?.customData || {};
        const result = await this.table.insertPresetTable(preset, customData);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'Preset table insertion failed' };
        }
      }

      case COMMANDS.INSERT_DYNAMIC_TABLE: {
        // 插入动态数据表格（从宿主绑定）
        const result = await this.table.insertDynamicTable(data);
        if (result && result.success) {
          return { ok: true, data: result };
        } else {
          return { ok: false, error: result?.error || 'Dynamic table insertion failed' };
        }
      }

      case COMMANDS.TABLE_CLICKED: {
        const tableData = await this.table.handleTableClick();
        if (tableData) {
          // 发送表格点击数据回宿主页面
          return { ok: true, data: tableData };
        } else {
          return { ok: false, error: 'No table data found at current position' };
        }
      }

      case COMMANDS.ANALYZE_SELECTION: {
        const selectionData = await this.selectionBinding.analyzeSelection();
        if (selectionData && selectionData.success) {
          return { ok: true, data: selectionData };
        } else {
          return { ok: false, error: selectionData?.error || 'Selection analysis failed' };
        }
      }

      case COMMANDS.BIND_SELECTION: {
        const bindingResult = await this.selectionBinding.bindSelection(data);
        if (bindingResult && bindingResult.success) {
          return { ok: true, data: bindingResult };
        } else {
          return { ok: false, error: bindingResult?.error || 'Selection binding failed' };
        }
      }

      case COMMANDS.BINDING_CLICKED: {
        const bindingClickData = await this.selectionBinding.handleBindingClick();
        if (bindingClickData && bindingClickData.success) {
          return { ok: true, data: bindingClickData };
        } else {
          return { ok: false, error: bindingClickData?.error || 'No binding control clicked' };
        }
      }

      case COMMANDS.ELEMENT_CLICKED: {
        const elementData = await this.elementDetection.detectClickedElement();
        if (elementData && elementData.success) {
          return { ok: true, data: elementData };
        } else {
          return { ok: false, error: elementData?.error || 'Element detection failed' };
        }
      }

      case COMMANDS.PRECISE_TABLE_CELL_CLICKED: {
        const cellData = await this.elementDetection.detectTableCellClick();
        if (cellData && cellData.success) {
          return { ok: true, data: cellData };
        } else {
          return { ok: false, error: cellData?.error || 'Precise table cell detection failed' };
        }
      }

      case COMMANDS.BIND_CHART_DATA: {
        const chartResult = await this.chartBinding.bindDataToChart(data);
        // 清理临时数据
        this.chartBinding.cleanupTempData();
        if (chartResult && chartResult.success) {
          return { ok: true, data: chartResult };
        } else {
          return { ok: false, error: chartResult?.error || 'Chart data binding failed' };
        }
      }

      case COMMANDS.CHART_CLICKED: {
        const chartData = await this.chartBinding.detectChartClick();
        if (chartData && chartData.success) {
          return { ok: true, data: chartData };
        } else {
          return { ok: false, error: chartData?.error || 'Chart click detection failed' };
        }
      }

      case COMMANDS.GET_CHART_SUMMARY: {
        const summaryData = await this.chartBinding.getChartBindingSummary();
        if (summaryData && summaryData.success) {
          return { ok: true, data: summaryData };
        } else {
          return { ok: false, error: summaryData?.error || 'Chart summary retrieval failed' };
        }
      }
      case COMMANDS.DETECT_CLICK_TYPE: {
        const clickTypeResult = await this.clickTypeDetector.detectClickType();
        if (clickTypeResult) {
          return { ok: true, data: clickTypeResult };
        } else {
          return { ok: false, error: 'Click type detection failed' };
        }
      }

       case COMMANDS.UPDATE_BINDING: {
        const updateResult = await this.selectionBinding.updateBinding(data);
        if (updateResult && updateResult.success) {
          return { ok: true, data: updateResult };
        } else {
          return { ok: false, error: updateResult?.error || 'Binding update failed' };
        }
      }

      case COMMANDS.UPDATE_CHART_BINDING: {
        const chartUpdateResult = await this.chartBinding.updateChartBinding(data);
        if (chartUpdateResult && chartUpdateResult.success) {
          return { ok: true, data: chartUpdateResult };
        } else {
          return { ok: false, error: chartUpdateResult?.error || 'Chart binding update failed' };
        }
      }

      case COMMANDS.UPDATE_ROW_BINDING: {
        const rowUpdateResult = await this.selectionBinding.updateRowBinding(data);
        if (rowUpdateResult && rowUpdateResult.success) {
          return { ok: true, data: rowUpdateResult };
        } else {
          return { ok: false, error: rowUpdateResult?.error || 'Row binding update failed' };
        }
      }


      default:
        return { ok: false, error: 'Unknown command' };
    }
  }
}

     