export const COMMANDS = {
    INSERT_TEXT: 'insertText',
    REPORT_ACTIVE_SDT: 'reportActiveSdt',
    INSERT_LINK: 'insertLink',
    LINK_CLICKED: 'linkClicked',
    INSERT_WORDART: 'insertWordArt',
    INSERT_PRESET_WORDART: 'insertPresetWordArt',
    TEST_SHAPE_INLINE: 'testShapeInline',
    INSERT_SHAPE_PARAGRAPH: 'insertShapeParagraph',
    INSERT_TABLE: 'insertTable',
    INSERT_PRESET_TABLE: 'insertPresetTable',
    INSERT_DYNAMIC_TABLE: 'insertDynamicTable',
    TABLE_CLICKED: 'tableClicked',
    BIND_SELECTION: 'bindSelection',
    ANALYZE_SELECTION: 'analyzeSelection',
    BINDING_CLICKED: 'bindingClicked',
    ELEMENT_CLICKED: 'elementClicked',
    PRECISE_TABLE_CELL_CLICKED: 'preciseTableCellClicked',
    BIND_CHART_DATA: 'bindChartData',
    CHART_CLICKED: 'chartClicked',
    GET_CHART_SUMMARY: 'getChartSummary',
    UPDATE_BINDING: 'updateBinding',
    UPDATE_CHART_BINDING: 'updateChartBinding',
    UPDATE_ROW_BINDING: 'updateRowBinding'
  };
  
  export const EVENTS = {
    PLUGIN_INITIALIZED: 'pluginInitialized',
    ACTIVE_SDT: 'activeSdt',
    SELECTION_CHANGED_FIRED: 'onSelectionChanged-fired'
  };
  
  export const TAG_PREFIX = 'bind:'; // 统一为数据绑定控件加前缀
  