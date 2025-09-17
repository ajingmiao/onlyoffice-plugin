import { HostBridge } from '../bridge/host-bridge.js';
import { PluginBridge } from '../bridge/plugin-bridge.js';
import { EditorService } from '../services/editor-service.js';
import { SdtService } from '../services/sdt-service.js';
import { SelectionService } from '../services/selection-service.js';
import { LinkService } from '../services/link-service.js'
import { WordArtService } from '../services/wordart-service.js'
import { ShapeService } from '../services/shape-service.js'
import { TableService } from '../services/table-service.js'
import { SelectionBindingService } from '../services/selection-binding-service.js'
import { ElementDetectionService } from '../services/element-detection-service.js'
import { ChartBindingService } from '../services/chart-binding-service.js'
import { CommandBus } from './command-bus.js';
import { EventBus } from './event-bus.js';

export function createContainer() {
  const host = new HostBridge();
  const plugin = new PluginBridge();

  const editor = new EditorService(plugin);
  const sdt = new SdtService(editor);
  const link = new LinkService(editor)
  const wordart = new WordArtService(editor)
  const shape = new ShapeService(editor)
  const table = new TableService(editor)
  const selectionBinding = new SelectionBindingService(editor)
  const elementDetection = new ElementDetectionService(editor)
  const chartBinding = new ChartBindingService(editor)
  const selection = new SelectionService(plugin, sdt);

  const events = new EventBus({ hostBridge: host });
  const commands = new CommandBus({
    sdtService: sdt,
    linkService: link,
    wordArtService: wordart,
    shapeService: shape,
    tableService: table,
    selectionBindingService: selectionBinding,
    elementDetectionService: elementDetection,
    chartBindingService: chartBinding
  });

  return {
    host, plugin, editor, sdt, link, wordart, shape, table, selectionBinding, elementDetection, chartBinding,
    selection, events, commands
  };
}
