# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an OnlyOffice plugin built with a clean architecture that implements document manipulation, chart detection, and data binding capabilities. The plugin follows a modular design with dependency injection and command/event bus patterns.

## Architecture

The codebase is organized using a layered architecture with clear separation of concerns:

### Core Layers
- **Entry Point**: `plugin.js` - Minimal bootstrap that creates container and starts lifecycle
- **Application Layer**: `app/` - Command bus, event bus, dependency container, and lifecycle management
- **Bridge Layer**: `bridge/` - Abstraction for OnlyOffice API and host communication
- **Service Layer**: `services/` - Business logic for specific document operations
- **Core Utilities**: `core/` - Shared utilities, logging, and constants

### Key Architectural Patterns
- **Dependency Injection**: `app/container.js` assembles all services and dependencies
- **Command Bus**: `app/command-bus.js` routes commands to appropriate services
- **Event Bus**: `app/event-bus.js` handles plugin events and host communication
- **Bridge Pattern**: `bridge/plugin-bridge.js` wraps OnlyOffice API, `bridge/host-bridge.js` handles host communication

## Development Commands

This project uses vanilla JavaScript modules and has no build system. Development is done directly with the source files.

### Testing Code Syntax
```bash
node --check plugin.js
node --check app/container.js
node --check app/command-bus.js
node --check app/lifecycle.js
node --check bridge/plugin-bridge.js
```

### OnlyOffice API Availability Check
```bash
curl -I http://localhost:9998/web-apps/apps/api/documents/api.js
```

## Key Services

### Document Manipulation Services
- `editor-service.js`: Core document operations (GetDocument, Range operations)
- `sdt-service.js`: Content controls (Structured Document Tags) management
- `selection-service.js`: Selection change detection and information extraction
- `link-service.js`: Hyperlink insertion with JSON data binding
- `wordart-service.js`: WordArt insertion and styling
- `shape-service.js`: Shape insertion (inline and paragraph level)
- `table-service.js`: Table creation, insertion, and interaction handling

### Detection and Binding Services
- `element-detection-service.js`: General element click detection
- `chart-binding-service.js`: Chart data binding and detection
- `chart-data-binder.js`: Chart data binding logic
- `chart-detector.js`: Chart detection utilities
- `chart-type-detector.js`: Chart type identification
- `selection-binding-service.js`: Selection-based data binding

## Command System

Commands are defined in `core/constants.js` and handled by the command bus. Key commands include:

### Document Operations
- `INSERT_TEXT`: Insert inline content control with placeholder text
- `INSERT_LINK`: Insert hyperlink with optional JSON data binding
- `INSERT_WORDART`: Insert WordArt with custom styling
- `INSERT_TABLE`: Insert tables with various presets
- `INSERT_SHAPE_INLINE/INSERT_SHAPE_PARAGRAPH`: Shape insertion

### Detection and Binding
- `LINK_CLICKED`: Detect and extract JSON data from clicked links
- `TABLE_CLICKED`: Handle table cell clicks
- `CHART_CLICKED`: Chart click detection and data extraction
- `ELEMENT_CLICKED`: General element detection
- `BIND_SELECTION`: Bind data to selected content

## Event Handling

The plugin listens to several OnlyOffice events configured in `config.json`:
- `onSelectionChanged`: Triggers element detection and data binding
- `onDocumentContentReady`: Plugin initialization
- `onClick`: Content control and element clicks
- `onHyperLinkClick`: Hyperlink interaction handling

## Host Communication

The host page (`E:\code\onlyoffice-plugin-demo\src\App.vue`) communicates with the plugin through:

### From Host to Plugin
- `serviceCommandSafe(command, data)`: Sends commands to plugin with queue management
- Commands include: `insertText`, `insertLink`, `bindChartData`, `getChartSummary`, etc.

### From Plugin to Host
- `host.sendInfo(eventType, data)`: Send data to host
- `host.onInternalCommand(callback)`: Receive commands from host
- Events: `linkClicked`, `tableClicked`, `chartClicked`, `elementClicked`, etc.

### Chart Binding Integration
The `bindChartData` function in the host generates sample chart data with:
- Chart metadata (title, type, data source, period)
- Business metrics (sales, growth rate, targets)
- Data series for visualization
- Binding metadata (ID, timestamps, permissions)

The plugin detects chart clicks and returns bound data to the host for processing.

## Adding New Features

1. **New Command**: Add constant to `core/constants.js`
2. **New Service**: Create `services/your-service.js` with editor dependency
3. **Container Registration**: Add service to `app/container.js`
4. **Command Routing**: Add case to `app/command-bus.js`
5. **Host Integration**: Update host page to handle new events

## File Naming Conventions

- Services: `*-service.js`
- Bridges: `*-bridge.js`
- Core utilities: descriptive names in `core/`
- Application components: descriptive names in `app/`

## Code Style

- ES6 modules with explicit imports/exports
- Class-based services with constructor injection
- Async/await for OnlyOffice API calls
- Comprehensive logging via `core/logger.js`
- Error handling with try-catch and result objects
- Consistent return format: `{ ok: boolean, data?, error? }`