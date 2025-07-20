# autocad-mcp-server
# AutoCAD MCP Server Documentation

## Overview

The AutoCAD MCP (Model Context Protocol) Server is a comprehensive Python application that provides programmatic control over AutoCAD through COM automation. It enables creation, manipulation, and management of AutoCAD entities with intelligent layer management, structured building elements, and complete CRUD operations.

## Features

### Core Capabilities
- **AutoCAD COM Integration**: Direct connection to AutoCAD instances via Windows COM
- **Intelligent Layer Management**: Automatic layer creation and assignment based on structure types
- **Structured Building Elements**: Specialized creation tools for walls, doors, windows, rooms, and more
- **Complete Entity Management**: Create, read, update, and delete operations for all AutoCAD entities
- **Color and Thickness Support**: Advanced styling options for all created entities
- **Entity Querying**: Comprehensive entity retrieval with filtering and grouping options

### Supported Entity Types
- Lines (with thickness support)
- Circles (with concentric thickness)
- Rectangles
- Arcs (with parallel thickness)
- Text annotations
- Complex structures (walls, doors, windows, rooms)

## Requirements

### System Requirements
- **Operating System**: Windows (COM automation requirement)
- **AutoCAD**: Any version with COM support (2000+)
- **Python**: 3.7 or higher

### Dependencies
```txt
mcp>=1.0.0
pywin32>=306
```

## Installation

### Method 1: Using uv (Recommended)
```bash
# Install uv
pip install uv

# Create project directory
mkdir autocad-mcp-server
cd autocad-mcp-server

# Create virtual environment
uv venv

# Activate environment (Windows)
.venv\Scripts\activate

# Install dependencies
uv pip install mcp pywin32
```

### Method 2: Using pip
```bash
# Create virtual environment
python -m venv venv

# Activate environment (Windows)
venv\Scripts\activate

# Install dependencies
pip install mcp pywin32
```

## Usage

### Starting the Server
```bash
python autocad_mcp_server.py
```

The server will automatically:
1. Attempt to connect to an existing AutoCAD instance
2. If none exists, launch a new AutoCAD instance
3. Create standard building layers
4. Initialize the MCP server for client connections

### Connection Behavior
- **Existing AutoCAD**: Connects to active AutoCAD application
- **No AutoCAD**: Launches new AutoCAD instance
- **Auto-Recovery**: Automatically reconnects if connection is lost

## Architecture

### Class Structure

#### AutoCADCOMServer
Main server class that handles:
- COM connection management
- MCP tool registration
- Entity creation and manipulation
- Layer management
- Error handling and recovery

### Key Components

#### Layer Management System
```python
layer_definitions = {
    "walls": {"name": "WALLS", "color": "white", "description": "Building walls and partitions"},
    "doors": {"name": "DOORS", "color": "green", "description": "Door openings and frames"},
    "windows": {"name": "WINDOWS", "color": "cyan", "description": "Window openings and frames"},
    # ... more layers
}
```

#### Color Management
Supports both named colors and AutoCAD Color Index (ACI):
```python
color_map = {
    "red": 1, "yellow": 2, "green": 3, "cyan": 4,
    "blue": 5, "magenta": 6, "white": 7, "black": 0,
    "bylayer": 256, "byblock": 0
}
```

## API Reference

### Core Information Tools

#### `get_drawing_info()`
**Description**: Retrieves comprehensive information about the current AutoCAD drawing.

**Returns**:
```json
{
  "filename": "drawing.dwg",
  "path": "C:\\Projects\\drawing.dwg",
  "entity_count": 150,
  "saved": true,
  "current_layer": "0",
  "total_layers": 8,
  "available_layers": [
    {
      "name": "0",
      "color": 7,
      "frozen": false,
      "locked": false
    }
  ]
}
```

#### `get_entities(max_entities?)`
**Description**: Retrieves all entities in the drawing with detailed properties and layer grouping.

**Parameters**:
- `max_entities` (optional): Maximum number of entities to retrieve

**Returns**:
```json
{
  "entities": [...],
  "entities_by_layer": {
    "WALLS": [...],
    "DOORS": [...]
  },
  "total_count": 150,
  "returned_count": 150,
  "layer_summary": {
    "WALLS": 25,
    "DOORS": 8
  }
}
```

### Structure Creation Tools

#### `create_structure(structure_type, geometry_data, options)`
**Description**: Creates structured building elements with automatic layer assignment.

**Parameters**:
- `structure_type`: "wall", "door", "window", "room", "furniture", etc.
- `geometry_data`: Geometry definition object
- `color` (optional): Color override
- `thickness` (optional): Structure thickness
- `custom_layer` (optional): Custom layer name
- `label` (optional): Text label for the structure

**Example - Wall Creation**:
```json
{
  "structure_type": "wall",
  "geometry_data": {
    "start": [0, 0],
    "end": [10, 0]
  },
  "thickness": 0.2,
  "label": "Exterior Wall North"
}
```

**Example - Room Creation**:
```json
{
  "structure_type": "room",
  "geometry_data": {
    "corner1": [0, 0],
    "corner2": [12, 10]
  },
  "label": "Living Room",
  "custom_layer": "ROOMS"
}
```

### Basic Entity Creation

#### `create_line(start, end, color?, thickness?)`
**Description**: Creates a line with optional thickness (parallel lines with end caps).

**Parameters**:
- `start`: Start point [x, y] or [x, y, z]
- `end`: End point [x, y] or [x, y, z]
- `color`: Color name or ACI number (default: "white")
- `thickness`: Line thickness creates parallel lines (default: 0.0)

#### `create_circle(center, radius, color?, thickness?)`
**Description**: Creates a circle with optional thickness (concentric circles).

**Parameters**:
- `center`: Center point [x, y] or [x, y, z]
- `radius`: Circle radius
- `color`: Color name or ACI number (default: "white")
- `thickness`: Creates concentric circles (default: 0.0)

#### `create_rectangle(corner1, corner2, color?, thickness?)`
**Description**: Creates a rectangle from two opposite corners.

#### `create_text(position, text, height?, color?)`
**Description**: Creates text annotation.

#### `create_arc(center, radius, start_angle, end_angle, color?, thickness?)`
**Description**: Creates an arc with optional thickness.

### Layer Management

#### `create_or_get_layer(layer_name, color?, description?)`
**Description**: Creates a new layer or retrieves existing layer.

#### `set_current_layer(layer_name)`
**Description**: Sets the active drawing layer.

### Entity Deletion Tools

#### `delete_entity_by_handle(handle)`
**Description**: Deletes a specific entity by its unique handle.

#### `delete_entities_by_handles(handles[])`
**Description**: Deletes multiple entities by their handles.

#### `delete_entities_by_type(entity_type)`
**Description**: Deletes all entities of a specific type.
- Types: "AcDbLine", "AcDbCircle", "AcDbText", "AcDbArc"

#### `delete_entities_by_layer(layer_name)`
**Description**: Deletes all entities on a specific layer.

#### `delete_entities_by_color(color)`
**Description**: Deletes all entities with a specific color.

#### `delete_entities_by_type_and_color(entity_type, color)`
**Description**: Deletes entities matching both type and color criteria.

#### `delete_last_entities(count?)`
**Description**: Deletes the most recently created entities.

#### `delete_all_entities(confirm)`
**Description**: Deletes all entities in the drawing (requires confirmation).

### Utility Tools

#### `undo_last_operation()`
**Description**: Performs undo operation (equivalent to Ctrl+Z).

#### `change_entity_color(handle, color)`
**Description**: Changes the color of an existing entity.

#### `zoom_extents()`
**Description**: Zooms to show all objects in the drawing.

## Layer System

### Predefined Structure Layers

| Layer Name | Color | Purpose |
|------------|-------|---------|
| WALLS | White | Building walls and partitions |
| DOORS | Green | Door openings and frames |
| WINDOWS | Cyan | Window openings and frames |
| FURNITURE | Yellow | Furniture and fixtures |
| ELECTRICAL | Red | Electrical fixtures and outlets |
| PLUMBING | Blue | Plumbing fixtures and pipes |
| HVAC | Magenta | HVAC ducts and equipment |
| STRUCTURE | Gray | Structural elements |
| ANNOTATION | White | Text and dimensions |
| SITE | Green | Site elements and landscaping |
| UTILITIES | Red | Utility lines and equipment |

### Automatic Layer Assignment

The system automatically assigns appropriate layers based on structure type:

```python
structure_mapping = {
    "wall" → "WALLS",
    "door" → "DOORS", 
    "window" → "WINDOWS",
    "furniture", "chair", "table" → "FURNITURE",
    "outlet", "switch", "light" → "ELECTRICAL",
    # ... etc
}
```

## Color System

### Named Colors
- `red`, `yellow`, `green`, `cyan`, `blue`, `magenta`
- `white`, `black`, `gray`, `light_gray`
- `bylayer`, `byblock`

### ACI Numbers
Direct AutoCAD Color Index numbers (0-255) are also supported.

## Error Handling

### Connection Recovery
- Automatic reconnection on connection loss
- Graceful handling of AutoCAD crashes
- Fallback to new AutoCAD instance creation

### Operation Safety
- Safe operation wrappers with timing delays
- Exception handling for individual entity creation
- Partial success reporting for batch operations

### Error Response Format
```json
{
  "error": "Description of the error",
  "context": "Additional context information"
}
```

## Examples

### Creating a Simple Floor Plan

```python
# Create exterior walls
create_structure("wall", {
    "start": [0, 0], "end": [20, 0]
}, thickness=0.2, label="South Wall")

create_structure("wall", {
    "start": [20, 0], "end": [20, 15]
}, thickness=0.2, label="East Wall")

# Add a door
create_structure("door", {
    "start": [8, 0], "end": [11, 0],
    "width": 3
}, label="Main Entrance")

# Add windows
create_structure("window", {
    "start": [15, 0], "end": [18, 0]
}, label="Living Room Window")
```

### Querying and Managing Entities

```python
# Get all entities
entities = get_entities()

# Get entities by layer
wall_entities = entities["entities_by_layer"]["WALLS"]

# Delete specific entities
delete_entities_by_layer("FURNITURE")

# Change colors
for entity in wall_entities:
    change_entity_color(entity["handle"], "yellow")
```

## Troubleshooting

### Common Issues

#### AutoCAD Not Found
**Problem**: "Not connected to AutoCAD" error
**Solution**: 
1. Ensure AutoCAD is installed
2. Run as Administrator if needed
3. Check Windows COM registration

#### COM Timeout Errors
**Problem**: Operations timing out
**Solution**: 
- Increase delay parameters in `safe_operation`
- Process entities in smaller batches
- Ensure AutoCAD is responsive

#### Layer Creation Failures
**Problem**: Layers not being created
**Solution**:
- Check layer name validity (no special characters)
- Ensure drawing is not read-only
- Verify sufficient permissions

### Performance Optimization

#### Large Drawings
- Use `max_entities` parameter for large drawings
- Process entities in batches of 100-500
- Consider using specific layer queries instead of full entity retrieval

#### Memory Management
- The server automatically handles COM object cleanup
- Use entity handles for references instead of keeping COM objects
- Batch similar operations together

## Development

### Extending the Server

#### Adding New Structure Types
1. Add to `layer_definitions` dictionary
2. Implement specific creation method
3. Add to structure type mapping
4. Update tool schema

#### Custom Entity Properties
1. Extend entity creation methods
2. Add property setting in `safe_operation` wrapper
3. Update return schemas

### Testing
- Test with various AutoCAD versions
- Verify COM object lifecycle management
- Test error recovery scenarios
- Validate layer creation and assignment

## License

This project uses the Model Context Protocol (MCP) framework and requires appropriate licensing for AutoCAD COM automation usage.
