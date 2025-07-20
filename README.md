# AutoCAD MCP Server - Complete Documentation

🏗️ **A powerful Python-based Model Context Protocol (MCP) server that provides programmatic control over AutoCAD through COM automation.**

Transform your AutoCAD workflow with intelligent building structure creation, automatic layer management, and comprehensive entity manipulation - all through Claude and other MCP clients.

## 🚀 Quick Start (30 seconds)

```bash
# 1. Clone and setup
git clone https://github.com/yourusername/autocad-mcp-server
cd autocad-mcp-server
pip install -e .

# 2. Start AutoCAD (must be running)
# Open AutoCAD application

# 3. Test the server
python test_script.py
```

## 📋 Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration with Claude Desktop](#configuration-with-claude-desktop)
- [Running the Server](#running-the-server)
- [Basic Usage Examples](#basic-usage-examples)
- [Advanced Features](#advanced-features)
- [API Reference](#api-reference)
- [Troubleshooting](#troubleshooting)
- [Development](#development)

## Prerequisites

### System Requirements
- **Operating System**: Windows (COM automation requirement)
- **AutoCAD**: Any version 2000+ with COM support
- **Python**: 3.8 or higher
- **Claude Desktop**: Latest version (for MCP integration)

### Required Software
1. **AutoCAD Installation**
   - Any AutoCAD version (AutoCAD, AutoCAD LT, Civil 3D, etc.)
   - Must have valid license and be able to run
   - COM automation must be enabled (default in most installations)

2. **Python Environment**
   - Python 3.8+ installed
   - Administrative privileges may be required for COM registration

## Installation

### Method 1: Using pip (Recommended)

```bash
# Clone the repository
git clone https://github.com/yourusername/autocad-mcp-server
cd autocad-mcp-server

# Create virtual environment
python -m venv venv

# Activate environment (Windows)
venv\Scripts\activate

# Install the package
pip install -e .
```

### Method 2: Using uv (Faster)

```bash
# Install uv if not already installed
pip install uv

# Clone and setup
git clone https://github.com/yourusername/autocad-mcp-server
cd autocad-mcp-server

# Create environment and install
uv venv
uv pip install -e .
```

### Dependencies Installed
- `mcp>=1.0.0` - Model Context Protocol framework
- `pywin32>=306` - Windows COM automation

## Configuration with Claude Desktop

### Step 1: Locate Claude Desktop Config

Find your Claude Desktop configuration file:
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`

### Step 2: Add MCP Server Configuration

Edit the config file to include your AutoCAD MCP server:

```json
{
  "mcpServers": {
    "autocad-mcp": {
      "command": "python",
      "args": [
        "C:\\path\\to\\your\\autocad-mcp-server\\test_script.py"
      ],
      "env": {
        "PYTHONPATH": "C:\\path\\to\\your\\autocad-mcp-server"
      }
    }
  }
}
```

**Important**: Replace `C:\\path\\to\\your\\autocad-mcp-server` with your actual project path.

### Step 3: Alternative Configuration (If Installed via pip)

If you installed the package globally:

```json
{
  "mcpServers": {
    "autocad-mcp": {
      "command": "autocad-mcp",
      "args": []
    }
  }
}
```

### Step 4: Restart Claude Desktop

1. Close Claude Desktop completely
2. Restart the application
3. Verify the MCP server appears in available tools

## Running the Server

### Method 1: Standalone Testing

```bash
# Activate your environment
venv\Scripts\activate  # Windows
# source venv/bin/activate  # macOS/Linux

# Start AutoCAD first (important!)
# Open AutoCAD application

# Run the test script
python test_script.py
```

**Expected Output:**
```
✅ Successfully connected to AutoCAD
📄 Drawing: Drawing1.dwg
📊 Entities: 0
🚀 MCP Server starting...
Server running and waiting for client connections...
```

### Method 2: Through Claude Desktop

1. **Start AutoCAD** (must be running before Claude connects)
2. **Open Claude Desktop**
3. **Start new conversation**
4. **Verify MCP connection** - You should see "autocad-mcp" in available tools
5. **Test with a simple command**:

```
Hi! Can you check what AutoCAD drawing is currently open?
```

### Method 3: Direct MCP Client

For advanced users or custom integrations:

```python
import asyncio
from autocad_mcp import AutoCADCOMServer

async def run_server():
    server = AutoCADCOMServer()
    await server.main()

asyncio.run(run_server())
```

## Basic Usage Examples

### Through Claude Desktop

Once configured, you can use natural language with Claude to control AutoCAD:

#### Example 1: Check Drawing Status
```
Claude: Can you tell me about the current AutoCAD drawing?
```

**Claude will use the MCP server to get drawing information and respond with details about the current drawing, layers, and entity count.**

#### Example 2: Create a Simple Room
```
Claude: Create a 12x10 foot room with walls that are 6 inches thick. Add a door on the south wall and a window on the east wall.
```

**Claude will:**
- Create a room structure using `create_structure`
- Add walls with proper thickness
- Place door and window openings
- Use appropriate layers (WALLS, DOORS, WINDOWS)
- Label the elements

#### Example 3: Layer Management
```
Claude: Show me all the layers in the current drawing and create a new layer called "HVAC" with red color.
```

#### Example 4: Entity Querying
```
Claude: List all the entities in the drawing grouped by layer. How many walls do we have?
```

### Direct API Usage (Advanced)

If you're building custom applications:

```python
from autocad_mcp import AutoCADCOMServer

# Initialize server
server = AutoCADCOMServer()

# Connect to AutoCAD
if server.connect_to_autocad():
    
    # Create a wall
    result = server.create_structure(
        structure_type="wall",
        geometry_data={
            "start": [0, 0],
            "end": [20, 0]
        },
        thickness=0.5,
        label="North Wall"
    )
    
    # Create a door
    door_result = server.create_structure(
        structure_type="door", 
        geometry_data={
            "start": [8, 0],
            "end": [11, 0],
            "width": 3
        },
        label="Main Entrance"
    )
    
    # Get all entities
    entities = server.get_entities()
    print(f"Total entities: {entities['total_count']}")
```

## Advanced Features

### Intelligent Layer Management

The server automatically assigns structures to appropriate layers:

```python
# Automatically goes to WALLS layer
create_structure("wall", geometry_data)

# Automatically goes to ELECTRICAL layer  
create_structure("outlet", geometry_data)

# Custom layer assignment
create_structure("furniture", geometry_data, custom_layer="BEDROOM_FURNITURE")
```

### Thickness and Styling

```python
# Wall with thickness (creates parallel lines with end caps)
create_structure("wall", {
    "start": [0, 0], 
    "end": [20, 0]
}, thickness=0.5, color="white")

# Circle with thickness (creates concentric circles)
create_circle([10, 10], 5, color="red", thickness=0.2)
```

### Batch Operations

```python
# Get all entities
entities = get_entities()

# Delete by criteria
delete_entities_by_layer("TEMP_LAYER")
delete_entities_by_color("red")
delete_entities_by_type("AcDbText")

# Modify existing entities
for entity in entities['entities']:
    if entity['type'] == 'AcDbLine':
        change_entity_color(entity['handle'], 'blue')
```

### Structure Labeling

```python
# Create room with automatic labeling
create_structure("room", {
    "corner1": [0, 0],
    "corner2": [12, 8]
}, label="Living Room", thickness=0.25)

# Labels automatically go to ANNOTATION layer
```

## API Reference

### Core Information Tools

#### `get_drawing_info()`
**Purpose**: Get comprehensive drawing information
**Returns**: Drawing metadata, layer info, entity counts

```python
{
  "filename": "floor_plan.dwg",
  "path": "C:\\Projects\\floor_plan.dwg", 
  "entity_count": 150,
  "saved": true,
  "current_layer": "WALLS",
  "total_layers": 8,
  "available_layers": [...]
}
```

#### `get_entities(max_entities?)`
**Purpose**: Retrieve all entities with layer grouping
**Parameters**: 
- `max_entities` (optional): Limit number of entities returned

**Returns**: Complete entity list with metadata
```python
{
  "entities": [...],
  "entities_by_layer": {
    "WALLS": [25 entities],
    "DOORS": [8 entities]
  },
  "total_count": 150,
  "layer_summary": {"WALLS": 25, "DOORS": 8}
}
```

### Structure Creation

#### `create_structure(type, geometry, options)`
**Purpose**: Create intelligent building structures
**Parameters**:
- `structure_type`: "wall", "door", "window", "room", "furniture", etc.
- `geometry_data`: Geometry definition
- `color` (optional): Override default layer color
- `thickness` (optional): Structure thickness  
- `custom_layer` (optional): Custom layer name
- `label` (optional): Text label

**Examples**:
```python
# Wall
create_structure("wall", {
    "start": [0, 0], "end": [20, 0]
}, thickness=0.25, label="Exterior Wall")

# Room  
create_structure("room", {
    "corner1": [0, 0], "corner2": [12, 10]
}, label="Master Bedroom")

# Door with swing
create_structure("door", {
    "start": [8, 0], "end": [11, 0], "width": 3
}, label="Front Door")
```

### Basic Entity Creation

#### `create_line(start, end, color?, thickness?)`
```python
create_line([0, 0], [10, 0], "red", 0.1)
```

#### `create_circle(center, radius, color?, thickness?)`
```python
create_circle([5, 5], 3, "blue", 0.05)
```

#### `create_rectangle(corner1, corner2, color?, thickness?)`
```python
create_rectangle([0, 0], [10, 8], "green", 0.0)
```

#### `create_text(position, text, height?, color?)`
```python
create_text([5, 5], "Room Label", 0.5, "black")
```

#### `create_arc(center, radius, start_angle, end_angle, color?, thickness?)`
```python
create_arc([0, 0], 5, 0, 90, "cyan", 0.0)
```

### Layer Management

#### `create_or_get_layer(name, color?, description?)`
```python
create_or_get_layer("CUSTOM_LAYER", "yellow", "Custom elements")
```

#### `set_current_layer(name)`
```python
set_current_layer("WALLS")
```

### Entity Deletion

#### `delete_entity_by_handle(handle)`
```python
delete_entity_by_handle("ABC123")
```

#### `delete_entities_by_layer(layer_name)`
```python
delete_entities_by_layer("TEMP_LAYER")
```

#### `delete_entities_by_type(entity_type)`
```python
delete_entities_by_type("AcDbText")  # Delete all text
```

#### `delete_entities_by_color(color)`
```python
delete_entities_by_color("red")
```

#### `delete_last_entities(count?)`
```python
delete_last_entities(5)  # Delete last 5 entities
```

### Utility Tools

#### `undo_last_operation()`
```python
undo_last_operation()  # Ctrl+Z equivalent
```

#### `change_entity_color(handle, color)`
```python
change_entity_color("ABC123", "blue")
```

#### `zoom_extents()`
```python
zoom_extents()  # Zoom to show all objects
```

## Predefined Layer System

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

The system intelligently assigns layers based on keywords:

```python
structure_mapping = {
    "wall", "partition" → "WALLS",
    "door", "opening" → "DOORS", 
    "window" → "WINDOWS",
    "furniture", "chair", "table", "bed" → "FURNITURE",
    "outlet", "switch", "light" → "ELECTRICAL",
    "toilet", "sink", "pipe" → "PLUMBING",
    "vent", "duct", "hvac" → "HVAC",
    "beam", "column" → "STRUCTURE",
    "text", "dimension" → "ANNOTATION"
}
```

## Real-World Examples

### Example 1: Complete Floor Plan with Claude

**User to Claude:**
```
Create a simple apartment floor plan:
- 20x15 foot room with 6-inch walls
- Front door (3 feet wide) on the south wall, centered
- Two windows (4 feet each) on the east and west walls
- Kitchen area in the northwest corner (8x6 feet)
- Add labels for each area
```

**Claude's Response Process:**
1. Uses `create_structure("room", ...)` for main room
2. Uses `create_structure("door", ...)` for entrance  
3. Uses `create_structure("window", ...)` for windows
4. Uses `create_structure("room", ...)` for kitchen
5. Uses `create_text(...)` for labels
6. Uses `zoom_extents()` to show complete drawing

### Example 2: Layer Management Workflow

**User to Claude:**
```
I need to organize my drawing. Can you:
1. Show me all current layers and their entity counts
2. Move all red entities to a new "HIGHLIGHTED" layer
3. Delete everything on the "TEMP" layer
4. Change all text to be green colored
```

**Claude's Process:**
1. `get_entities()` to analyze current state
2. `create_or_get_layer("HIGHLIGHTED", "red")`
3. Filter entities by color, move to new layer
4. `delete_entities_by_layer("TEMP")`
5. `change_entity_color()` for all text entities

### Example 3: Electrical Layout

**User to Claude:**
```
Add electrical outlets and switches to this room:
- Outlets every 8 feet along the walls
- Light switches by each door
- Ceiling light in the center
- Put everything on the electrical layer with proper labels
```

**Claude's Process:**
1. `get_entities()` to find existing walls and doors
2. Calculate outlet positions along walls
3. `create_structure("outlet", ...)` for each outlet
4. `create_structure("switch", ...)` by doors
5. `create_structure("light", ...)` at room center
6. All automatically go to ELECTRICAL layer

## Troubleshooting

### Common Issues and Solutions

#### ❌ "Not connected to AutoCAD"

**Symptoms:**
- Error messages about AutoCAD connection
- Server fails to start

**Solutions:**
1. **Start AutoCAD First**
   ```bash
   # Always start AutoCAD before the MCP server
   # Open AutoCAD application manually
   ```

2. **Check AutoCAD Installation**
   ```bash
   # Verify AutoCAD is installed and licensed
   # Try opening AutoCAD manually
   ```

3. **Administrative Privileges**
   ```bash
   # Run command prompt as Administrator
   # Try running the server with elevated privileges
   ```

4. **COM Registration Issues**
   ```bash
   # Re-register AutoCAD COM components
   # Run as Administrator:
   regsvr32 "C:\Program Files\Autodesk\AutoCAD 202X\acad.exe" /regserver
   ```

#### ❌ Claude Desktop Not Connecting

**Symptoms:**
- MCP server not appearing in Claude
- Tools not available

**Solutions:**
1. **Check Configuration Path**
   ```json
   // Verify paths in claude_desktop_config.json are correct
   {
     "mcpServers": {
       "autocad-mcp": {
         "command": "python",
         "args": ["C:\\CORRECT\\PATH\\test_script.py"]
       }
     }
   }
   ```

2. **Restart Claude Desktop**
   - Close Claude completely (check system tray)
   - Restart the application
   - Wait for full initialization

3. **Check Python Environment**
   ```bash
   # Verify Python and dependencies are accessible
   python -c "import mcp; print('MCP available')"
   python -c "import win32com.client; print('COM available')"
   ```

4. **Test Server Standalone**
   ```bash
   # Test the server works independently
   python test_script.py
   ```

#### ❌ COM Timeout Errors

**Symptoms:**
- Operations taking too long
- "COM operation timed out" errors

**Solutions:**
1. **Close Unnecessary AutoCAD Drawings**
   - Keep only one drawing open
   - Close large/complex drawings

2. **Increase Delay Parameters**
   ```python
   # In server.py, increase delay in safe_operation
   def safe_operation(self, operation_func, delay: float = 2.0):  # Increased delay
   ```

3. **Process in Smaller Batches**
   ```python
   # For large operations, break into smaller chunks
   entities = get_entities(max_entities=100)  # Limit batch size
   ```

#### ❌ Layer Creation Issues

**Symptoms:**
- Layers not being created
- "Layer already exists" errors

**Solutions:**
1. **Check Drawing Permissions**
   - Ensure drawing is not read-only
   - Save the drawing first

2. **Verify Layer Names**
   ```python
   # Use valid layer names (no special characters)
   create_or_get_layer("VALID_LAYER_NAME", "white")
   ```

3. **Layer Naming Conflicts**
   - Check for existing layers with same name
   - Use unique layer names

#### ❌ Performance Issues

**Symptoms:**
- Slow response times
- High memory usage
- AutoCAD becomes unresponsive

**Solutions:**
1. **Optimize Entity Queries**
   ```python
   # Use max_entities for large drawings
   entities = get_entities(max_entities=500)
   
   # Query specific layers instead of all entities
   delete_entities_by_layer("TARGET_LAYER")
   ```

2. **Batch Operations**
   ```python
   # Group similar operations together
   handles = ["handle1", "handle2", "handle3"]
   delete_entities_by_handles(handles)  # Batch delete
   ```

3. **AutoCAD Settings**
   - Turn off automatic save during operations
   - Disable visual effects/animations
   - Close unnecessary AutoCAD palettes

### Debugging Tips

#### Enable Verbose Logging
```python
# Add to test_script.py for debugging
import logging
logging.basicConfig(level=logging.DEBUG)

# Add debug prints in operations
print(f"DEBUG: Operation started - {operation_name}")
```

#### Test Individual Components
```python
# Test AutoCAD connection separately
from autocad_mcp import AutoCADCOMServer
server = AutoCADCOMServer()
print(f"Connection: {server.connect_to_autocad()}")

# Test specific operations
result = server.create_line([0,0], [10,0])
print(f"Line creation: {result}")
```

#### Monitor AutoCAD State
```python
# Check AutoCAD responsiveness
info = server.get_drawing_info()
print(f"AutoCAD responsive: {'error' not in info}")
```

## Development

### Extending the Server

#### Adding New Structure Types

1. **Define Layer Mapping**
```python
# Add to layer_definitions in server.py
self.layer_definitions["custom_type"] = {
    "name": "CUSTOM_LAYER", 
    "color": "yellow", 
    "description": "Custom structure type"
}
```

2. **Create Structure Method**
```python
def create_custom_structure(self, geometry_data, color="yellow", thickness=0.0):
    """Create custom structure"""
    # Implementation here
    pass
```

3. **Add to Structure Mapping**
```python
# Add to get_layer_for_structure_type method
if "custom" in structure_type_lower:
    return "CUSTOM_LAYER"
```

4. **Update MCP Tools**
```python
# Add to setup_tools method
types.Tool(
    name="create_custom_structure",
    description="Create custom structure type",
    inputSchema={...}
)
```

#### Adding New Entity Properties

```python
def create_line_with_custom_properties(self, start, end, **kwargs):
    """Enhanced line creation with custom properties"""
    line = self.create_line(start, end, kwargs.get('color', 'white'))
    
    # Add custom properties
    if 'lineweight' in kwargs:
        line.Lineweight = kwargs['lineweight']
    
    if 'linetype' in kwargs:
        line.Linetype = kwargs['linetype']
    
    return line
```

### Testing

#### Unit Tests
```python
# tests/test_server.py
import unittest
from autocad_mcp import AutoCADCOMServer

class TestAutoCADServer(unittest.TestCase):
    def setUp(self):
        self.server = AutoCADCOMServer()
    
    def test_connection(self):
        result = self.server.connect_to_autocad()
        self.assertTrue(result)
    
    def test_layer_creation(self):
        result = self.server.create_or_get_layer("TEST_LAYER")
        self.assertTrue(result)
```

#### Integration Tests
```python
# Test complete workflows
def test_room_creation_workflow():
    server = AutoCADCOMServer()
    
    # Create room
    result = server.create_structure("room", {
        "corner1": [0, 0],
        "corner2": [10, 10]
    }, label="Test Room")
    
    assert result["success"]
    
    # Verify entities created
    entities = server.get_entities()
    assert entities["total_count"] > 0
```

### Performance Optimization

#### For Large Drawings
```python
# Batch processing for large operations
def process_large_dataset(self, data_list, batch_size=100):
    """Process large datasets in batches"""
    for i in range(0, len(data_list), batch_size):
        batch = data_list[i:i+batch_size]
        self.process_batch(batch)
        time.sleep(0.1)  # Prevent COM overload
```

#### Memory Management
```python
# Clean up COM objects explicitly
def cleanup_com_objects(self):
    """Clean up COM references"""
    if self.doc:
        self.doc = None
    if self.acad_app:
        self.acad_app = None
    pythoncom.CoUninitialize()
```

### Contributing

1. **Fork the repository**
2. **Create feature branch**: `git checkout -b feature/new-feature`
3. **Add tests** for new functionality
4. **Update documentation**
5. **Submit pull request**

#### Code Style
- Follow PEP 8 guidelines
- Use type hints where possible
- Add docstrings for all public methods
- Include error handling for COM operations

#### Testing Requirements
- All new features must include tests
- Maintain >90% code coverage
- Test with multiple AutoCAD versions if possible

## FAQ

### Q: Can I use this with AutoCAD LT?
**A**: Yes, AutoCAD LT supports COM automation and should work with this server.

### Q: Does this work with other CAD software?
**A**: This server is specifically designed for AutoCAD's COM interface. Other CAD software would require different implementations.

### Q: Can I run multiple instances?
**A**: You can connect to multiple AutoCAD instances, but each MCP server instance connects to one AutoCAD instance.

### Q: Is this compatible with AutoCAD on Mac?
**A**: No, this uses Windows COM automation which is not available on Mac. AutoCAD for Mac uses a different architecture.

### Q: Can I modify existing drawings?
**A**: Yes, the server can modify any drawing that AutoCAD can open, including reading entities, modifying properties, and adding new content.

### Q: How do I handle coordinate systems?
**A**: The server uses AutoCAD's current coordinate system. All coordinates are in AutoCAD's current units (typically decimal units).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Legal Notice

This software requires a valid AutoCAD license for operation. AutoCAD is a trademark of Autodesk, Inc. This project is not affiliated with or endorsed by Autodesk.

**Ready to revolutionize your AutoCAD workflow? Start creating intelligent drawings with Claude today!** 
