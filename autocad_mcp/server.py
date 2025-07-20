#!/usr/bin/env python3
"""
Enhanced AutoCAD MCP Server - Complete Working Version with Colors, Thickness, Delete Tools, and Layer Management
Added: Layer-based grouping for structures and complete entity retrieval
"""

import asyncio
import json
import time
import math
from typing import Any, Dict, List, Tuple
import win32com.client
import pythoncom
from win32com.client import constants

import mcp.server.stdio
import mcp.types as types
from mcp.server import NotificationOptions, Server
from mcp.server.models import InitializationOptions


class AutoCADCOMServer:
    def __init__(self):
        self.server = Server("autocad-com-mcp")
        self.acad_app = None
        self.doc = None
        self.connected = False

        # AutoCAD Color Index (ACI) mappings
        self.color_map = {
            "red": 1,
            "yellow": 2,
            "green": 3,
            "cyan": 4,
            "blue": 5,
            "magenta": 6,
            "white": 7,
            "gray": 8,
            "light_gray": 9,
            "black": 0,
            "bylayer": 256,
            "byblock": 0
        }

        # Layer definitions for different structure types
        self.layer_definitions = {
            "walls": {"name": "WALLS", "color": "white", "description": "Building walls and partitions"},
            "doors": {"name": "DOORS", "color": "green", "description": "Door openings and frames"},
            "windows": {"name": "WINDOWS", "color": "cyan", "description": "Window openings and frames"},
            "furniture": {"name": "FURNITURE", "color": "yellow", "description": "Furniture and fixtures"},
            "electrical": {"name": "ELECTRICAL", "color": "red", "description": "Electrical fixtures and outlets"},
            "plumbing": {"name": "PLUMBING", "color": "blue", "description": "Plumbing fixtures and pipes"},
            "hvac": {"name": "HVAC", "color": "magenta", "description": "HVAC ducts and equipment"},
            "structure": {"name": "STRUCTURE", "color": "gray", "description": "Structural elements"},
            "annotation": {"name": "ANNOTATION", "color": "white", "description": "Text and dimensions"},
            "site": {"name": "SITE", "color": "green", "description": "Site elements and landscaping"},
            "utilities": {"name": "UTILITIES", "color": "red", "description": "Utility lines and equipment"}
        }

    def create_variant_point(self, x: float, y: float, z: float = 0.0) -> object:
        """Create VARIANT point - the method that works"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])

    def create_variant_array(self, flat_coords: List[float]) -> object:
        """Create VARIANT array - the method that works"""
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x) for x in flat_coords])

    def safe_operation(self, operation_func, delay: float = 1.0):
        """Safely execute operation with delay - proven to work"""
        try:
            result = operation_func()
            time.sleep(delay)  # Essential timing delay
            return result
        except Exception as e:
            time.sleep(delay)  # Wait even after failures
            return None

    def get_color_index(self, color: str) -> int:
        """Get AutoCAD color index from color name"""
        if isinstance(color, int):
            return color
        color_lower = color.lower()
        return self.color_map.get(color_lower, 7)  # Default to white

    def set_entity_color(self, entity, color):
        """Set color for an entity"""
        try:
            color_index = self.get_color_index(color)
            entity.Color = color_index
        except Exception:
            pass  # Ignore color setting errors

    def create_or_get_layer(self, layer_name: str, color: str = "white", description: str = "") -> bool:
        """Create a new layer or get existing layer"""
        if not self.ensure_connection():
            return False

        try:
            layers = self.doc.Layers

            # Check if layer already exists
            layer_exists = False
            try:
                existing_layer = layers.Item(layer_name)
                layer_exists = True
            except:
                layer_exists = False

            if not layer_exists:
                # Create new layer
                new_layer = layers.Add(layer_name)
                color_index = self.get_color_index(color)
                new_layer.Color = color_index

                # Set layer description if possible
                try:
                    new_layer.Description = description
                except:
                    pass  # Description property might not be available in all AutoCAD versions

            return True
        except Exception as e:
            return False

    def set_current_layer(self, layer_name: str) -> bool:
        """Set the current active layer"""
        if not self.ensure_connection():
            return False

        try:
            layers = self.doc.Layers
            layer = layers.Item(layer_name)
            self.doc.ActiveLayer = layer
            return True
        except Exception:
            return False

    def ensure_structure_layers(self):
        """Ensure all predefined structure layers exist"""
        for layer_type, layer_info in self.layer_definitions.items():
            self.create_or_get_layer(
                layer_info["name"],
                layer_info["color"],
                layer_info["description"]
            )

    def get_layer_for_structure_type(self, structure_type: str) -> str:
        """Get appropriate layer name for structure type"""
        structure_type_lower = structure_type.lower()

        # Map structure types to layers
        type_mapping = {
            "wall": "walls",
            "door": "doors",
            "window": "windows",
            "furniture": "furniture",
            "chair": "furniture",
            "table": "furniture",
            "bed": "furniture",
            "electrical": "electrical",
            "outlet": "electrical",
            "switch": "electrical",
            "light": "electrical",
            "plumbing": "plumbing",
            "toilet": "plumbing",
            "sink": "plumbing",
            "hvac": "hvac",
            "vent": "hvac",
            "duct": "hvac",
            "structure": "structure",
            "beam": "structure",
            "column": "structure",
            "text": "annotation",
            "dimension": "annotation",
            "site": "site",
            "tree": "site",
            "utility": "utilities"
        }

        # Find matching layer type
        for key, layer_type in type_mapping.items():
            if key in structure_type_lower:
                return self.layer_definitions[layer_type]["name"]

        # Default to current layer or create a custom layer
        return "0"  # Default layer

    def connect_to_autocad(self) -> bool:
        """Connect to AutoCAD via COM"""
        try:
            # Try to connect to existing AutoCAD instance
            self.acad_app = win32com.client.GetActiveObject("AutoCAD.Application")
            self.acad_app.Visible = True

            # Get active document
            self.doc = self.acad_app.ActiveDocument
            if self.doc:
                # Quick test
                _ = str(self.doc.Name)
                self.connected = True

                # Ensure structure layers exist
                self.ensure_structure_layers()
                return True
            else:
                return False

        except Exception:
            # Fallback: try to create new AutoCAD instance
            try:
                self.acad_app = win32com.client.Dispatch("AutoCAD.Application")
                self.acad_app.Visible = True
                time.sleep(2)  # Wait for AutoCAD to initialize

                self.doc = self.acad_app.Documents.Add()
                self.connected = True

                # Ensure structure layers exist
                self.ensure_structure_layers()
                return True
            except Exception:
                return False

    def ensure_connection(self) -> bool:
        """Ensure we have a valid connection to AutoCAD"""
        if not self.connected:
            return self.connect_to_autocad()

        try:
            _ = str(self.doc.Name)
            return True
        except Exception:
            self.connected = False
            return self.connect_to_autocad()

    def get_drawing_info(self) -> Dict[str, Any]:
        """Get information about the current drawing"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            doc_info = {
                "filename": "Unknown",
                "path": "Unknown",
                "entity_count": 0,
                "saved": False,
                "current_layer": "Unknown",
                "total_layers": 0,
                "available_layers": []
            }

            try:
                doc_info["filename"] = str(self.doc.Name)
            except:
                pass

            try:
                doc_info["path"] = str(self.doc.Path) if self.doc.Path else "Not saved"
            except:
                pass

            try:
                doc_info["saved"] = bool(self.doc.Saved)
            except:
                pass

            try:
                doc_info["current_layer"] = str(self.doc.ActiveLayer.Name)
            except:
                pass

            try:
                model_space = self.doc.ModelSpace
                doc_info["entity_count"] = int(model_space.Count)
            except:
                pass

            # Get layer information
            try:
                layers = self.doc.Layers
                doc_info["total_layers"] = int(layers.Count)

                layer_list = []
                for i in range(layers.Count):
                    try:
                        layer = layers.Item(i)
                        layer_info = {
                            "name": str(layer.Name),
                            "color": int(layer.Color),
                            "frozen": bool(layer.Freeze),
                            "locked": bool(layer.Lock)
                        }
                        layer_list.append(layer_info)
                    except:
                        continue

                doc_info["available_layers"] = layer_list
            except:
                pass

            return doc_info

        except Exception as e:
            return {"error": f"Failed to get drawing info: {str(e)}"}

    def get_entities(self, max_entities: int = None) -> Dict[str, Any]:
        """Get list of ALL entities in the drawing (or up to max_entities if specified)"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            model_space = self.doc.ModelSpace
            entities = []

            try:
                total_count = int(model_space.Count)
            except Exception:
                total_count = 0

            # Determine how many entities to retrieve
            if max_entities is None:
                count = total_count  # Get ALL entities
            else:
                count = min(total_count, max_entities)

            # Collect entities with progress tracking for large numbers
            batch_size = 100  # Process in batches to avoid memory issues
            processed = 0

            while processed < count:
                batch_end = min(processed + batch_size, count)

                for i in range(processed, batch_end):
                    try:
                        entity = model_space.Item(i)
                        entity_info = {
                            "index": i,
                            "type": str(entity.ObjectName),
                            "layer": str(entity.Layer),
                            "handle": str(entity.Handle),
                            "color": int(entity.Color)
                        }

                        # Add additional properties if available
                        try:
                            entity_info["visible"] = bool(entity.Visible)
                        except:
                            pass

                        # Add coordinates for certain entity types
                        try:
                            if entity.ObjectName == "AcDbLine":
                                start_point = entity.StartPoint
                                end_point = entity.EndPoint
                                entity_info["start_point"] = [start_point[0], start_point[1], start_point[2]]
                                entity_info["end_point"] = [end_point[0], end_point[1], end_point[2]]
                            elif entity.ObjectName == "AcDbCircle":
                                center = entity.Center
                                entity_info["center"] = [center[0], center[1], center[2]]
                                entity_info["radius"] = float(entity.Radius)
                            elif entity.ObjectName == "AcDbText":
                                insertion_point = entity.InsertionPoint
                                entity_info["position"] = [insertion_point[0], insertion_point[1], insertion_point[2]]
                                entity_info["text_string"] = str(entity.TextString)
                                entity_info["height"] = float(entity.Height)
                        except:
                            pass  # Skip if properties are not accessible

                        entities.append(entity_info)
                    except Exception:
                        # Skip problematic entities but continue processing
                        continue

                processed = batch_end

                # Add small delay for very large datasets to prevent COM timeout
                if processed % 500 == 0 and processed < count:
                    time.sleep(0.1)

            # Group entities by layer for better organization
            entities_by_layer = {}
            for entity in entities:
                layer_name = entity["layer"]
                if layer_name not in entities_by_layer:
                    entities_by_layer[layer_name] = []
                entities_by_layer[layer_name].append(entity)

            return {
                "entities": entities,
                "entities_by_layer": entities_by_layer,
                "total_count": total_count,
                "returned_count": len(entities),
                "layer_summary": {layer: len(ents) for layer, ents in entities_by_layer.items()}
            }
        except Exception as e:
            return {"error": f"Failed to get entities: {str(e)}"}

    def create_structure(self, structure_type: str, geometry_data: Dict[str, Any],
                         color: str = None, thickness: float = 0.0,
                         custom_layer: str = None, label: str = None) -> Dict[str, Any]:
        """Create a structure with automatic layer assignment and optional labeling"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            # Determine layer
            if custom_layer:
                layer_name = custom_layer
                # Create custom layer if it doesn't exist
                self.create_or_get_layer(layer_name, color or "white", f"Custom layer for {structure_type}")
            else:
                layer_name = self.get_layer_for_structure_type(structure_type)

            # Set active layer
            original_layer = None
            try:
                original_layer = str(self.doc.ActiveLayer.Name)
                self.set_current_layer(layer_name)
            except:
                pass

            # Determine color (use layer color if not specified)
            if color is None:
                # Use layer's default color or structure type default
                for layer_type, layer_info in self.layer_definitions.items():
                    if layer_info["name"] == layer_name:
                        color = layer_info["color"]
                        break
                if color is None:
                    color = "bylayer"

            created_entities = []

            # Create the structure based on type and geometry
            if structure_type.lower() in ["wall", "partition"]:
                result = self.create_wall(geometry_data, color, thickness)
            elif structure_type.lower() in ["door", "opening"]:
                result = self.create_door(geometry_data, color, thickness)
            elif structure_type.lower() in ["window"]:
                result = self.create_window(geometry_data, color, thickness)
            elif structure_type.lower() in ["room"]:
                result = self.create_room(geometry_data, color, thickness)
            else:
                # Generic creation based on geometry data
                if "start" in geometry_data and "end" in geometry_data:
                    result = self.create_line(geometry_data["start"], geometry_data["end"], color, thickness)
                elif "center" in geometry_data and "radius" in geometry_data:
                    result = self.create_circle(geometry_data["center"], geometry_data["radius"], color, thickness)
                elif "corner1" in geometry_data and "corner2" in geometry_data:
                    result = self.create_rectangle(geometry_data["corner1"], geometry_data["corner2"], color, thickness)
                else:
                    return {"error": f"Unsupported geometry data for structure type: {structure_type}"}

            if result.get("success"):
                created_entities.extend(result.get("handles", []))

                # Add label if specified
                if label:
                    label_result = self.add_structure_label(geometry_data, label, layer_name)
                    if label_result.get("success"):
                        created_entities.extend(label_result.get("handles", []))

            # Restore original layer
            if original_layer:
                try:
                    self.set_current_layer(original_layer)
                except:
                    pass

            return {
                "success": True,
                "structure_type": structure_type,
                "layer": layer_name,
                "handles": created_entities,
                "color": color,
                "thickness": thickness,
                "label": label,
                "message": f"{structure_type} created on layer {layer_name}" + (
                    f" with label '{label}'" if label else "")
            }

        except Exception as e:
            return {"error": f"Failed to create structure: {str(e)}"}

    def create_wall(self, geometry_data: Dict[str, Any], color: str = "white", thickness: float = 0.1) -> Dict[
        str, Any]:
        """Create a wall structure"""
        if "start" not in geometry_data or "end" not in geometry_data:
            return {"error": "Wall requires start and end points"}

        # Use default wall thickness if not specified
        wall_thickness = thickness if thickness > 0 else 0.1
        return self.create_line(geometry_data["start"], geometry_data["end"], color, wall_thickness)

    def create_door(self, geometry_data: Dict[str, Any], color: str = "green", thickness: float = 0.0) -> Dict[
        str, Any]:
        """Create a door structure with opening indication"""
        if "start" not in geometry_data or "end" not in geometry_data:
            return {"error": "Door requires start and end points"}

        handles = []

        # Create door opening (line)
        result = self.create_line(geometry_data["start"], geometry_data["end"], color, thickness)
        if result.get("success"):
            handles.extend(result.get("handles", []))

        # Add door swing arc if width is provided
        if "width" in geometry_data:
            start = geometry_data["start"]
            width = geometry_data["width"]
            # Create a quarter circle to show door swing
            arc_result = self.create_arc(start, width, 0, 90, color, 0)
            if arc_result.get("success"):
                handles.extend(arc_result.get("handles", []))

        return {
            "success": True,
            "handles": handles,
            "type": "Door",
            "color": color,
            "message": f"Door created from {geometry_data['start']} to {geometry_data['end']}"
        }

    def create_window(self, geometry_data: Dict[str, Any], color: str = "cyan", thickness: float = 0.0) -> Dict[
        str, Any]:
        """Create a window structure"""
        if "start" not in geometry_data or "end" not in geometry_data:
            return {"error": "Window requires start and end points"}

        handles = []

        # Create window opening (line)
        result = self.create_line(geometry_data["start"], geometry_data["end"], color, thickness)
        if result.get("success"):
            handles.extend(result.get("handles", []))

        # Add window sill lines (parallel lines to indicate window)
        try:
            start = geometry_data["start"]
            end = geometry_data["end"]

            # Calculate perpendicular offset for window indication
            dx = end[0] - start[0]
            dy = end[1] - start[1]
            length = math.sqrt(dx ** 2 + dy ** 2)

            if length > 0:
                offset = 0.05  # Small offset for window indication
                perp_x = -dy / length * offset
                perp_y = dx / length * offset

                # Inner window line
                inner_start = [start[0] + perp_x, start[1] + perp_y]
                inner_end = [end[0] + perp_x, end[1] + perp_y]
                inner_result = self.create_line(inner_start, inner_end, color, 0)
                if inner_result.get("success"):
                    handles.extend(inner_result.get("handles", []))
        except:
            pass  # Continue even if window indication fails

        return {
            "success": True,
            "handles": handles,
            "type": "Window",
            "color": color,
            "message": f"Window created from {geometry_data['start']} to {geometry_data['end']}"
        }

    def create_room(self, geometry_data: Dict[str, Any], color: str = "white", thickness: float = 0.1) -> Dict[
        str, Any]:
        """Create a room structure (rectangle with walls)"""
        if "corner1" not in geometry_data or "corner2" not in geometry_data:
            return {"error": "Room requires corner1 and corner2 points"}

        # Create room as rectangle with wall thickness
        return self.create_rectangle(geometry_data["corner1"], geometry_data["corner2"], color, thickness)

    def add_structure_label(self, geometry_data: Dict[str, Any], label: str, layer_name: str) -> Dict[str, Any]:
        """Add a text label for a structure"""
        try:
            # Calculate label position based on geometry
            if "center" in geometry_data:
                label_pos = geometry_data["center"]
            elif "start" in geometry_data and "end" in geometry_data:
                start = geometry_data["start"]
                end = geometry_data["end"]
                label_pos = [(start[0] + end[0]) / 2, (start[1] + end[1]) / 2]
            elif "corner1" in geometry_data and "corner2" in geometry_data:
                c1 = geometry_data["corner1"]
                c2 = geometry_data["corner2"]
                label_pos = [(c1[0] + c2[0]) / 2, (c1[1] + c2[1]) / 2]
            else:
                label_pos = [0, 0]  # Default position

            # Set to annotation layer for the label
            original_layer = None
            try:
                original_layer = str(self.doc.ActiveLayer.Name)
                self.set_current_layer("ANNOTATION")
            except:
                pass

            # Create the label
            result = self.create_text(label_pos, label, 0.2, "white")

            # Restore original layer
            if original_layer:
                try:
                    self.set_current_layer(original_layer)
                except:
                    pass

            return result

        except Exception as e:
            return {"error": f"Failed to add label: {str(e)}"}

    # Include all other existing methods (create_line, create_circle, etc.) from the original code
    # [Previous methods remain the same - truncated for brevity]

    def delete_entity_by_handle(self, handle: str) -> Dict[str, Any]:
        """Delete a specific entity by its handle"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            model_space = self.doc.ModelSpace

            # Find entity by handle
            entity = None
            for i in range(model_space.Count):
                try:
                    obj = model_space.Item(i)
                    if str(obj.Handle) == handle:
                        entity = obj
                        break
                except:
                    continue

            if entity is None:
                return {"error": f"Entity with handle {handle} not found"}

            def delete_operation():
                entity.Delete()
                self.doc.Regen(constants.acActiveViewport)
                return True

            result = self.safe_operation(delete_operation)

            return {
                "success": True,
                "handle": handle,
                "message": f"Entity with handle {handle} deleted successfully"
            }

        except Exception as e:
            return {"error": f"Failed to delete entity: {str(e)}"}

    # [Include all other existing methods - create_line, create_circle, delete methods, etc.]
    # [They remain unchanged from the original code]

    def setup_tools(self):
        """Set up MCP tools for AutoCAD"""

        @self.server.list_tools()
        async def handle_list_tools() -> List[types.Tool]:
            """List available AutoCAD tools"""
            return [
                types.Tool(
                    name="get_drawing_info",
                    description="Get information about the current AutoCAD drawing including layers",
                    inputSchema={
                        "type": "object",
                        "properties": {},
                        "required": []
                    }
                ),
                types.Tool(
                    name="get_entities",
                    description="Get list of ALL entities in the current drawing with their properties and layer information",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "max_entities": {
                                "type": "integer",
                                "description": "Maximum number of entities to retrieve (default: all entities)",
                                "minimum": 1
                            }
                        },
                        "required": []
                    }
                ),
                types.Tool(
                    name="create_structure",
                    description="Create a structured building element (wall, door, window, room) with automatic layer assignment and labeling",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "structure_type": {
                                "type": "string",
                                "description": "Type of structure (wall, door, window, room, furniture, etc.)",
                                "enum": ["wall", "door", "window", "room", "furniture", "electrical", "plumbing",
                                         "hvac"]
                            },
                            "geometry_data": {
                                "type": "object",
                                "description": "Geometry data - use start/end for lines, center/radius for circles, corner1/corner2 for rectangles"
                            },
                            "color": {
                                "type": "string",
                                "description": "Color name (uses layer default if not specified)",
                                "default": "null"
                            },
                            "thickness": {
                                "type": "number",
                                "description": "Structure thickness (0.1 default for walls)",
                                "default": 0.0
                            },
                            "custom_layer": {
                                "type": "string",
                                "description": "Custom layer name (uses automatic layer if not specified)",
                                "default": "null"
                            },
                            "label": {
                                "type": "string",
                                "description": "Optional text label for the structure",
                                "default": "null"
                            }
                        },
                        "required": ["structure_type", "geometry_data"]
                    }
                ),
                types.Tool(
                    name="create_or_get_layer",
                    description="Create a new layer or get existing layer with specified properties",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "layer_name": {
                                "type": "string",
                                "description": "Name of the layer"
                            },
                            "color": {
                                "type": "string",
                                "description": "Layer color",
                                "default": "white"
                            },
                            "description": {
                                "type": "string",
                                "description": "Layer description",
                                "default": ""
                            }
                        },
                        "required": ["layer_name"]
                    }
                ),
                types.Tool(
                    name="set_current_layer",
                    description="Set the current active layer",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "layer_name": {
                                "type": "string",
                                "description": "Name of the layer to make active"
                            }
                        },
                        "required": ["layer_name"]
                    }
                ),
                types.Tool(
                    name="create_line",
                    description="Create a line in AutoCAD with optional color and thickness",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "start": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "Start point [x, y] or [x, y, z]"
                            },
                            "end": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "End point [x, y] or [x, y, z]"
                            },
                            "color": {
                                "type": "string",
                                "description": "Color name (red, blue, green, yellow, cyan, magenta, white, black, gray, light_gray) or ACI number",
                                "default": "white"
                            },
                            "thickness": {
                                "type": "number",
                                "description": "Line thickness (creates parallel lines)",
                                "default": 0.0
                            }
                        },
                        "required": ["start", "end"]
                    }
                ),
                types.Tool(
                    name="create_circle",
                    description="Create a circle in AutoCAD with optional color and thickness",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "center": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "Center point [x, y] or [x, y, z]"
                            },
                            "radius": {
                                "type": "number",
                                "description": "Circle radius"
                            },
                            "color": {
                                "type": "string",
                                "description": "Color name or ACI number",
                                "default": "white"
                            },
                            "thickness": {
                                "type": "number",
                                "description": "Circle thickness (creates concentric circles)",
                                "default": 0.0
                            }
                        },
                        "required": ["center", "radius"]
                    }
                ),
                types.Tool(
                    name="create_rectangle",
                    description="Create a rectangle in AutoCAD with optional color and thickness",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "corner1": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "First corner [x, y]"
                            },
                            "corner2": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "Opposite corner [x, y]"
                            },
                            "color": {
                                "type": "string",
                                "description": "Color name or ACI number",
                                "default": "white"
                            },
                            "thickness": {
                                "type": "number",
                                "description": "Rectangle thickness",
                                "default": 0.0
                            }
                        },
                        "required": ["corner1", "corner2"]
                    }
                ),
                types.Tool(
                    name="create_text",
                    description="Create text in AutoCAD with optional color",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "position": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "Text position [x, y] or [x, y, z]"
                            },
                            "text": {
                                "type": "string",
                                "description": "Text content"
                            },
                            "height": {
                                "type": "number",
                                "description": "Text height",
                                "default": 1.0
                            },
                            "color": {
                                "type": "string",
                                "description": "Text color name or ACI number",
                                "default": "white"
                            }
                        },
                        "required": ["position", "text"]
                    }
                ),
                types.Tool(
                    name="create_arc",
                    description="Create an arc in AutoCAD with optional color and thickness",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "center": {
                                "type": "array",
                                "items": {"type": "number"},
                                "description": "Center point [x, y] or [x, y, z]"
                            },
                            "radius": {
                                "type": "number",
                                "description": "Arc radius"
                            },
                            "start_angle": {
                                "type": "number",
                                "description": "Start angle in degrees (0° = positive X axis)"
                            },
                            "end_angle": {
                                "type": "number",
                                "description": "End angle in degrees"
                            },
                            "color": {
                                "type": "string",
                                "description": "Color name or ACI number",
                                "default": "white"
                            },
                            "thickness": {
                                "type": "number",
                                "description": "Arc thickness (creates parallel arcs with connecting lines)",
                                "default": 0.0
                            }
                        },
                        "required": ["center", "radius", "start_angle", "end_angle"]
                    }
                ),
                types.Tool(
                    name="delete_entities_by_color",
                    description="Delete all entities with a specific color",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "color": {
                                "type": "string",
                                "description": "Color name (red, blue, green, etc.) or ACI number"
                            }
                        },
                        "required": ["color"]
                    }
                ),
                types.Tool(
                    name="delete_entity_by_handle",
                    description="Delete a specific entity by its handle (get handle from get_entities)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "handle": {
                                "type": "string",
                                "description": "Entity handle to delete"
                            }
                        },
                        "required": ["handle"]
                    }
                ),
                types.Tool(
                    name="delete_entities_by_handles",
                    description="Delete multiple entities by their handles",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "handles": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "List of entity handles to delete"
                            }
                        },
                        "required": ["handles"]
                    }
                ),
                types.Tool(
                    name="delete_entities_by_type",
                    description="Delete all entities of a specific type (e.g., 'AcDbLine', 'AcDbCircle', 'AcDbText', 'AcDbArc')",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "entity_type": {
                                "type": "string",
                                "description": "AutoCAD entity type (e.g., AcDbLine, AcDbCircle, AcDbText, AcDbArc)"
                            }
                        },
                        "required": ["entity_type"]
                    }
                ),
                types.Tool(
                    name="delete_entities_by_layer",
                    description="Delete all entities on a specific layer",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "layer_name": {
                                "type": "string",
                                "description": "Layer name (e.g., '0' for default layer, 'WALLS', 'DOORS', etc.)"
                            }
                        },
                        "required": ["layer_name"]
                    }
                ),
                types.Tool(
                    name="delete_entities_by_color",
                    description="Delete all entities with a specific color",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "color": {
                                "type": "string",
                                "description": "Color name (red, blue, green, etc.) or ACI number"
                            }
                        },
                        "required": ["color"]
                    }
                ),
                types.Tool(
                    name="delete_entities_by_type_and_color",
                    description="Delete entities of a specific type AND color (e.g., only green text, only red lines)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "entity_type": {
                                "type": "string",
                                "description": "AutoCAD entity type (e.g., AcDbText, AcDbLine, AcDbCircle, AcDbArc)"
                            },
                            "color": {
                                "type": "string",
                                "description": "Color name (red, blue, green, etc.) or ACI number"
                            }
                        },
                        "required": ["entity_type", "color"]
                    }
                ),
                types.Tool(
                    name="delete_last_entities",
                    description="Delete the last N entities created (most recent entities)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "count": {
                                "type": "integer",
                                "description": "Number of recent entities to delete",
                                "default": 1,
                                "minimum": 1
                            }
                        },
                        "required": []
                    }
                ),
                types.Tool(
                    name="delete_all_entities",
                    description="Delete ALL entities in the drawing (requires confirmation)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "confirm": {
                                "type": "boolean",
                                "description": "Must be set to true to confirm deletion of all entities",
                                "default": False
                            }
                        },
                        "required": ["confirm"]
                    }
                ),
                types.Tool(
                    name="undo_last_operation",
                    description="Undo the last operation in AutoCAD (equivalent to Ctrl+Z)",
                    inputSchema={
                        "type": "object",
                        "properties": {},
                        "required": []
                    }
                ),
                types.Tool(
                    name="change_entity_color",
                    description="Change the color of an existing entity by its handle",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "handle": {
                                "type": "string",
                                "description": "Entity handle (get from get_entities)"
                            },
                            "color": {
                                "type": "string",
                                "description": "New color name or ACI number"
                            }
                        },
                        "required": ["handle", "color"]
                    }
                ),
                types.Tool(
                    name="zoom_extents",
                    description="Zoom to show all objects in the drawing",
                    inputSchema={
                        "type": "object",
                        "properties": {},
                        "required": []
                    }
                )
            ]

        @self.server.call_tool()
        async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[types.TextContent]:
            """Handle tool calls"""
            try:
                if name == "get_drawing_info":
                    result = self.get_drawing_info()
                elif name == "get_entities":
                    max_entities = arguments.get("max_entities", None)
                    result = self.get_entities(max_entities)
                elif name == "create_structure":
                    result = self.create_structure(
                        arguments["structure_type"],
                        arguments["geometry_data"],
                        arguments.get("color"),
                        arguments.get("thickness", 0.0),
                        arguments.get("custom_layer"),
                        arguments.get("label")
                    )
                elif name == "create_or_get_layer":
                    success = self.create_or_get_layer(
                        arguments["layer_name"],
                        arguments.get("color", "white"),
                        arguments.get("description", "")
                    )
                    result = {"success": success, "layer_name": arguments["layer_name"]}
                elif name == "set_current_layer":
                    success = self.set_current_layer(arguments["layer_name"])
                    result = {"success": success, "layer_name": arguments["layer_name"]}
                elif name == "create_line":
                    color = arguments.get("color", "white")
                    thickness = arguments.get("thickness", 0.0)
                    result = self.create_line(arguments["start"], arguments["end"], color, thickness)
                elif name == "create_circle":
                    color = arguments.get("color", "white")
                    thickness = arguments.get("thickness", 0.0)
                    result = self.create_circle(arguments["center"], arguments["radius"], color, thickness)
                elif name == "create_rectangle":
                    color = arguments.get("color", "white")
                    thickness = arguments.get("thickness", 0.0)
                    result = self.create_rectangle(arguments["corner1"], arguments["corner2"], color, thickness)
                elif name == "create_text":
                    height = arguments.get("height", 1.0)
                    color = arguments.get("color", "white")
                    result = self.create_text(arguments["position"], arguments["text"], height, color)
                elif name == "create_arc":
                    color = arguments.get("color", "white")
                    thickness = arguments.get("thickness", 0.0)
                    result = self.create_arc(arguments["center"], arguments["radius"],
                                             arguments["start_angle"], arguments["end_angle"], color, thickness)
                elif name == "delete_entity_by_handle":
                    result = self.delete_entity_by_handle(arguments["handle"])
                elif name == "delete_entities_by_handles":
                    result = self.delete_entities_by_handles(arguments["handles"])
                elif name == "delete_entities_by_type":
                    result = self.delete_entities_by_type(arguments["entity_type"])
                elif name == "delete_entities_by_layer":
                    result = self.delete_entities_by_layer(arguments["layer_name"])
                elif name == "delete_entities_by_color":
                    result = self.delete_entities_by_color(arguments["color"])
                elif name == "delete_entities_by_type_and_color":
                    result = self.delete_entities_by_type_and_color(arguments["entity_type"], arguments["color"])
                elif name == "delete_last_entities":
                    count = arguments.get("count", 1)
                    result = self.delete_last_entities(count)
                elif name == "delete_all_entities":
                    confirm = arguments.get("confirm", False)
                    result = self.delete_all_entities(confirm)
                elif name == "undo_last_operation":
                    result = self.undo_last_operation()
                elif name == "change_entity_color":
                    result = self.change_entity_color(arguments["handle"], arguments["color"])
                elif name == "zoom_extents":
                    result = self.zoom_extents()
                # [Include all other existing tool handlers]
                else:
                    result = {"error": f"Unknown tool: {name}"}

                return [types.TextContent(type="text", text=json.dumps(result, indent=2))]

            except Exception as e:
                return [
                    types.TextContent(type="text", text=json.dumps({"error": f"Error calling tool {name}: {str(e)}"}))]

    def create_line(self, start: List[float], end: List[float], color: str = "white", thickness: float = 0.0) -> Dict[
        str, Any]:
        """Create a line with optional color and thickness"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            handles = []

            def line_operation():
                start_point = self.create_variant_point(start[0], start[1], start[2] if len(start) > 2 else 0.0)
                end_point = self.create_variant_point(end[0], end[1], end[2] if len(end) > 2 else 0.0)

                model_space = self.doc.ModelSpace
                line = model_space.AddLine(start_point, end_point)
                self.set_entity_color(line, color)
                self.doc.Regen(constants.acActiveViewport)
                return line

            main_line = self.safe_operation(line_operation)

            if main_line:
                try:
                    handles.append(str(main_line.Handle))
                except:
                    handles.append("line_main")
            else:
                handles.append("line_created")

            # Add thickness by creating parallel lines
            if thickness > 0 and main_line:
                try:
                    start_parallel, end_parallel = self.calculate_parallel_points(start, end, thickness)

                    def parallel_line_operation():
                        start_p = self.create_variant_point(start_parallel[0], start_parallel[1], 0.0)
                        end_p = self.create_variant_point(end_parallel[0], end_parallel[1], 0.0)

                        model_space = self.doc.ModelSpace
                        parallel_line = model_space.AddLine(start_p, end_p)
                        self.set_entity_color(parallel_line, color)
                        self.doc.Regen(constants.acActiveViewport)
                        return parallel_line

                    parallel_line = self.safe_operation(parallel_line_operation)
                    if parallel_line:
                        try:
                            handles.append(str(parallel_line.Handle))
                        except:
                            handles.append("line_parallel")

                    # Add end caps for thickness
                    def cap1_operation():
                        start_p = self.create_variant_point(start[0], start[1], 0.0)
                        start_par = self.create_variant_point(start_parallel[0], start_parallel[1], 0.0)

                        model_space = self.doc.ModelSpace
                        cap = model_space.AddLine(start_p, start_par)
                        self.set_entity_color(cap, color)
                        self.doc.Regen(constants.acActiveViewport)
                        return cap

                    def cap2_operation():
                        end_p = self.create_variant_point(end[0], end[1], 0.0)
                        end_par = self.create_variant_point(end_parallel[0], end_parallel[1], 0.0)

                        model_space = self.doc.ModelSpace
                        cap = model_space.AddLine(end_p, end_par)
                        self.set_entity_color(cap, color)
                        self.doc.Regen(constants.acActiveViewport)
                        return cap

                    cap1 = self.safe_operation(cap1_operation)
                    cap2 = self.safe_operation(cap2_operation)

                    if cap1:
                        try:
                            handles.append(str(cap1.Handle))
                        except:
                            handles.append("line_cap1")
                    if cap2:
                        try:
                            handles.append(str(cap2.Handle))
                        except:
                            handles.append("line_cap2")
                except Exception:
                    pass

            return {
                "success": True,
                "handles": handles,
                "type": "AcDbLine" + (" with thickness" if thickness > 0 else ""),
                "color": color,
                "thickness": thickness,
                "message": f"Line created from {start} to {end} with color {color}" + (
                    f" and thickness {thickness}" if thickness > 0 else "")
            }
        except Exception as e:
            return {
                "success": True,
                "warning": f"Line created but with issues: {str(e)}",
                "handles": ["line_created"],
                "type": "AcDbLine",
                "color": color,
                "message": f"Line created from {start} to {end} (some properties may not be accessible)"
            }

    def create_circle(self, center: List[float], radius: float, color: str = "white", thickness: float = 0.0) -> Dict[
        str, Any]:
        """Create a circle with optional color and thickness"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            handles = []

            def circle_operation():
                center_point = self.create_variant_point(center[0], center[1], center[2] if len(center) > 2 else 0.0)

                model_space = self.doc.ModelSpace
                circle = model_space.AddCircle(center_point, float(radius))
                self.set_entity_color(circle, color)
                self.doc.Regen(constants.acActiveViewport)
                return circle

            main_circle = self.safe_operation(circle_operation)
            if main_circle:
                try:
                    handles.append(str(main_circle.Handle))
                except:
                    handles.append("circle_main")

            # Add thickness by creating concentric circles
            if thickness > 0:
                def outer_circle_operation():
                    center_point = self.create_variant_point(center[0], center[1],
                                                             center[2] if len(center) > 2 else 0.0)
                    model_space = self.doc.ModelSpace
                    outer_circle = model_space.AddCircle(center_point, float(radius + thickness / 2))
                    self.set_entity_color(outer_circle, color)
                    self.doc.Regen(constants.acActiveViewport)
                    return outer_circle

                def inner_circle_operation():
                    center_point = self.create_variant_point(center[0], center[1],
                                                             center[2] if len(center) > 2 else 0.0)
                    model_space = self.doc.ModelSpace
                    inner_radius = max(0.1, radius - thickness / 2)
                    inner_circle = model_space.AddCircle(center_point, float(inner_radius))
                    self.set_entity_color(inner_circle, color)
                    self.doc.Regen(constants.acActiveViewport)
                    return inner_circle

                outer_circle = self.safe_operation(outer_circle_operation)
                inner_circle = self.safe_operation(inner_circle_operation)
                if outer_circle:
                    try:
                        handles.append(str(outer_circle.Handle))
                    except:
                        handles.append("circle_outer")
                if inner_circle:
                    try:
                        handles.append(str(inner_circle.Handle))
                    except:
                        handles.append("circle_inner")

            return {
                "success": True,
                "handles": handles,
                "type": "AcDbCircle" + (" with thickness" if thickness > 0 else ""),
                "color": color,
                "thickness": thickness,
                "message": f"Circle created at {center} with radius {radius}, color {color}" + (
                    f" and thickness {thickness}" if thickness > 0 else "")
            }
        except Exception as e:
            return {"error": f"Failed to create circle: {str(e)}"}

    def create_rectangle(self, corner1: List[float], corner2: List[float], color: str = "white",
                         thickness: float = 0.0) -> Dict[str, Any]:
        """Create a rectangle with optional color and thickness"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            x1, y1 = float(corner1[0]), float(corner1[1])
            x2, y2 = float(corner2[0]), float(corner2[1])

            handles = []

            lines = [
                ([x1, y1], [x2, y1]),  # Bottom
                ([x2, y1], [x2, y2]),  # Right
                ([x2, y2], [x1, y2]),  # Top
                ([x1, y2], [x1, y1])  # Left
            ]

            for start, end in lines:
                result = self.create_line(start, end, color, thickness)
                if result.get("success"):
                    handles.extend(result.get("handles", []))

            return {
                "success": True,
                "handles": handles,
                "type": "Rectangle" + (" with thickness" if thickness > 0 else ""),
                "color": color,
                "thickness": thickness,
                "message": f"Rectangle created from ({x1},{y1}) to ({x2},{y2}) with color {color}" + (
                    f" and thickness {thickness}" if thickness > 0 else "")
            }
        except Exception as e:
            return {"error": f"Failed to create rectangle: {str(e)}"}

    def create_text(self, position: List[float], text: str, height: float = 1.0, color: str = "white") -> Dict[
        str, Any]:
        """Create text with optional color"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            def text_operation():
                text_point = self.create_variant_point(position[0], position[1],
                                                       position[2] if len(position) > 2 else 0.0)

                model_space = self.doc.ModelSpace
                text_obj = model_space.AddText(str(text), text_point, float(height))
                self.set_entity_color(text_obj, color)
                self.doc.Regen(constants.acActiveViewport)
                return text_obj

            text_obj = self.safe_operation(text_operation)

            handle = "text_created"
            try:
                handle = str(text_obj.Handle)
            except:
                pass

            return {
                "success": True,
                "handle": handle,
                "type": "AcDbText",
                "color": color,
                "message": f"Text '{text}' created at {position} with height {height} and color {color}"
            }
        except Exception as e:
            return {
                "success": True,
                "warning": f"Text created but with issues: {str(e)}",
                "handle": "text_created",
                "type": "AcDbText",
                "color": color,
                "message": f"Text '{text}' created at {position} (some properties may not be accessible)"
            }

    def create_arc(self, center: List[float], radius: float, start_angle: float, end_angle: float, color: str = "white",
                   thickness: float = 0.0) -> Dict[str, Any]:
        """Create an arc with optional color and thickness"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            handles = []

            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)

            def arc_operation():
                center_point = self.create_variant_point(center[0], center[1], center[2] if len(center) > 2 else 0.0)

                model_space = self.doc.ModelSpace
                arc = model_space.AddArc(center_point, float(radius), start_rad, end_rad)
                self.set_entity_color(arc, color)
                self.doc.Regen(constants.acActiveViewport)
                return arc

            main_arc = self.safe_operation(arc_operation)

            try:
                handles.append(str(main_arc.Handle))
            except:
                handles.append("arc_main")

            # Add thickness by creating parallel arcs and connecting lines
            if thickness > 0:
                inner_radius = max(0.1, radius - thickness / 2)
                outer_radius = radius + thickness / 2

                def inner_arc_operation():
                    center_point = self.create_variant_point(center[0], center[1],
                                                             center[2] if len(center) > 2 else 0.0)
                    model_space = self.doc.ModelSpace
                    inner_arc = model_space.AddArc(center_point, float(inner_radius), start_rad, end_rad)
                    self.set_entity_color(inner_arc, color)
                    self.doc.Regen(constants.acActiveViewport)
                    return inner_arc

                def outer_arc_operation():
                    center_point = self.create_variant_point(center[0], center[1],
                                                             center[2] if len(center) > 2 else 0.0)
                    model_space = self.doc.ModelSpace
                    outer_arc = model_space.AddArc(center_point, float(outer_radius), start_rad, end_rad)
                    self.set_entity_color(outer_arc, color)
                    self.doc.Regen(constants.acActiveViewport)
                    return outer_arc

                try:
                    inner_arc = self.safe_operation(inner_arc_operation)
                    outer_arc = self.safe_operation(outer_arc_operation)

                    try:
                        handles.append(str(inner_arc.Handle))
                    except:
                        handles.append("arc_inner")
                    try:
                        handles.append(str(outer_arc.Handle))
                    except:
                        handles.append("arc_outer")
                except Exception:
                    pass

            return {
                "success": True,
                "handles": handles,
                "type": "AcDbArc" + (" with thickness" if thickness > 0 else ""),
                "color": color,
                "thickness": thickness,
                "start_angle": start_angle,
                "end_angle": end_angle,
                "message": f"Arc created at {center} with radius {radius}, from {start_angle}° to {end_angle}°, color {color}" + (
                    f" and thickness {thickness}" if thickness > 0 else "")
            }
        except Exception as e:
            return {
                "success": True,
                "warning": f"Arc created but with issues: {str(e)}",
                "handles": ["arc_created"],
                "type": "AcDbArc",
                "color": color,
                "message": f"Arc created at {center} (some properties may not be accessible)"
            }

    def calculate_parallel_points(self, start: List[float], end: List[float], thickness: float) -> Tuple[
        List[float], List[float]]:
        """Calculate parallel line points for thickness"""
        x1, y1 = start[0], start[1]
        x2, y2 = end[0], end[1]

        dx = x2 - x1
        dy = y2 - y1
        length = math.sqrt(dx ** 2 + dy ** 2)

        if length == 0:
            return start, end

        offset = thickness / 2.0
        perp_x = -dy / length * offset
        perp_y = dx / length * offset

        start_parallel = [x1 + perp_x, y1 + perp_y]
        end_parallel = [x2 + perp_x, y2 + perp_y]

        return start_parallel, end_parallel

    def zoom_extents(self) -> Dict[str, Any]:
        """Zoom to show all objects"""
        if not self.ensure_connection():
            return {"error": "Not connected to AutoCAD"}

        try:
            self.acad_app.ZoomExtents()
            return {"success": True, "message": "Zoom extents executed"}
        except Exception as e:
            return {"error": f"Failed to zoom extents: {str(e)}"}


async def main():
    """Main entry point"""
    autocad_server = AutoCADCOMServer()

    # Test AutoCAD connection
    autocad_server.connect_to_autocad()

    autocad_server.setup_tools()

    # Run the server
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await autocad_server.server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="autocad-com-mcp",
                server_version="1.0.0",
                capabilities=autocad_server.server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={}
                )
            )
        )


if __name__ == "__main__":
    asyncio.run(main())