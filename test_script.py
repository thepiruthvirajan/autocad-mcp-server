#!/usr/bin/env python3
"""
AutoCAD test with delays and error handling - slow and steady approach
"""

import win32com.client
import pythoncom
import time


def create_variant_point(x, y, z=0):
    """Helper function to create VARIANT point"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])


def create_variant_array(flat_coords):
    """Helper function to create VARIANT array from flat coordinates"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x) for x in flat_coords])


def safe_operation(operation_name, operation_func, delay=0.5):
    """Safely execute an operation with error handling and delay"""
    try:
        print(f"   Trying {operation_name}...")
        result = operation_func()
        time.sleep(delay)  # Wait between operations
        print(f"   ✓ {operation_name} succeeded!")
        return result, True
    except Exception as e:
        print(f"   ✗ {operation_name} failed: {e}")
        time.sleep(delay)  # Wait even after failures
        return None, False


def test_autocad_slow_steady():
    """Test AutoCAD operations slowly with delays"""
    print("=== Testing AutoCAD Slowly and Steadily ===")

    try:
        # Step 1: Connect to AutoCAD
        print("1. Connecting to AutoCAD...")
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
        print("   ✓ Connected to AutoCAD")
        time.sleep(1)

        # Step 2: Get document
        print("2. Getting active document...")
        doc = acad.ActiveDocument
        print(f"   ✓ Document: {doc.Name}")
        time.sleep(1)

        # Step 3: Get ModelSpace
        print("3. Getting ModelSpace...")
        model_space = doc.ModelSpace
        initial_count = model_space.Count
        print(f"   ✓ ModelSpace has {initial_count} entities")
        time.sleep(1)

        successful_operations = []

        # Step 4: Test line creation with VARIANT
        print("4. Testing line creation with VARIANT...")

        def create_line():
            start_point = create_variant_point(0, 0, 0)
            end_point = create_variant_point(10, 10, 0)
            line = model_space.AddLine(start_point, end_point)
            doc.Regen(1)
            return line

        line_result, line_success = safe_operation("Line creation", create_line, 1.0)
        if line_success:
            successful_operations.append("Line")

        # Step 5: Test another line (different coordinates)
        print("5. Testing second line...")

        def create_line2():
            start_point = create_variant_point(15, 0, 0)
            end_point = create_variant_point(25, 10, 0)
            line = model_space.AddLine(start_point, end_point)
            doc.Regen(1)
            return line

        line2_result, line2_success = safe_operation("Second line", create_line2, 1.0)
        if line2_success:
            successful_operations.append("Second Line")

        # Step 6: Test circle creation with VARIANT
        print("6. Testing circle creation with VARIANT...")

        def create_circle():
            center_point = create_variant_point(35, 5, 0)
            radius = 5.0
            circle = model_space.AddCircle(center_point, radius)
            doc.Regen(1)
            return circle

        circle_result, circle_success = safe_operation("Circle creation", create_circle, 1.0)
        if circle_success:
            successful_operations.append("Circle")

        # Step 7: Test text creation with VARIANT
        print("7. Testing text creation with VARIANT...")

        def create_text():
            text_point = create_variant_point(0, 15, 0)
            text_obj = model_space.AddText("TEST", text_point, 2.0)
            doc.Regen(1)
            return text_obj

        text_result, text_success = safe_operation("Text creation", create_text, 1.0)
        if text_success:
            successful_operations.append("Text")

        # Step 8: Test rectangle with 4 lines (one at a time)
        print("8. Testing rectangle with 4 individual lines...")
        rectangle_lines = []
        x1, y1, x2, y2 = 50, 0, 60, 10

        # Bottom line
        def create_bottom_line():
            return model_space.AddLine(create_variant_point(x1, y1, 0), create_variant_point(x2, y1, 0))

        bottom_line, bottom_success = safe_operation("Bottom line", create_bottom_line, 1.0)
        if bottom_success:
            rectangle_lines.append("Bottom")

        # Right line
        def create_right_line():
            return model_space.AddLine(create_variant_point(x2, y1, 0), create_variant_point(x2, y2, 0))

        right_line, right_success = safe_operation("Right line", create_right_line, 1.0)
        if right_success:
            rectangle_lines.append("Right")

        # Top line
        def create_top_line():
            return model_space.AddLine(create_variant_point(x2, y2, 0), create_variant_point(x1, y2, 0))

        top_line, top_success = safe_operation("Top line", create_top_line, 1.0)
        if top_success:
            rectangle_lines.append("Top")

        # Left line
        def create_left_line():
            return model_space.AddLine(create_variant_point(x1, y2, 0), create_variant_point(x1, y1, 0))

        left_line, left_success = safe_operation("Left line", create_left_line, 1.0)
        if left_success:
            rectangle_lines.append("Left")

        if rectangle_lines:
            successful_operations.append(f"Rectangle ({len(rectangle_lines)} lines)")

        # Step 9: Test polyline (simple triangle)
        print("9. Testing simple polyline triangle...")

        def create_triangle():
            points = create_variant_array([
                70, 0, 0,  # Point 1
                80, 0, 0,  # Point 2
                75, 10, 0  # Point 3
            ])
            poly = model_space.AddPolyline(points)
            poly.Closed = True
            doc.Regen(1)
            return poly

        triangle_result, triangle_success = safe_operation("Triangle polyline", create_triangle, 1.0)
        if triangle_success:
            successful_operations.append("Triangle")

        # Step 10: Test LightWeightPolyline (2D rectangle)
        print("10. Testing LightWeightPolyline rectangle...")

        def create_lw_rectangle():
            points_2d = create_variant_array([
                85, 0,  # Point 1
                95, 0,  # Point 2
                95, 10,  # Point 3
                85, 10  # Point 4
            ])
            lwpoly = model_space.AddLightWeightPolyline(points_2d)
            lwpoly.Closed = True
            doc.Regen(1)
            return lwpoly

        lw_result, lw_success = safe_operation("LightWeight rectangle", create_lw_rectangle, 1.0)
        if lw_success:
            successful_operations.append("LW Rectangle")

        # Step 11: Test arc
        print("11. Testing arc...")

        def create_arc():
            center = create_variant_point(105, 5, 0)
            radius = 5.0
            start_angle = 0.0
            end_angle = 1.5708  # 90 degrees
            arc = model_space.AddArc(center, radius, start_angle, end_angle)
            doc.Regen(1)
            return arc

        arc_result, arc_success = safe_operation("Arc", create_arc, 1.0)
        if arc_success:
            successful_operations.append("Arc")

        # Step 12: Final operations
        print("12. Final operations...")
        time.sleep(1)

        # Get final count
        final_count = model_space.Count
        objects_created = final_count - initial_count
        print(f"   ✓ Created {objects_created} new objects")

        # Zoom extents
        try:
            acad.ZoomExtents()
            print("   ✓ Zoom extents completed")
        except Exception as e:
            print(f"   ✗ Zoom extents failed: {e}")

        # Summary
        print("\n=== SUMMARY ===")
        print(f"Total objects created: {objects_created}")
        print(f"Successful operations: {successful_operations}")
        print(f"Success rate: {len(successful_operations)}/11 operations")

        if len(successful_operations) > 0:
            print("🎉 SUCCESS! Some operations worked!")
            print("Working operations:")
            for op in successful_operations:
                print(f"   ✓ {op}")
            return True
        else:
            print("❌ No operations succeeded.")
            return False

    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        return False


if __name__ == "__main__":
    print("Starting slow and steady test...")
    print("This test adds delays between operations to avoid timing issues.")
    print("Please wait patiently...\n")

    success = test_autocad_slow_steady()

    if success:
        print("\n✅ Found working operations!")
        print("These can be used reliably in the MCP server.")
    else:
        print("\n❌ Need to investigate further.")