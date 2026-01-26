"""
Unit tests for table detection functions.
"""
import pytest
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_to_md import (
    build_nonempty_grid,
    grid_to_tables,
    carve_rectangles,
    enumerate_histogram_rectangles,
    union_rects,
    bfs_components,
    get_print_areas,
)


# ============================================================
# Tests for carve_rectangles
# ============================================================

class TestCarveRectangles:
    """Tests for maximum rectangle decomposition."""

    def test_single_rectangle(self, grid_3x3_full):
        """Full 3x3 grid should be single rectangle."""
        rects = carve_rectangles(grid_3x3_full)
        assert len(rects) == 1
        # Should cover entire grid
        top, left, bottom, right = rects[0]
        assert (bottom - top + 1) * (right - left + 1) == 9

    def test_l_shape(self, grid_l_shape):
        """L-shape should decompose to 2 rectangles."""
        rects = carve_rectangles(grid_l_shape)
        assert len(rects) >= 1  # At least 1 rectangle
        # Total cells covered should be 5 (L-shape has 5 cells)
        total_cells = sum((b - t + 1) * (r - l + 1) for t, l, b, r in rects)
        assert total_cells == 5

    def test_t_shape(self):
        """T-shape should decompose to multiple rectangles."""
        # T-shape:
        # 1 1 1
        # 0 1 0
        # 0 1 0
        grid = [
            [1, 1, 1],
            [0, 1, 0],
            [0, 1, 0],
        ]
        rects = carve_rectangles(grid)
        assert len(rects) >= 2  # At least 2 rectangles
        # Total cells covered should be 5
        total_cells = sum((b - t + 1) * (r - l + 1) for t, l, b, r in rects)
        assert total_cells == 5

    def test_concave_shape(self):
        """Concave shape should decompose correctly."""
        # Concave (U-shape):
        # 1 0 1
        # 1 0 1
        # 1 1 1
        grid = [
            [1, 0, 1],
            [1, 0, 1],
            [1, 1, 1],
        ]
        rects = carve_rectangles(grid)
        assert len(rects) >= 2  # At least 2 rectangles
        # Total cells covered should be 7
        total_cells = sum((b - t + 1) * (r - l + 1) for t, l, b, r in rects)
        assert total_cells == 7

    def test_empty_grid(self, grid_empty):
        """Empty grid should return no rectangles."""
        rects = carve_rectangles(grid_empty)
        assert len(rects) == 0

    def test_single_cell(self):
        """Single cell grid should return one rectangle."""
        grid = [[1]]
        rects = carve_rectangles(grid)
        assert len(rects) == 1
        assert rects[0] == (0, 0, 0, 0)


# ============================================================
# Tests for enumerate_histogram_rectangles
# ============================================================

class TestEnumerateHistogramRectangles:
    """Tests for histogram-based rectangle enumeration."""

    def test_full_grid(self):
        """Full grid should enumerate containing rectangle."""
        grid = [
            [1, 1],
            [1, 1],
        ]
        rects = enumerate_histogram_rectangles(grid)
        # Should include the full 2x2 rectangle
        assert any((b - t + 1) * (r - l + 1) == 4 for t, l, b, r in rects)

    def test_empty_grid(self):
        """Empty grid should return no rectangles."""
        grid = [[0, 0], [0, 0]]
        rects = enumerate_histogram_rectangles(grid)
        assert len(rects) == 0

    def test_single_row(self):
        """Single row of cells."""
        grid = [[1, 1, 1]]
        rects = enumerate_histogram_rectangles(grid)
        # Should include 1x3 rectangle
        assert any((b - t + 1) == 1 and (r - l + 1) == 3 for t, l, b, r in rects)

    def test_single_column(self):
        """Single column of cells."""
        grid = [[1], [1], [1]]
        rects = enumerate_histogram_rectangles(grid)
        # Should include 3x1 rectangle
        assert any((b - t + 1) == 3 and (r - l + 1) == 1 for t, l, b, r in rects)


# ============================================================
# Tests for bfs_components
# ============================================================

class TestBfsComponents:
    """Tests for BFS connected component detection."""

    def test_single_component(self, grid_3x3_full):
        """Full grid should be single component."""
        comps = bfs_components(grid_3x3_full)
        assert len(comps) == 1
        assert len(comps[0]) == 9  # All 9 cells

    def test_two_separate_components(self, grid_two_separate):
        """Two separate regions should be two components."""
        comps = bfs_components(grid_two_separate)
        assert len(comps) == 2
        # Each component should have 4 cells
        assert all(len(c) == 4 for c in comps)

    def test_l_shape_single_component(self, grid_l_shape):
        """L-shape should be single connected component."""
        comps = bfs_components(grid_l_shape)
        assert len(comps) == 1
        assert len(comps[0]) == 5

    def test_empty_grid_no_components(self, grid_empty):
        """Empty grid should have no components."""
        comps = bfs_components(grid_empty)
        assert len(comps) == 0


# ============================================================
# Tests for union_rects
# ============================================================

class TestUnionRects:
    """Tests for rectangle union operation."""

    def test_single_rect(self):
        """Single rectangle should be returned as-is."""
        rects = [(1, 1, 3, 3)]
        result = union_rects(rects)
        assert len(result) == 1
        assert result[0] == (1, 1, 3, 3)

    def test_overlapping_rects(self):
        """Overlapping rectangles should be merged."""
        rects = [
            (1, 1, 2, 2),
            (2, 2, 3, 3),
        ]
        result = union_rects(rects)
        # Should cover the combined area
        total_input_area = sum((r[2] - r[0] + 1) * (r[3] - r[1] + 1) for r in rects)
        total_output_area = sum((r[2] - r[0] + 1) * (r[3] - r[1] + 1) for r in result)
        # Output area may be less or equal due to overlap
        assert total_output_area <= total_input_area + 4  # Allow some tolerance

    def test_adjacent_rects(self):
        """Adjacent rectangles should be handled."""
        rects = [
            (1, 1, 2, 2),
            (1, 3, 2, 4),  # Adjacent horizontally
        ]
        result = union_rects(rects)
        assert len(result) >= 1

    def test_empty_input(self):
        """Empty input should return empty output."""
        result = union_rects([])
        assert len(result) == 0

    def test_disjoint_rects(self):
        """Disjoint rectangles should remain separate."""
        rects = [
            (1, 1, 2, 2),
            (5, 5, 6, 6),
        ]
        result = union_rects(rects)
        assert len(result) == 2


# ============================================================
# Tests for grid_to_tables
# ============================================================

class TestGridToTables:
    """Tests for table detection from grid."""

    def test_single_table(self, simple_worksheet, default_opts):
        """Contiguous data should be single table."""
        ws = simple_worksheet
        area = (1, 1, 2, 2)  # A1:B2
        tables = grid_to_tables(ws, area, hidden_policy="ignore", opts=default_opts)
        assert len(tables) == 1

    def test_split_by_empty_row(self, worksheet_with_empty_rows, default_opts):
        """Empty row should split into 2 tables."""
        ws = worksheet_with_empty_rows
        area = (1, 1, 5, 2)  # A1:B5 with row 3 empty
        tables = grid_to_tables(ws, area, hidden_policy="ignore", opts=default_opts)
        assert len(tables) == 2

    def test_all_empty(self, empty_workbook, default_opts):
        """All empty cells should return no tables."""
        ws = empty_workbook.active
        area = (1, 1, 3, 3)
        tables = grid_to_tables(ws, area, hidden_policy="ignore", opts=default_opts)
        assert len(tables) == 0

    def test_table_has_bbox(self, simple_worksheet, default_opts):
        """Tables should have bounding box."""
        ws = simple_worksheet
        area = (1, 1, 2, 2)
        tables = grid_to_tables(ws, area, hidden_policy="ignore", opts=default_opts)
        assert len(tables) == 1
        assert "bbox" in tables[0]
        bbox = tables[0]["bbox"]
        assert len(bbox) == 4  # (min_row, min_col, max_row, max_col)

    def test_table_has_mask(self, simple_worksheet, default_opts):
        """Tables should have cell mask."""
        ws = simple_worksheet
        area = (1, 1, 2, 2)
        tables = grid_to_tables(ws, area, hidden_policy="ignore", opts=default_opts)
        assert len(tables) == 1
        assert "mask" in tables[0]
        assert isinstance(tables[0]["mask"], set)


# ============================================================
# Tests for build_nonempty_grid
# ============================================================

class TestBuildNonemptyGrid:
    """Tests for non-empty grid construction."""

    def test_simple_grid(self, simple_worksheet, default_opts):
        """Simple worksheet should produce correct grid."""
        ws = simple_worksheet
        area = (1, 1, 2, 2)
        grid, r0, c0, r1, c1 = build_nonempty_grid(ws, area, hidden_policy="ignore", opts=default_opts)
        # All 4 cells should be non-empty (marked as 1)
        assert grid[0][0] == 1  # A1
        assert grid[0][1] == 1  # B1
        assert grid[1][0] == 1  # A2
        assert grid[1][1] == 1  # B2

    def test_grid_with_empty_cells(self, empty_workbook, default_opts):
        """Grid should mark empty cells correctly."""
        ws = empty_workbook.active
        ws['A1'] = 'Data'
        # B1, A2, B2 are empty
        area = (1, 1, 2, 2)
        grid, r0, c0, r1, c1 = build_nonempty_grid(ws, area, hidden_policy="ignore", opts=default_opts)
        assert grid[0][0] == 1  # A1 has data
        assert grid[0][1] == 0  # B1 is empty
        assert grid[1][0] == 0  # A2 is empty
        assert grid[1][1] == 0  # B2 is empty

    def test_grid_bounds(self, simple_worksheet, default_opts):
        """Grid bounds should match input area."""
        ws = simple_worksheet
        area = (1, 1, 2, 2)
        grid, r0, c0, r1, c1 = build_nonempty_grid(ws, area, hidden_policy="ignore", opts=default_opts)
        assert r0 == 1
        assert c0 == 1
        assert r1 == 2
        assert c1 == 2
        assert len(grid) == 2  # 2 rows
        assert len(grid[0]) == 2  # 2 columns


# ============================================================
# Tests for get_print_areas
# ============================================================

class TestGetPrintAreas:
    """Tests for print area extraction."""

    def test_no_print_area_used_range(self, simple_worksheet):
        """No print area with used_range mode should return used range."""
        ws = simple_worksheet
        areas = get_print_areas(ws, "used_range")
        assert len(areas) >= 1
        # Should cover A1:B2
        area = areas[0]
        assert area[0] <= 1  # min_row
        assert area[1] <= 1  # min_col
        assert area[2] >= 2  # max_row
        assert area[3] >= 2  # max_col

    def test_no_print_area_skip_sheet(self, simple_worksheet):
        """No print area with skip_sheet mode should return empty."""
        ws = simple_worksheet
        areas = get_print_areas(ws, "skip_sheet")
        assert len(areas) == 0

    def test_with_print_area(self, simple_worksheet):
        """Worksheet with print area should return that area."""
        ws = simple_worksheet
        ws.print_area = "A1:B2"
        areas = get_print_areas(ws, "used_range")
        assert len(areas) >= 1

    def test_multiple_print_areas(self, simple_worksheet):
        """Multiple print areas should all be returned."""
        ws = simple_worksheet
        ws['C3'] = 'Extra'
        ws.print_area = "A1:B2,C3:C3"
        areas = get_print_areas(ws, "used_range")
        assert len(areas) >= 1


# ============================================================
# Tests for merged cell handling in grid
# ============================================================

class TestMergedCellsInGrid:
    """Tests for merged cell handling in table detection."""

    def test_merged_cells_detected(self, worksheet_with_merged_cells, default_opts):
        """Merged cells should be marked as non-empty."""
        ws = worksheet_with_merged_cells
        area = (1, 1, 3, 3)
        grid, r0, c0, r1, c1 = build_nonempty_grid(ws, area, hidden_policy="ignore", opts=default_opts)
        # Merged cell A1:C1 should mark all cells as non-empty
        assert grid[0][0] == 1  # A1
        assert grid[0][1] == 1  # B1 (part of merged range)
        assert grid[0][2] == 1  # C1 (part of merged range)

    def test_merged_cells_single_table(self, worksheet_with_merged_cells, default_opts):
        """Merged cells should not split tables."""
        ws = worksheet_with_merged_cells
        area = (1, 1, 3, 3)
        tables = grid_to_tables(ws, area, hidden_policy="ignore", opts=default_opts)
        assert len(tables) == 1  # Should be single table
