"""
Test suite for CorelDRAWer-Skill generators.
Run with: python3 -m pytest tests/ -v
Or:        python3 tests/test_generators.py
"""

import json
import os
import sys
import tempfile
import unittest

# Add parent dir to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from generate_column import generate_svg, DEFAULT_DATA, PATTERNS
from generate_cross_section import generate_cross_section, DEMO_DATA


class TestColumnGenerator(unittest.TestCase):
    """Tests for generate_column.py — stratigraphic column SVG generator."""

    def test_default_data_generates(self):
        """Default 14-layer Zigui section should generate without error."""
        result = generate_svg(DEFAULT_DATA)
        self.assertIsInstance(result, str)
        self.assertIn('<svg', result)
        self.assertIn('cdr-background', result)
        self.assertIn('cdr-body', result)
        self.assertIn('cdr-legend', result)

    def test_single_layer(self):
        """Single layer should generate valid SVG."""
        data = {
            "title": "Test",
            "location": "",
            "layers": [{
                "erathem": "A", "system": "B", "series": "C",
                "formation": "D", "symbol": "X", "thick": 10,
                "descr": "test layer",
                "c": 0, "m": 0, "y": 0, "k": 10,
                "pattern": "sand", "grain": 3
            }]
        }
        result = generate_svg(data)
        self.assertIn('<svg', result)
        self.assertIn('layer_0', result)

    def test_thin_layers(self):
        """Very thin layers (<3.5m) should be clamped to min height."""
        data = {
            "title": "Thin",
            "layers": [
                {"erathem": "A", "system": "B", "series": "C", "formation": "D",
                 "symbol": "X", "thick": 0.1, "descr": "very thin",
                 "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "mud"}
            ]
        }
        result = generate_svg(data)
        self.assertIn('<svg', result)

    def test_all_patterns(self):
        """All 18 patterns should render without error."""
        for pname in PATTERNS:
            data = {
                "title": f"Test {pname}",
                "layers": [{
                    "erathem": "A", "system": "B", "series": "C",
                    "formation": pname, "symbol": "T",
                    "thick": 20, "descr": f"testing {pname}",
                    "c": 0, "m": 0, "y": 0, "k": 20,
                    "pattern": pname
                }]
            }
            result = generate_svg(data)
            self.assertIn('<svg', result, f"Pattern {pname} failed")

    def test_fossil_column_shows(self):
        """Fossil column should appear when fossil data present."""
        data = {
            "title": "Fossils",
            "layers": [{
                "erathem": "A", "system": "B", "series": "C",
                "formation": "D", "symbol": "X", "thick": 10,
                "descr": "with fossils",
                "c": 0, "m": 0, "y": 0, "k": 10,
                "pattern": "lime", "fossils": ["trilobite", "brachiopod"]
            }]
        }
        result = generate_svg(data)
        self.assertIn('fossil-icon', result)

    def test_structure_column_shows(self):
        """Structure column should appear when structure data present."""
        data = {
            "title": "Structures",
            "layers": [{
                "erathem": "A", "system": "B", "series": "C",
                "formation": "D", "symbol": "X", "thick": 10,
                "descr": "with structures",
                "c": 0, "m": 0, "y": 0, "k": 10,
                "pattern": "sand", "structures": ["cross_bed", "ripple"]
            }]
        }
        result = generate_svg(data)
        self.assertIn('structure-icon', result)

    def test_contact_symbols(self):
        """Contact symbols should appear when specified."""
        for contact in ["conformity", "disconformity", "unconformity"]:
            data = {
                "title": f"Contact {contact}",
                "layers": [
                    {"erathem": "A", "system": "B", "series": "C", "formation": "D",
                     "symbol": "X", "thick": 10, "descr": "top",
                     "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand",
                     "contact": contact},
                    {"erathem": "A", "system": "B", "series": "C", "formation": "E",
                     "symbol": "Y", "thick": 15, "descr": "bottom",
                     "c": 0, "m": 0, "y": 0, "k": 20, "pattern": "lime"}
                ]
            }
            result = generate_svg(data)
            self.assertIn('<svg', result)

    def test_output_to_file(self):
        """Should write SVG to file when output_path given."""
        with tempfile.NamedTemporaryFile(suffix='.svg', delete=False) as f:
            path = f.name
        try:
            result = generate_svg(DEFAULT_DATA, path)
            self.assertEqual(result, path)
            with open(path) as f:
                content = f.read()
                self.assertIn('<svg', content)
        finally:
            os.unlink(path)

    def test_empty_layers_raises(self):
        """Empty layers list should raise ValueError."""
        with self.assertRaises(ValueError):
            generate_svg({"title": "Empty", "layers": []})

    def test_zero_thickness_raises(self):
        """Zero total thickness should raise ValueError."""
        with self.assertRaises(ValueError):
            generate_svg({
                "title": "Zero",
                "layers": [{
                    "erathem": "A", "system": "B", "series": "C",
                    "formation": "D", "symbol": "X", "thick": 0,
                    "descr": "zero", "c": 0, "m": 0, "y": 0, "k": 0,
                    "pattern": "pure"
                }]
            })

    def test_age_annotation(self):
        """Age annotations should appear when specified."""
        data = {
            "title": "Ages",
            "layers": [{
                "erathem": "A", "system": "B", "series": "C",
                "formation": "D", "symbol": "X", "thick": 10,
                "descr": "with age", "age_ma": 541,
                "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"
            }]
        }
        result = generate_svg(data)
        self.assertIn('age-ma', result)


class TestCrossSectionGenerator(unittest.TestCase):
    """Tests for generate_cross_section.py."""

    def test_demo_data_generates(self):
        """Default 3-borehole demo should generate without error."""
        result = generate_cross_section(DEMO_DATA)
        self.assertIsInstance(result, str)
        self.assertIn('<svg', result)
        self.assertIn('cdr-background', result)
        self.assertIn('cdr-body', result)
        self.assertIn('surface-line', result)
        self.assertIn('layer-band', result)
        self.assertIn('fault-line', result)

    def test_two_boreholes(self):
        """Minimal 2-borehole section should work."""
        data = {
            "title": "2-BH Test",
            "boreholes": [
                {"id": "A", "x": 0, "elevation": 100, "depth": 30,
                 "layers": [{"formation": "F1", "thick": 15, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"},
                           {"formation": "F2", "thick": 15, "c": 0, "m": 0, "y": 0, "k": 30, "pattern": "lime"}]},
                {"id": "B", "x": 200, "elevation": 90, "depth": 25,
                 "layers": [{"formation": "F1", "thick": 10, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"},
                           {"formation": "F2", "thick": 15, "c": 0, "m": 0, "y": 0, "k": 30, "pattern": "lime"}]}
            ]
        }
        result = generate_cross_section(data)
        self.assertIn('<svg', result)
        self.assertIn('surface-line', result)

    def test_no_faults(self):
        """Section without faults should work."""
        data = {
            "title": "No Faults",
            "boreholes": [
                {"id": "A", "x": 0, "elevation": 100, "depth": 20,
                 "layers": [{"formation": "F1", "thick": 20, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"}]},
                {"id": "B", "x": 150, "elevation": 95, "depth": 20,
                 "layers": [{"formation": "F1", "thick": 20, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"}]}
            ],
            "faults": []
        }
        result = generate_cross_section(data)
        self.assertNotIn('fault-line', result)

    def test_vertical_exaggeration_clamped(self):
        """VE should be clamped to valid range."""
        data = {
            "title": "VE Test",
            "vertical_exaggeration": 50,  # should be clamped to 20
            "boreholes": [
                {"id": "A", "x": 0, "elevation": 100, "depth": 20,
                 "layers": [{"formation": "F1", "thick": 20, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"}]},
                {"id": "B", "x": 150, "elevation": 95, "depth": 20,
                 "layers": [{"formation": "F1", "thick": 20, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"}]}
            ]
        }
        result = generate_cross_section(data)
        self.assertIn('<svg', result)

    def test_output_to_file(self):
        """Should write SVG to file."""
        with tempfile.NamedTemporaryFile(suffix='.svg', delete=False) as f:
            path = f.name
        try:
            result = generate_cross_section(DEMO_DATA, path)
            self.assertEqual(result, path)
            with open(path) as f:
                self.assertIn('<svg', f.read())
        finally:
            os.unlink(path)


class TestVBAGenerator(unittest.TestCase):
    """Tests for cdr_com_auto.py VBA generator."""

    def test_vba_generates(self):
        """VBA code should generate for column data."""
        from cdr_com_auto import generate_vba_code
        vba = generate_vba_code(DEFAULT_DATA)
        self.assertIn('Public Sub DrawColumn()', vba)
        self.assertIn('End Sub', vba)
        self.assertIn('On Error GoTo ErrHandler', vba)
        self.assertIn('BeginCommandGroup', vba)
        self.assertIn('EndCommandGroup', vba)

    def test_vba_includes_fossils(self):
        """VBA should include fossil handling when data has fossils."""
        from cdr_com_auto import generate_vba_code
        data = {
            "title": "Test",
            "layers": [{
                "erathem": "A", "system": "B", "series": "C",
                "formation": "D", "symbol": "X", "thick": 10,
                "descr": "test", "fossils": ["trilobite"],
                "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"
            }]
        }
        vba = generate_vba_code(data)
        self.assertIn('trilobite', vba)


class TestDXF(unittest.TestCase):
    """Tests for dxf_export.py."""

    def test_dxf_generates(self):
        """DXF should produce valid DXF with layers."""
        from dxf_export import generate_dxf
        import tempfile, os
        with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as f:
            path = f.name
        try:
            generate_dxf(DEFAULT_DATA, path)
            with open(path) as f:
                content = f.read()
                self.assertIn("SECTION", content)
                self.assertIn("EOF", content)
        finally:
            os.unlink(path)

    def test_cross_section_two_bh(self):
        """Cross-section with fault should render correctly."""
        data = {
            "title": "Test",
            "boreholes": [
                {"id": "A", "x": 0, "elevation": 100, "depth": 20,
                 "layers": [{"formation": "F1", "thick": 20, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"}]},
                {"id": "B", "x": 100, "elevation": 90, "depth": 20,
                 "layers": [{"formation": "F1", "thick": 20, "c": 0, "m": 0, "y": 0, "k": 10, "pattern": "sand"}]}
            ],
            "faults": [{"x": 50, "dip": 60, "direction": "E", "type": "normal", "throw": 10}]
        }
        result = generate_cross_section(data)
        self.assertIn("fault-line", result)


class TestCLI(unittest.TestCase):
    """Tests for the coreldrawer.py unified CLI."""

    def test_cli_help(self):
        """CLI should not crash on --help."""
        import subprocess
        result = subprocess.run(
            [sys.executable, 'coreldrawer.py', '--help'],
            capture_output=True, text=True,
            cwd=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        )
        self.assertEqual(result.returncode, 0)
        self.assertIn('usage:', result.stdout.lower())


if __name__ == '__main__':
    unittest.main()
