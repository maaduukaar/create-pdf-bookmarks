# -*- coding: utf-8 -*-
"""
Tests for PDF bookmark generator.
These tests verify that bookmark extraction, hierarchy building, coordinate resolution,
and PDF embedding work perfectly and do not regress.
"""

import os
import sys
import json
import shutil
import pytest
import fitz

# Resolve paths to ensure bookmarks.py can be imported
TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(TESTS_DIR)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import bookmarks


def test_reference_bookmark_extraction():
    """
    Test that bookmarks extracted from the reference PDF document
    fully match the reference JSON file in content, levels, pages,
    nesting structure, and exact coordinate positions.
    """
    pdf_path = os.path.join(TESTS_DIR, "reference.pdf")
    ref_json_path = os.path.join(TESTS_DIR, "reference_bookmarks.json")
    
    assert os.path.exists(pdf_path), f"Reference PDF not found at {pdf_path}"
    assert os.path.exists(ref_json_path), f"Reference JSON not found at {ref_json_path}"
    
    # 1. Extract TOC entries from pages 5 and 6
    entries = bookmarks.extract_toc_from_pdf_pages(
        pdf_path, 
        page_numbers=[5, 6], 
        show_output=False, 
        use_numbering=False
    )
    assert entries, "Failed to extract any TOC entries from reference.pdf"
    
    # 2. Find exact coordinates for entries
    entries_with_coords = bookmarks.find_exact_coordinates(
        pdf_path, 
        entries, 
        show_output=False
    )
    
    # 3. Build bookmark tree
    generated_tree = bookmarks.build_bookmark_tree(entries_with_coords)
    
    # 4. Load the reference JSON
    with open(ref_json_path, "r", encoding="utf-8") as f:
        reference_tree = json.load(f)
        
    # 5. Assert equality of trees
    assert len(generated_tree) == len(reference_tree), (
        f"Top-level bookmarks count mismatch: "
        f"got {len(generated_tree)}, expected {len(reference_tree)}"
    )
    
    def assert_nodes_equal(node_gen, node_ref, path=""):
        # Check title
        assert node_gen["title"] == node_ref["title"], f"Title mismatch at {path}: {node_gen['title']} != {node_ref['title']}"
        
        # Check destination (page and type/coordinates)
        assert len(node_gen["dest"]) == len(node_ref["dest"]), f"Destination length mismatch at {path} for {node_gen['title']}"
        assert node_gen["dest"][0] == node_ref["dest"][0], f"Page mismatch at {path} for {node_gen['title']}"
        assert node_gen["dest"][1] == node_ref["dest"][1], f"Destination type mismatch at {path} for {node_gen['title']}"
        
        if len(node_gen["dest"]) > 2:
            # Compare coordinates with a small delta to prevent tiny float rounding errors
            for i in range(2, len(node_gen["dest"])):
                val_gen = node_gen["dest"][i]
                val_ref = node_ref["dest"][i]
                if isinstance(val_gen, (int, float)) and isinstance(val_ref, (int, float)):
                    assert val_gen == pytest.approx(val_ref, abs=1e-4), f"Coordinate mismatch at index {i} at {path} for {node_gen['title']}"
                else:
                    assert val_gen == val_ref, f"Destination value mismatch at index {i} at {path} for {node_gen['title']}"
        
        # Check formatting attributes
        assert node_gen.get("bold") == node_ref.get("bold"), f"Bold mismatch at {path} for {node_gen['title']}"
        assert node_gen.get("italic") == node_ref.get("italic"), f"Italic mismatch at {path} for {node_gen['title']}"
        assert node_gen.get("color") == node_ref.get("color"), f"Color mismatch at {path} for {node_gen['title']}"
        
        # Check children recursively
        children_gen = node_gen.get("children", [])
        children_ref = node_ref.get("children", [])
        assert len(children_gen) == len(children_ref), f"Children count mismatch at {path} for {node_gen['title']}"
        
        for i, (child_gen, child_ref) in enumerate(zip(children_gen, children_ref)):
            assert_nodes_equal(child_gen, child_ref, f"{path} -> child[{i}]")

    for i, (gen, ref) in enumerate(zip(generated_tree, reference_tree)):
        assert_nodes_equal(gen, ref, f"root[{i}]")


def test_pdf_embedding(tmp_path):
    """
    Test copying the reference PDF, embedding bookmarks into it,
    saving the result, and checking that the bookmarks can be correctly retrieved
    from the PDF file with correct page links and hierarchy levels.
    """
    # Paths
    ref_pdf = os.path.join(TESTS_DIR, "reference.pdf")
    ref_json = os.path.join(TESTS_DIR, "reference_bookmarks.json")
    
    # 1. Copy reference PDF to a temporary file
    temp_pdf_path = os.path.join(tmp_path, "reference.pdf")
    shutil.copyfile(ref_pdf, temp_pdf_path)
    
    # 2. Copy reference JSON to a temporary file
    temp_json_path = os.path.join(tmp_path, "reference_bookmarks.json")
    shutil.copyfile(ref_json, temp_json_path)
    
    # 3. Embed bookmarks
    success = bookmarks.embed_bookmarks_to_pdf(
        temp_pdf_path, 
        temp_json_path, 
        show_output=False
    )
    assert success, "embed_bookmarks_to_pdf returned False"
    
    # The output should be saved alongside reference.pdf as reference_with_bookmarks.pdf
    expected_output_path = os.path.join(tmp_path, "reference_with_bookmarks.pdf")
    assert os.path.exists(expected_output_path), f"Output PDF file not created at {expected_output_path}"
    
    # 4. Open the generated PDF and read back the table of contents (TOC)
    doc = fitz.open(expected_output_path)
    toc = doc.get_toc()
    doc.close()
    
    # 5. Load original json to count expected total bookmarks
    with open(ref_json, "r", encoding="utf-8") as f:
        ref_tree = json.load(f)
        
    def count_nodes(nodes):
        count = len(nodes)
        for node in nodes:
            count += count_nodes(node.get("children", []))
        return count
        
    expected_count = count_nodes(ref_tree)
    assert len(toc) == expected_count, f"TOC count mismatch: got {len(toc)}, expected {expected_count}"
    
    # Verify the first few bookmarks in TOC to ensure they match our expectation
    # PyMuPDF TOC format is: [level, title, page_num] where page_num is 1-based (but custom link dict can also be present)
    # The first bookmark should have level 1 and page 7
    assert toc[0][0] == 1
    assert toc[0][2] == 7
    # The second bookmark should have level 2 and page 7
    assert toc[1][0] == 2
    assert toc[1][2] == 7


@pytest.mark.parametrize("input_text, expected_title, expected_level, expected_page", [
    ("1.2.3 Название раздела ...... 42", "1.2.3. Название раздела", 3, 42),
    ("1.2.3 Название раздела    42", "1.2.3. Название раздела", 3, 42),
    ("1. Назначение программы 7", "1. Назначение программы", 1, 7),
    ("1.1 Наименование программы 7", "1.1. Наименование программы", 2, 7),
    ("2.1.1 Требования к оборудованию 13", "2.1.1. Требования к оборудованию", 3, 13),
    ("Введение ...... 5", "Введение", 1, 5),
    ("Перечень терминов 3", "Перечень терминов", 1, 3),
])
def test_parse_toc_line_from_pdf(input_text, expected_title, expected_level, expected_page):
    """Unit test for parsing visual TOC lines from PDF."""
    title, level, page = bookmarks.parse_toc_line_from_pdf(input_text)
    assert title == expected_title
    assert level == expected_level
    assert page == expected_page


@pytest.mark.parametrize("invalid_text", [
    "",
    "   ",
    "Просто какой-то текст без номера страницы",
    "1.2.3 Раздел без страницы",
    "123",
])
def test_parse_toc_line_from_pdf_invalid(invalid_text):
    """Unit test that invalid strings return (None, None, None)."""
    title, level, page = bookmarks.parse_toc_line_from_pdf(invalid_text)
    assert title is None
    assert level is None
    assert page is None


def test_determine_level_by_numbering():
    """Unit test for determining level purely by heading numbering system."""
    entries = [
        {"title": "Перечень терминов", "page": 3},
        {"title": "1. Назначение программы", "page": 10},
        {"title": "1.1 Наименование программы", "page": 10},
        {"title": "1.2. Функции программы", "page": 10},
        {"title": "2. Условия выполнения программы", "page": 13},
        {"title": "2.1.1 Требования к оборудованию", "page": 13},
    ]
    
    result = bookmarks.determine_level_by_numbering(entries)
    
    assert result[0]["level"] == 1  # No numbering
    assert result[1]["level"] == 1  # "1."
    assert result[2]["level"] == 2  # "1.1"
    assert result[3]["level"] == 2  # "1.2."
    assert result[4]["level"] == 1  # "2."
    assert result[5]["level"] == 3  # "2.1.1"


@pytest.mark.parametrize("page_str, expected_list", [
    ("2", [2]),
    ("2,3", [2, 3]),
    ("2-4", [2, 3, 4]),
    ("2,3,5-7", [2, 3, 5, 6, 7]),
    (" 2,  3-5 ", [2, 3, 4, 5]),
])
def test_parse_page_range(page_str, expected_list):
    """Unit test for parsing string page ranges."""
    assert bookmarks.parse_page_range(page_str) == expected_list


@pytest.mark.parametrize("invalid_page_str", [
    "abc",
    "2-a",
    "2-",
    "-3",
    "2,,3",
])
def test_parse_page_range_invalid(invalid_page_str):
    """Unit test that invalid page ranges raise ValueError."""
    with pytest.raises(ValueError):
        bookmarks.parse_page_range(invalid_page_str)


@pytest.mark.parametrize("docx_line, expected_title, expected_level, expected_page", [
    ("1. Назначение программы 7", "1 Назначение программы", 1, 7),
    ("1.1. Наименование программы 7", "1.1 Наименование программы", 2, 7),
    ("3.4.2.1. Название раздела 69", "3.4.2.1 Название раздела", 4, 69),
])
def test_parse_toc_line(docx_line, expected_title, expected_level, expected_page):
    """Unit test for parsing standard DOCX TOC lines."""
    title, level, page = bookmarks.parse_toc_line(docx_line)
    assert title == expected_title
    assert level == expected_level
    assert page == expected_page


def test_generate_search_variants():
    """Unit test for generating coordinate search variants for PDF pages."""
    # Test with standard dotted title containing more than 3 words in name
    variants = bookmarks.generate_search_variants("1.2.3. Введение в сложную систему")
    
    assert "1.2.3. Введение в сложную систему" in variants
    assert "1.2.3 Введение в сложную систему" in variants
    assert "Введение в сложную систему" in variants
    assert "Введение в сложную" in variants
    
    # Test with title without dots
    variants_nodot = bookmarks.generate_search_variants("1.2 Введение")
    assert "1.2 Введение" in variants_nodot
    assert "1.2. Введение" in variants_nodot
    assert "Введение" in variants_nodot
