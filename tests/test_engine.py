"""Unit tests for randomoffice.engine.layout_geometry and loss_reporter."""

import pytest

from randomoffice.engine.layout_geometry import (
    BoundingBox,
    SlideLayoutType,
    classify_slide_layout,
    emu_to_inches,
    inches_to_emu,
    inches_to_pt,
    pt_to_inches,
    px_to_inches,
)
from randomoffice.engine.loss_reporter import DocumentLossAuditor


def test_bounding_box_aspect_ratio_and_scaling() -> None:
    # 4 x 2 inches box (aspect ratio 2.0)
    box = BoundingBox(x=1.0, y=1.0, width=4.0, height=2.0)
    assert box.aspect_ratio == 2.0

    # Scale to fit 2.0 x 2.0 bounds -> width becomes 2.0, height becomes 1.0 (aspect ratio 2.0 preserved)
    fitted = box.scale_to_fit(2.0, 2.0, keep_aspect_ratio=True)
    assert fitted.width == 2.0
    assert fitted.height == 1.0
    assert fitted.aspect_ratio == 2.0
    assert fitted.x == 1.0 and fitted.y == 1.0

    # Serialization roundtrip
    data = fitted.to_dict()
    restored = BoundingBox.from_dict(data)
    assert restored.width == fitted.width
    assert restored.height == fitted.height


def test_bounding_box_validations() -> None:
    with pytest.raises(ValueError, match="negative"):
        BoundingBox(0, 0, -1, 10)

    box = BoundingBox(0, 0, 10, 10)
    with pytest.raises(ValueError, match="strictly positive"):
        box.scale_to_fit(0, 5)


def test_unit_conversions() -> None:
    assert inches_to_emu(1.0) == 914_400
    assert emu_to_inches(914_400) == 1.0
    assert inches_to_pt(1.0) == 72.0
    assert pt_to_inches(72.0) == 1.0
    assert px_to_inches(96.0) == 1.0


def test_slide_layout_classification() -> None:
    assert classify_slide_layout(has_title=False, body_paragraphs_count=0, tables_count=0, images_count=0) == SlideLayoutType.BLANK
    assert classify_slide_layout(has_title=True, body_paragraphs_count=0, tables_count=0, images_count=0) == SlideLayoutType.TITLE_ONLY
    assert classify_slide_layout(has_title=True, body_paragraphs_count=1, tables_count=0, images_count=1) == SlideLayoutType.PICTURE_WITH_CAPTION
    assert classify_slide_layout(has_title=True, body_paragraphs_count=5, tables_count=0, images_count=0) == SlideLayoutType.TWO_COLUMN
    assert classify_slide_layout(has_title=True, body_paragraphs_count=2, tables_count=0, images_count=0) == SlideLayoutType.TITLE_AND_CONTENT


def test_document_loss_auditor() -> None:
    # Word audit
    word_items = [
        {"type": "paragraph", "text": "Heading 1", "level": "h1"},
        {"type": "table", "rows": [["A", "B"], ["C", "D"]]},
    ]
    word_loss = DocumentLossAuditor.audit_word_items(word_items)
    assert len(word_loss) >= 1
    assert word_loss[0].feature == "table_cell_shading_and_borders"

    # PPTX audit with overlapping images
    pptx_slides = [
        {"slide_number": 1, "images": ["data:image/png;base64,AAA", "data:image/png;base64,BBB"], "tables": [[["T1"]]]}
    ]
    pptx_loss = DocumentLossAuditor.audit_pptx_slides(pptx_slides)
    features = {entry.feature for entry in pptx_loss}
    assert "pptx_floating_image_coordinates" in features
    assert "pptx_table_geometry" in features

