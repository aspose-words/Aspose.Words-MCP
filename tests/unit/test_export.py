import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

from core.export import build_pdf_opts


def test_build_pdf_opts_default():
    """Test default PDF options with no custom_properties_export."""
    opts = build_pdf_opts({})
    assert opts is not None
    assert opts.custom_properties_export == aw.saving.PdfCustomPropertiesExport.NONE


def test_build_pdf_opts_custom_properties_export_none():
    """Test custom_properties_export=none."""
    opts = build_pdf_opts({'custom_properties_export': 'none'})
    assert opts.custom_properties_export == aw.saving.PdfCustomPropertiesExport.NONE


def test_build_pdf_opts_custom_properties_export_standard():
    """Test custom_properties_export=standard."""
    opts = build_pdf_opts({'custom_properties_export': 'standard'})
    assert opts.custom_properties_export == aw.saving.PdfCustomPropertiesExport.STANDARD


def test_build_pdf_opts_custom_properties_export_metadata():
    """Test custom_properties_export=metadata."""
    opts = build_pdf_opts({'custom_properties_export': 'metadata'})
    assert opts.custom_properties_export == aw.saving.PdfCustomPropertiesExport.METADATA


def test_build_pdf_opts_custom_properties_export_invalid():
    """Test invalid custom_properties_export raises ValueError."""
    with pytest.raises(ValueError, match='custom_properties_export must be one of'):
        build_pdf_opts({'custom_properties_export': 'invalid'})


def test_build_pdf_opts_custom_properties_export_case_insensitive():
    """Test custom_properties_export is case-insensitive."""
    opts = build_pdf_opts({'custom_properties_export': 'METADATA'})
    assert opts.custom_properties_export == aw.saving.PdfCustomPropertiesExport.METADATA


def test_build_pdf_opts_compliance_and_custom_properties_export():
    """Test compliance and custom_properties_export work together."""
    opts = build_pdf_opts(
        {
            'compliance': 'PDF_A1B',
            'custom_properties_export': 'metadata',
        }
    )
    assert opts.compliance == aw.saving.PdfCompliance.PDF_A1B
    assert opts.custom_properties_export == aw.saving.PdfCustomPropertiesExport.METADATA
