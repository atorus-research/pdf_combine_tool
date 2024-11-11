import pytest
from unittest.mock import Mock, patch
import os
import pandas as pd
import fitz
import tempfile
import shutil
from src.pdf_util import PDFUtility
from src.pdf_compiler import PDFCompiler
import win32com.client


@pytest.fixture
def mock_gui():
    """Fixture to provide a mock GUI instance"""
    gui = Mock()
    gui.entry_var1 = Mock(get=Mock(return_value="/test/path"))
    gui.entry_var2 = Mock(get=Mock(return_value="/test/metadata.csv"))
    gui.entry_var5 = Mock(get=Mock(return_value="test_password"))
    gui.final_run_var = Mock(get=Mock(return_value=False))
    gui.logger = Mock()
    return gui


@pytest.fixture
def pdf_util(mock_gui):
    """Fixture to provide a PDFUtility instance"""
    return PDFUtility(mock_gui)


@pytest.fixture
def mock_word():
    """Fixture to mock Word COM automation"""
    with patch('win32com.client.gencache.EnsureDispatch') as mock:
        word_app = Mock()
        word_app.Documents = Mock()
        word_app.Documents.Open = Mock()
        word_app.Quit = Mock()
        mock.return_value = word_app
        yield mock


@pytest.fixture
def test_metadata_content():
    """Fixture to provide test metadata content"""
    return '''TLF,Title3,Title4,Title5,ProgName,Seq,OutputName,Order
F,Figure 1,Test Plot,Safety Pop,g_test,1,F_1.1,1
F,Figure 2,Another Plot,Full Pop,g_test,2,F_1.2,2'''


@pytest.fixture
def test_metadata_file(test_dir, test_metadata_content):
    """Fixture to provide a test metadata file"""
    metadata_path = os.path.join(test_dir, "test_metadata.csv")
    with open(metadata_path, "w") as f:
        f.write(test_metadata_content)
    return metadata_path


@pytest.fixture
def test_dir():
    """Fixture to provide a temporary directory"""
    temp_dir = tempfile.mkdtemp()
    yield temp_dir
    shutil.rmtree(temp_dir, ignore_errors=True)


@pytest.fixture
def test_files(test_dir):
    """Fixture to create and provide test files"""
    # Create test PDF
    pdf_path = os.path.join(test_dir, "test.pdf")
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((50, 50), "Test PDF")
    doc.save(pdf_path)
    doc.close()

    # Create test RTF
    rtf_path = os.path.join(test_dir, "test.rtf")
    with open(rtf_path, "w") as f:
        f.write(r"{\rtf1\ansi\Test RTF}")

    # Create test metadata CSV
    csv_path = os.path.join(test_dir, "test_metadata.csv")
    test_data = {
        'TLF': ['F', 'F'],
        'Title3': ['Test1', 'Test2'],
        'Title4': ['Description1', 'Description2'],
        'Title5': ['Pop1', 'Pop2'],
        'ProgName': ['prog1', 'prog2'],
        'Seq': [1, 2],
        'OutputName': ['F_14.1', 'F_14.2'],
        'Order': [1, 2]
    }
    pd.DataFrame(test_data).to_csv(csv_path, index=False)

    return pdf_path, rtf_path, csv_path


class TestPDFUtility:
    @pytest.mark.integration
    def test_meta_data_to_dict(self, test_metadata_file):
        """Test metadata parsing with and without population"""
        # Test with population included
        result = PDFUtility.meta_data_to_dict(test_metadata_file, title_sep="-", add_popul=True)
        assert len(result) == 2
        assert all(isinstance(k, str) for k in result.keys())
        assert all(isinstance(v, str) for v in result.values())
        assert any('Pop' in v for v in result.values())

        # Test without population
        result = PDFUtility.meta_data_to_dict(test_metadata_file, title_sep="-", add_popul=False)
        assert len(result) == 2
        assert not any('Pop' in v for v in result.values())

    def test_get_tlf_list(self, test_files):
        """Test TLF list extraction from metadata"""
        _, _, csv_path = test_files
        files, count = PDFUtility.get_tlf_list(csv_path)

        assert count == 2
        assert len(files) == 2
        assert all(f.endswith('.rtf') for f in files)
        assert all('F_' in f for f in files)

    def test_combine_pdfs_simple(self, pdf_util, test_dir):
        """Test PDF combination without TOC"""
        # Create test PDFs
        test_pdfs = []
        for i in range(2):
            pdf_path = os.path.join(test_dir, f'test{i}.pdf')
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 50), f"Test PDF {i}")
            doc.save(pdf_path)
            doc.close()
            test_pdfs.append(pdf_path)

        # Test without password
        output_path = os.path.join(test_dir, 'combined.pdf')
        result = pdf_util.combine_pdfs_simple(
            test_pdfs,
            output_path,
            use_password=False
        )

        assert result is True
        assert os.path.exists(output_path)

        with fitz.open(output_path) as doc:
            assert doc.page_count == 2

        # Test with password protection
        output_path_protected = os.path.join(test_dir, 'combined_protected.pdf')
        result = pdf_util.combine_pdfs_simple(
            test_pdfs,
            output_path_protected,
            use_password=True,
            password="test_password"
        )

        assert result is True
        assert os.path.exists(output_path_protected)

    @pytest.mark.parametrize("file_count", [1, 5, 10])
    def test_combine_multiple_pdfs(self, pdf_util, test_dir, file_count):
        """Test combining different numbers of PDFs"""
        test_pdfs = []
        for i in range(file_count):
            pdf_path = os.path.join(test_dir, f'test{i}.pdf')
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 50), f"Test PDF {i}")
            doc.save(pdf_path)
            doc.close()
            test_pdfs.append(pdf_path)

        output_path = os.path.join(test_dir, 'combined.pdf')
        result = pdf_util.combine_pdfs_simple(test_pdfs, output_path)

        assert result is True
        with fitz.open(output_path) as doc:
            assert doc.page_count == file_count

    @pytest.mark.win32
    def test_rtf_conversion(self, mock_word, test_dir, pdf_util):
        """Test RTF to PDF conversion using Word automation"""
        # Create test RTF
        rtf_path = os.path.join(test_dir, "test.rtf")
        with open(rtf_path, "w") as f:
            f.write(r"{\rtf1\ansi\Test RTF}")

        pdf_util.rtf_file_to_pdf(
            file_name="test.rtf",
            input_dir=test_dir,
            output_dir=test_dir,
            pause_time=0.1
        )

        # Verify Word automation calls
        mock_word.assert_called_once()
        word_app = mock_word.return_value
        word_app.Documents.Open.assert_called_once()
        word_app.Quit.assert_called_once()

    @pytest.mark.benchmark
    def test_pdf_combine_performance(self, benchmark, pdf_util, test_dir):
        """Test PDF combination performance"""
        # Create test PDFs
        test_pdfs = []
        for i in range(3):
            pdf_path = os.path.join(test_dir, f'test{i}.pdf')
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 50), f"Test PDF {i}")
            doc.save(pdf_path)
            doc.close()
            test_pdfs.append(pdf_path)

        output_path = os.path.join(test_dir, 'combined.pdf')

        # Run benchmark
        result = benchmark(
            pdf_util.combine_pdfs_simple,
            test_pdfs,
            output_path,
            use_password=False
        )

        assert result is True
        assert os.path.exists(output_path)

    @pytest.mark.parametrize("file_count", [1, 5, 10])
    def test_combine_multiple_pdfs(self, pdf_util, test_dir, file_count):
        """Test combining different numbers of PDFs"""
        test_pdfs = []
        for i in range(file_count):
            pdf_path = os.path.join(test_dir, f'test{i}.pdf')
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((50, 50), f"Test PDF {i}")
            doc.save(pdf_path)
            doc.close()
            test_pdfs.append(pdf_path)

        output_path = os.path.join(test_dir, 'combined.pdf')
        result = pdf_util.combine_pdfs_simple(test_pdfs, output_path)

        assert result is True
        with fitz.open(output_path) as doc:
            assert doc.page_count == file_count


class TestPDFCompiler:
    @pytest.fixture
    def pdf_compiler(self, mock_gui):
        """Fixture to provide a PDFCompiler instance"""
        util_mock = Mock()
        return PDFCompiler(mock_gui, util_mock)

    def test_get_toc_page_numb(self, test_dir):
        """Test TOC page number extraction"""
        pdf_path = os.path.join(test_dir, 'test_toc.pdf')
        doc = fitz.open()
        for i in range(3):
            page = doc.new_page()
            page.insert_text((50, 50), f"Test Page {i}")
        doc.save(pdf_path)
        doc.close()

        page_count = PDFCompiler.get_toc_page_numb(pdf_path)
        assert page_count == 3

    def test_update_toc_pages(self, test_dir):
        """Test TOC page number updating"""
        test_content = (
            "Title 1 *page:1\n"
            "Title 2 *page:2\n"
            "Title 3 *page:3\n"
        )

        input_file = os.path.join(test_dir, 'test_toc.txt')
        with open(input_file, 'w', encoding='latin-1') as f:
            f.write(test_content)

        updated_content = PDFCompiler.update_toc_pages(
            input_file=input_file,
            page_char="*page:",
            w_page=50,
            page_numb_to_add=2
        )

        assert "*page:3" in updated_content
        assert "*page:4" in updated_content
        assert "*page:5" in updated_content

    @pytest.mark.parametrize("page_width,expected", [
        (50, True),  # Normal width
        (10, True),  # Very narrow
        (200, True)  # Very wide
    ])
    def test_update_toc_pages_different_widths(self, test_dir, page_width, expected):
        """Test TOC page updating with different page widths"""
        test_content = "Title 1 *page:1\n"
        input_file = os.path.join(test_dir, 'test_toc.txt')
        with open(input_file, 'w', encoding='latin-1') as f:
            f.write(test_content)

        updated_content = PDFCompiler.update_toc_pages(
            input_file=input_file,
            page_char="*page:",
            w_page=page_width,
            page_numb_to_add=1
        )

        assert bool(updated_content) is expected

    @pytest.mark.parametrize("test_input,expected", [
        ("normal.pdf", True),
        ("missing.pdf", False)
    ])
    def test_file_existence_handling(self, test_dir, test_input, expected):
        """Test handling of different file inputs"""
        if expected:
            # Create test file if it should exist
            pdf_path = os.path.join(test_dir, test_input)
            doc = fitz.open()
            doc.new_page()
            doc.save(pdf_path)
            doc.close()

        assert os.path.exists(os.path.join(test_dir, test_input)) == expected


if __name__ == '__main__':
    pytest.main(['-v'])