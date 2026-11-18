"""覆盖会改变原文或影响自动交付的转换回归。"""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'dependencies'))
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


class ConversionTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.directory = Path(temp.name)
        self.output = self.directory / 'result.docx'

    def convert(self, source, *options):
        result = subprocess.run(
            [sys.executable, '-I', '-B', '-X', 'utf8', ROOT / 'scripts/convert.py',
             '--text' if isinstance(source, str) else '--input', source,
             '--output', self.output, '--offline',
             '--no-embed-fonts', '--no-cover', *options],
            capture_output=True, encoding='utf-8', cwd=self.directory, timeout=30,
        )
        return result.returncode, json.loads(result.stdout)

    def test_table_escape_and_recover_extra_cells(self):
        code, report = self.convert(
            '| A | B |\n| --- | --- |\n| one \\| two | keep |\n| x | y | extra |'
        )
        self.assertEqual(code, 0)
        table = Document(self.output).tables[0]
        self.assertEqual([c.text for c in table.rows[1].cells], ['one | two', 'keep'])
        self.assertIn('extra', table.rows[2].cells[-1].text)
        self.assertEqual(report['diagnostics'][0]['line'], 4)
        self.assertIsNotNone(table.rows[0]._tr.find('.//' + qn('w:tblHeader')))

    def test_strict_does_not_overwrite(self):
        self.output.write_bytes(b'original')
        code, report = self.convert('![missing](none.png)', '--strict', '--overwrite')
        self.assertNotEqual(code, 0)
        self.assertFalse(report['ok'])
        self.assertEqual(report['diagnostics'][0]['line'], 1)
        self.assertEqual(self.output.read_bytes(), b'original')

    def test_fences_breaks_and_literal_text(self):
        source = ('Title\n===\n\n## —— Subtitle\n\n'
                  '**Name**：value  \nNext \\*literal\\*\n\n'
                  '````md\n```py\nprint(1)\n```\n````\n\n'
                  '~~~text\nexample\n~~~\n\nAfter')
        code, report = self.convert(source)
        self.assertEqual(code, 0, report)
        paragraphs = [(p.style.name, p.text) for p in Document(self.output).paragraphs]
        self.assertIn(('Heading 1', 'Title'), paragraphs)
        self.assertIn(('Heading 2', '—— Subtitle'), paragraphs)
        self.assertIn(('Normal', 'Name：value\nNext *literal*'), paragraphs)
        self.assertIn(('Code Block', '```py\nprint(1)\n```'), paragraphs)
        self.assertIn(('Code Block', 'example'), paragraphs)
        self.assertIn(('Normal', 'After'), paragraphs)

    def test_numbering_and_hyperlink(self):
        code, report = self.convert('5. first\n6. [link](https://example.invalid)')
        self.assertEqual(code, 0, report)
        doc = Document(self.output)
        num_id = doc.paragraphs[0]._p.pPr.numPr.numId.val
        numbering = doc.part.numbering_part.element
        instance = next(n for n in numbering.findall(qn('w:num'))
                        if int(n.get(qn('w:numId'))) == num_id)
        self.assertEqual(instance.find('.//' + qn('w:startOverride')).get(qn('w:val')), '5')
        self.assertEqual(doc.paragraphs[1].hyperlinks[0].url, 'https://example.invalid')

    def test_utf16_input_and_image_alt_text(self):
        image = self.directory / 'small image.png'
        image.write_bytes((ROOT / 'verification/sample.png').read_bytes())
        source = self.directory / '中文.md'
        source.write_text('# 中文\n\n![图片说明](<small image.png> "Title")', encoding='utf-16')
        code, report = self.convert(source, '--strict')
        self.assertEqual(code, 0, report)
        shape = Document(self.output).inline_shapes[0]
        self.assertEqual(shape._inline.docPr.get('descr'), '图片说明')
        self.assertLess(shape.width.inches, 6)

    def test_markdown_link_destinations(self):
        code, report = self.convert('[label](https://example.invalid/a_(b) "title")')
        self.assertEqual(code, 0, report)
        paragraph = Document(self.output).paragraphs[0]
        self.assertEqual(paragraph.text, 'label')
        self.assertEqual(paragraph.hyperlinks[0].url, 'https://example.invalid/a_(b)')

    def test_unrendered_syntax_and_invalid_arguments(self):
        code, report = self.convert('# Topic[^1]\n\n> $x^2$\n\n~~~mermaid\nA --> B\n~~~')
        self.assertEqual(code, 0, report)
        self.assertEqual({(d['code'], d['line']) for d in report['diagnostics']},
                         {('unsupported_footnote', 1), ('unsupported_math', 3),
                          ('unsupported_code_rendering', 5)})
        code, report = self.convert('sample', '--invalid')
        self.assertEqual(code, 2)
        self.assertFalse(report['ok'])

    def test_optional_chapter_break_and_native_toc(self):
        code, report = self.convert('# First\n\n# Second', '--no-h1-page-break', '--toc', '--strict')
        self.assertEqual(code, 0, report)
        doc = Document(self.output)
        self.assertEqual([p.text for p in doc.paragraphs if p.style.name == 'Heading 1'],
                         ['First', 'Second'])
        self.assertEqual(doc._element.xml.count('w:type="page"'), 1)
        self.assertIn(' TOC ', doc._element.xml)

    def test_note_properties_are_omitted_and_source_lines_preserved(self):
        source = ('---\ntitle: Metadata title\nsource: https://example.invalid\n'
                  '作者:\n  - "[[@author]]"\ndescription: |\n  Hidden description\n'
                  'tags:\n- internal\n---\n\n# Body\n\n![missing](none.png)')
        code, report = self.convert(source)
        self.assertEqual(code, 0, report)
        self.assertEqual(report['elements'], 2)
        self.assertEqual([(d['code'], d['line']) for d in report['diagnostics']],
                         [('image_missing', 14)])
        paragraphs = [p.text for p in Document(self.output).paragraphs]
        self.assertEqual(paragraphs, ['Body', '[图片：missing](none.png)'])

    def test_opening_divider_and_unclosed_properties_are_preserved(self):
        cases = (('---\n\n# Body\n\n---', 'Body'),
                 ('---\ntitle: Unclosed\n\n# Body', 'title: Unclosed'))
        for source, expected in cases:
            with self.subTest(source=source):
                code, report = self.convert(source, '--overwrite')
                self.assertEqual(code, 0, report)
                self.assertIn(expected, [p.text for p in Document(self.output).paragraphs])
                self.assertIn('w:pBdr', Document(self.output)._element.xml)

    def test_template_modes(self):
        template = self.directory / 'template.docx'
        doc = Document()
        doc.styles['Normal'].font.size = Pt(16)
        doc.add_paragraph('Before')
        doc.add_paragraph('{{正文}}')
        doc.add_paragraph('After')
        doc.sections[0].header.paragraphs[0].text = 'Company'
        doc.sections[0].footer.paragraphs[0].text = 'Footer'
        doc.save(template)
        options = ('--template', template, '--template-mode')
        code, report = self.convert('Inserted', *options, 'placeholders')
        self.assertEqual(code, 0, report)
        doc = Document(self.output)
        self.assertEqual([p.text for p in doc.paragraphs], ['Before', 'Inserted', 'After'])
        self.assertEqual(doc.sections[0].header.paragraphs[0].text, 'Company')
        self.assertEqual(doc.sections[0].footer.paragraphs[0].text, 'Footer')
        self.assertEqual(doc.styles['Normal'].font.size.pt, 16)
        code, report = self.convert('Inserted', *options, 'styles', '--overwrite')
        self.assertEqual(code, 0, report)
        self.assertEqual([p.text for p in Document(self.output).paragraphs], ['Inserted'])


if __name__ == '__main__':
    unittest.main()
