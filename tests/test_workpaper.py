"""Regression tests for evidence semantics, native layout, and edit scope."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile,ZIP_DEFLATED
import copy,hashlib,json,subprocess,sys,unittest
from PIL import Image
from openpyxl import Workbook,load_workbook
from openpyxl.chart import BarChart,Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.datavalidation import DataValidation
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'scripts'))
from make_demo import make_demo
from build_workpaper import build
from workpaper_model import load_manifest,validate_manifest
from patch_workbook import patch
from verify_workbook import read_parts,verify,sha
from extract_outline_links import extract
from extract_report import load_style_levels,walk,extract_section
from split_sentences import build as split_coverage

class WorkpaperTests(unittest.TestCase):
    def setUp(self):
        self.tmp=TemporaryDirectory();self.addCleanup(self.tmp.cleanup);self.folder=Path(self.tmp.name);self.report=make_demo(self.folder/'demo');self.manifest=self.folder/'demo/manifest.json'
    def data(self):return load_manifest(self.manifest)
    def test_editable_layout_links_metadata_and_media_dedup(self):
        w=load_workbook(self.folder/'demo/demo.xlsx');s=w['3.4']
        self.assertEqual(s['A3'].value,'示例公司持续投入品牌建设。');self.assertEqual(s['A3'].data_type,'s');self.assertIn('A3:I16',{str(r)for r in s.merged_cells.ranges})
        self.assertEqual(s['J3'].hyperlink.target,'https://example.com/demo-report.pdf');self.assertIn('采集 2026-11-01',s['J4'].value)
        self.assertEqual(len(s._images),5);self.assertEqual(self.report['media_files'],2)
        self.assertTrue((self.folder/'demo/previews/3.4.png').is_file());self.assertFalse(self.report['native_excel_open'])
        self.assertEqual(self.report['coverage']['correction'],1)
    def test_formula_like_text_is_literal_and_table_values_are_typed(self):
        d=self.data();d.pop('_base');d['sheets'][0]['rows'][0]['text']='=1+1';d['sheets'][0]['tables']=[{'columns':[{'title':'日期'},{'title':'比例','format':'0.0%'},{'title':'金额','format':'0.00'}],'rows':[[{'date':'2026-11-01'},.25,100.25]]}]
        p=self.folder/'demo/literal.json';p.write_text(json.dumps(d),encoding='utf8');out=self.folder/'literal.xlsx';build(p,out);w=load_workbook(out);s=w['3.4']
        self.assertEqual(s['A3'].data_type,'s');self.assertEqual(s['A3'].value,'=1+1');self.assertTrue(any(c.data_type=='d'for row in s for c in row));self.assertTrue(any(c.value==.25 and c.number_format=='0.0%'for row in s for c in row))
    def test_reject_broker_or_personal_sources(self):
        for bad in ['broker_research','personal_post']:
            with self.subTest(bad=bad):
                d=self.data();d['sheets'][0]['rows'][0]['sources'][0]['source_class']=bad
                with self.assertRaisesRegex(ValueError,'not permitted'):validate_manifest(d)
    def test_archive_cannot_be_full_support(self):
        d=self.data();source=d['sheets'][0]['rows'][0]['sources'][0];source.update(source_class='archive',kind='archive_excerpt')
        with self.assertRaisesRegex(ValueError,'support needs'):validate_manifest(d)
    def test_archive_cannot_complete_missing_sentence_coverage(self):
        d=self.data();row=d['sheets'][0]['rows'][0];row['sentence_ids']=['s1','s2'];row['sources'][0]['claim_ids']=['s1']
        archive=copy.deepcopy(row['sources'][0]);archive.update(source_class='archive',kind='archive_excerpt',claim_ids=['s2']);row['sources'].append(archive)
        with self.assertRaisesRegex(ValueError,'support needs'):validate_manifest(d)
    def test_database_deferral_requires_explicit_user_scope(self):
        d=self.data();d['sheets'][0]['rows'][-1]['user_deferred']=False
        with self.assertRaisesRegex(ValueError,'explicit user'):validate_manifest(d)
    def test_physical_and_printed_pages_are_distinct(self):
        d=self.data();src=d['sheets'][0]['rows'][0]['sources'][0];src.update(source_class='company_filing',kind='pdf_crop',pdf_page=43,printed_page=42);d['demo']=False;d['sheets'][0]['rows']=d['sheets'][0]['rows'][:1];validate_manifest(d)
        d.pop('_base');p=self.folder/'demo/pdf.json';p.write_text(json.dumps(d),encoding='utf8');out=self.folder/'pages.xlsx';build(p,out);s=load_workbook(out)['3.4'];self.assertIn('PDF第43页；印刷第42页',s['J2'].value);self.assertTrue(s['J1'].hyperlink.target.endswith('#page=43'))
    def test_missing_screenshot_or_unmapped_sentence_rejected(self):
        d=self.data();d['sheets'][0]['rows'][0]['sources'][0]['image']='missing.png'
        with self.assertRaisesRegex(ValueError,'missing or invalid'):validate_manifest(d)
        d=self.data();d['sheets'][0]['rows'][0]['sentence_ids']=['s1','s2']
        with self.assertRaisesRegex(ValueError,'claim_ids'):validate_manifest(d)
    def baseline(self,chart_on_target=False):
        w=Workbook();s=w.active;s.title='3.4';s['A1']='原始文本';other=w.create_sheet('保留');other['A1']=100;other['A2']='=A1*2';other['A3']='留存合并';other.merge_cells('A3:D4');other['E1']='Link';other['E1'].hyperlink='https://example.com/keep';dv=DataValidation(type='list',formula1='"A,B"');other.add_data_validation(dv);dv.add(other['F1'])
        chart=BarChart();chart.add_data(Reference(other,min_col=1,min_row=1,max_row=2));(s if chart_on_target else other).add_chart(chart,'H1');other.add_image(XLImage(str(self.folder/'demo/product.png')),'A8');path=self.folder/'baseline.xlsx';w.save(path)
        parts=read_parts(path);parts['customXml/marker.xml']=b'<preserve>synthetic sentinel</preserve>'
        with ZipFile(path,'w',ZIP_DEFLATED)as z:
            for key,value in parts.items():z.writestr(key,value)
        return path
    def test_patch_preserves_unrelated_chart_formula_image_and_package(self):
        original=self.baseline();before=read_parts(original);digest=sha(original.read_bytes());out=self.folder/'patched.xlsx';report=patch(original,self.folder/'demo/demo.xlsx','3.4',out,digest);after=read_parts(out)
        self.assertTrue(report['other_worksheets_unchanged']);self.assertEqual(sha(original.read_bytes()),digest);self.assertEqual(before['customXml/marker.xml'],after['customXml/marker.xml'])
        for part in ['xl/worksheets/sheet2.xml','xl/charts/chart1.xml','xl/drawings/drawing1.xml','xl/workbook.xml']:
            if part in before:self.assertEqual(before[part],after[part],part)
        s=load_workbook(out)['保留'];self.assertEqual(s['A2'].value,'=A1*2');self.assertEqual(len(s._charts),1);self.assertEqual(len(s._images),1);self.assertEqual(len(s.data_validations.dataValidation),1)
    def test_patch_rejects_stale_input_and_unsupported_target_chart(self):
        original=self.baseline()
        with self.assertRaisesRegex(ValueError,'changed'):patch(original,self.folder/'demo/demo.xlsx','3.4',self.folder/'bad.xlsx','0'*64)
        original=self.baseline(chart_on_target=True)
        with self.assertRaisesRegex(ValueError,'charts/objects'):patch(original,self.folder/'demo/demo.xlsx','3.4',self.folder/'bad.xlsx')
    def test_patch_preserves_target_protection_and_implicit_styles(self):
        original=self.baseline();w=load_workbook(original);w['3.4'].protection.sheet=True;w['3.4'].oddHeader.center.text='Synthetic header';w.save(original)
        edited=self.folder/'plain-edit.xlsx';e=Workbook();e.active.title='3.4';e.active['A1']='New plain text';e.save(edited)
        out=self.folder/'protected-patch.xlsx';patch(original,edited,'3.4',out)
        result=load_workbook(out)['3.4'];self.assertTrue(result.protection.sheet);self.assertEqual(result.oddHeader.center.text,'Synthetic header');self.assertEqual(result['A1'].value,'New plain text');self.assertGreater(result['A1'].style_id,0)
    def test_docx_links_and_sparse_chart_points_survive_extraction(self):
        p=self.folder/'sample.docx';w='http://schemas.openxmlformats.org/wordprocessingml/2006/main';c='http://schemas.openxmlformats.org/drawingml/2006/chart';r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        doc=f'<w:document xmlns:w="{w}" xmlns:r="{r}"><w:body><w:p><w:pPr><w:pStyle w:val="H2"/></w:pPr><w:r><w:t>为什么转型</w:t></w:r></w:p><w:p><w:r><w:t>公司持续投入。</w:t></w:r><w:hyperlink r:id="source"><w:r><w:t>正式资料</w:t></w:r></w:hyperlink></w:p></w:body></w:document>'
        chart=f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:lineChart><c:ser><c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>2020</c:v></c:pt><c:pt idx="1"><c:v>2021</c:v></c:pt><c:pt idx="2"><c:v>2022</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="2"><c:v>8</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>'
        with ZipFile(p,'w')as z:
            z.writestr('word/document.xml',doc);z.writestr('word/styles.xml',f'<w:styles xmlns:w="{w}"><w:style w:styleId="H2"><w:name w:val="Heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>');z.writestr('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="source" Target="https://example.com/filing.pdf" Type="'+r+'/hyperlink"/></Relationships>');z.writestr('word/charts/chart1.xml',chart)
        data=extract(p,True);self.assertEqual(data['paragraphs'][1]['links'][0]['url'],'https://example.com/filing.pdf');self.assertEqual(data['charts'][0]['series'][0]['values'],[None,None,'8'])
        with ZipFile(p)as z:section=extract_section(walk(z,load_style_levels(z)),'为什么','3.4')
        coverage=split_coverage(section);self.assertTrue(coverage['subsections'][0]['paragraphs'][0]['sentences'][0]['review_required'])

if __name__=='__main__':unittest.main()
