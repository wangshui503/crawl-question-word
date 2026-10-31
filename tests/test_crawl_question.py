import base64
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

spec = importlib.util.spec_from_file_location('crawler', Path(__file__).resolve().parents[1] / 'crawl_question.py')
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


class Tests(unittest.TestCase):
    def test_expired_token(self):
        payload = base64.urlsafe_b64encode(json.dumps({'exp': 1}).encode()).decode().rstrip('=')
        with self.assertRaises(m.AuthError):
            m.normalize_token('Bearer header.' + payload + '.signature')

    def test_reject_webpage_url(self):
        with self.assertRaises(m.CrawlError):
            m.Client('token', 'https://app.ji-guang.top/#/home')

    def test_401_no_retry(self):
        client = m.Client('token')
        response = Mock(status_code=401)
        client.session.get = Mock(return_value=response)
        with self.assertRaises(m.AuthError):
            client.get('/user/subjects/test')
        self.assertEqual(client.session.get.call_count, 1)

    def test_non_json(self):
        client = m.Client('token')
        response = Mock(status_code=200)
        response.json.side_effect = ValueError('HTML')
        client.session.get = Mock(return_value=response)
        with self.assertRaisesRegex(m.CrawlError, 'JSON'):
            client.get('/test')

    def test_pagination_and_repeat_guard(self):
        client = m.Client('token')
        client.get = Mock(side_effect=[
            {'items': [{'id': 'a'}], 'has_more': True},
            {'items': [{'id': 'b'}], 'has_more': False},
        ])
        self.assertEqual([x['id'] for x in client.pages('/s', 'items', page_size=100)], ['a', 'b'])
        self.assertEqual(client.get.call_args_list[1].args[1]['skip'], 1)
        client.get = Mock(return_value={'questions': [{'question_id': 'q'}], 'has_more': True})
        with self.assertRaisesRegex(m.CrawlError, '分页重复'):
            list(client.pages('/c', 'questions'))

    def test_full_export_and_explanation_fallback(self):
        from docx import Document
        client = m.Client('token')

        def get(path, params=None):
            return {
                '/user/subjects/': {'items': [{'id': 's', 'name': '卫法'}], 'has_more': False},
                '/user/subjects/s': {'chapters': [{'id': 'c', 'name': '法律责任'}]},
                '/user/chapters/c': {'questions': [{'question_id': 'q'}], 'has_more': False},
                '/user/questions/q': {'content': '题目内容', 'options': {'A': '选项一'}, 'correct_answer': 'A', 'explanation': '标准解析'},
            }[path]

        client.get = get
        with tempfile.TemporaryDirectory() as directory:
            m.crawl(client, '三年级', directory, delay=0)
            files = list(Path(directory).glob('*.docx'))
            self.assertEqual(len(files), 1)
            self.assertEqual(files[0].name, '卫法题库.docx')
            text = '\n'.join(p.text for p in Document(files[0]).paragraphs)
            for content in ('第1题', '题目内容', 'A．选项一', '答案：A', '解析：标准解析'):
                self.assertIn(content, text)
            self.assertEqual(list(Path(directory).glob('*.tmp.docx')), [])

    def test_empty_failed_chapter_does_not_export(self):
        client = m.Client('token')
        def get(path, params=None):
            if path == '/user/subjects/':
                return {'items': [{'id': 's', 'name': '卫法'}], 'has_more': False}
            if path == '/user/subjects/s':
                return {'chapters': [{'id': 'c', 'name': '卫法'}]}
            raise m.CrawlError('接口失败')
        client.get = get
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(m.CrawlError):
                m.crawl(client, '三年级', directory, delay=0)
            self.assertEqual(list(Path(directory).glob('*.docx')), [])

    def test_html_cleanup_and_multiline(self):
        self.assertEqual(m.readable_text('<p>第一行&nbsp;<br>第二行</p><script>隐藏</script>'), '第一行\n第二行')
        self.assertEqual(m.readable_text('1 < 2 &amp; 3 > 2\x00'), '1 < 2 & 3 > 2')

    def test_all_options_and_nested_values(self):
        row = m.question_row({'content': '<p>题干</p>', 'options': {'F': ['甲', '乙'], 'A': '<b>选项</b>'}, 'correct_answer': ['A', 'F'], 'ai_explanation': '<p></p>', 'explanation': '备用解析'}, '学科', '章节', 'q')
        self.assertEqual(row['选项'], {'A': '选项', 'F': '甲\n乙'})
        self.assertEqual(row['答案'], 'A\nF')
        self.assertEqual(row['解析'], '备用解析')

    def test_explanation_switch_and_missing_answer(self):
        from docx import Document
        row = m.question_row({'content': '题目', 'options': {}, 'explanation': '不应导出的解析'}, '学科', '章节', 'q')
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'test.docx'
            m.save_word([row], path, '学科', '章节', include_explanation=False)
            text = '\n'.join(p.text for p in Document(path).paragraphs)
            self.assertIn('答案：未提供', text)
            self.assertNotIn('不应导出的解析', text)

    def test_partial_result_notice(self):
        from docx import Document
        row = m.question_row({'content': '题目'}, '学科', '章节', 'q')
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'partial.docx'
            m.save_word([row], path, '学科', '章节', failures=2)
            text = '\n'.join(p.text for p in Document(path).paragraphs)
            self.assertIn('部分结果', text)
            self.assertIn('2 项失败', text)

    def merge_client(self, fail_question=False):
        client = m.Client('token')
        calls = []

        def get(path, params=None):
            calls.append((path, params))
            if path == '/user/subjects/':
                return {'items': [{'id': 's1', 'name': ' 卫法 '}, {'id': 's2', 'name': '卫法基础'}, {'id': 's3', 'name': '其他学科'}], 'has_more': False}
            if path == '/user/subjects/s1':
                return {'chapters': [{'id': 'c1', 'name': '总则'}, {'id': 'c2', 'name': '法律责任'}]}
            if path == '/user/chapters/c1':
                if params['skip'] == 0:
                    return {'questions': [{'id': 'q1'}], 'has_more': True}
                return {'questions': [{'id': 'shared'}], 'has_more': False}
            if path == '/user/chapters/c2':
                return {'questions': [{'id': 'shared'}, {'id': 'q3'}], 'has_more': False}
            if path.startswith('/user/questions/'):
                qid = path.rsplit('/', 1)[-1]
                if fail_question and qid == 'q3':
                    raise m.CrawlError('题目详情失败')
                return {'content': '卫法题目' + qid, 'options': {'A': '选项A', 'B': '选项B'}, 'correct_answer': 'A', 'explanation': '说明' + qid}
            raise AssertionError('不应请求非目标学科：' + path)

        client.get = get
        return client, calls

    def test_only_exact_subject_all_chapters_and_pages_merged_and_deduplicated(self):
        from docx import Document
        client, calls = self.merge_client()
        with tempfile.TemporaryDirectory() as directory:
            m.crawl(client, '', directory, page_size=1, delay=0)
            files = list(Path(directory).glob('*.docx'))
            self.assertEqual([f.name for f in files], ['卫法题库.docx'])
            doc = Document(files[0])
            self.assertEqual([p.text for p in doc.paragraphs if p.style.name == 'Heading 1'], ['第1题', '第2题', '第3题'])
            text = '\n'.join(p.text for p in doc.paragraphs)
            self.assertIn('卫法题目q1', text)
            self.assertIn('卫法题目q3', text)
            self.assertEqual(text.count('卫法题目shared'), 1)
            self.assertIn('来源：卫法 / 总则', text)
            self.assertIn('来源：卫法 / 法律责任', text)
        self.assertNotIn('tag', calls[0][1])
        self.assertFalse(any(path in ('/user/subjects/s2', '/user/subjects/s3') for path, _ in calls))
        self.assertEqual(sum(path == '/user/questions/shared' for path, _ in calls), 1)
        self.assertTrue(any(path == '/user/chapters/c1' and params['skip'] == 1 for path, params in calls))

    def test_partial_fetch_still_writes_one_merged_document(self):
        from docx import Document
        client, _ = self.merge_client(fail_question=True)
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(m.CrawlError, '部分结果'):
                m.crawl(client, '', directory, page_size=1, delay=0)
            files = list(Path(directory).glob('*.docx'))
            self.assertEqual([p.name for p in files], ['卫法题库-部分结果.docx'])
            text = '\n'.join(p.text for p in Document(files[0]).paragraphs)
            self.assertIn('2 道题', text)
            self.assertIn('部分结果', text)

    def test_no_matching_subject_does_not_fetch_details_or_write(self):
        client = m.Client('token')
        client.get = Mock(side_effect=[
            {'items': [{'id': 's', 'name': '卫法基础'}, {'id': 'other', 'name': '其他'}], 'has_more': False},
        ])
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(m.CrawlError, '没有找到'):
                m.crawl(client, '', directory, delay=0)
            self.assertEqual(list(Path(directory).iterdir()), [])
        self.assertEqual(client.get.call_count, 1)

    def test_empty_matching_chapter_does_not_write(self):
        client = m.Client('token')
        client.get = Mock(side_effect=[
            {'items': [{'id': 's', 'name': '卫法'}], 'has_more': False},
            {'chapters': [{'id': 'c', 'name': '卫法'}]},
            {'questions': [], 'has_more': False},
        ])
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(m.CrawlError, '没有成功获取'):
                m.crawl(client, '', directory, delay=0)
            self.assertEqual(list(Path(directory).iterdir()), [])

    def test_chinese_and_latin_fonts(self):
        from docx import Document
        from docx.oxml.ns import qn
        row = m.question_row({'content': '中文 ABC'}, '学科', '卫法', 'q')
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'fonts.docx'
            m.save_word([row], path, '', '卫法')
            doc = Document(path)
            for name in ('Normal', 'Title', 'Heading 1', 'Footer'):
                fonts = doc.styles[name].element.rPr.rFonts
                self.assertEqual(fonts.get(qn('w:eastAsia')), 'SimSun')
                self.assertEqual(fonts.get(qn('w:ascii')), 'Times New Roman')
                self.assertEqual(fonts.get(qn('w:hAnsi')), 'Times New Roman')

    def test_cli_defaults_to_weifa_and_no_tag_filter(self):
        with patch.dict(m.os.environ, {'JIGUANG_TOKEN': 'token'}), patch.object(m, 'Client') as client, patch.object(m, 'crawl') as crawl:
            self.assertEqual(m.main([]), 0)
            self.assertEqual(crawl.call_args.args[1], '')
            self.assertEqual(crawl.call_args.kwargs['subject_name'], '卫法')
            client.return_value.session.close.assert_called_once()

    def test_custom_subject_and_tag_forwarded(self):
        with patch.dict(m.os.environ, {'JIGUANG_TOKEN': 'token'}), patch.object(m, 'Client'), patch.object(m, 'crawl') as crawl:
            self.assertEqual(m.main(['--subject', '其他学科', '--tag', '三年级', '--no-explanation']), 0)
            self.assertEqual(crawl.call_args.args[1], '三年级')
            self.assertEqual(crawl.call_args.kwargs['subject_name'], '其他学科')
            self.assertFalse(crawl.call_args.kwargs['include_explanation'])


if __name__ == '__main__':
    unittest.main()
