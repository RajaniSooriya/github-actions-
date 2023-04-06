import json
import unittest
from unittest.mock import patch, MagicMock
from app import app, generateScripts, generateSingleScript, extractAndGenerate


class TestApp(unittest.TestCase):

    def setUp(self):
        self.app = app.test_client()

    def test_root(self):
        response = self.app.get('/')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_data(), b'Hello from ScriptGenAI')

    def test_generate_scripts(self):
        slides = ['S1: This is the first slide.', 'S2: This is the second slide.']
        generated_scripts = generateScripts(slides)
        self.assertTrue(isinstance(generated_scripts, list))
        self.assertEqual(len(generated_scripts), 2)
        self.assertTrue(isinstance(generated_scripts[0], str))

    def test_generate_single_script(self):
        slides = 'S1: This is the first slide.\nS2: This is the second slide.'
        generated_script = generateSingleScript(slides)
        self.assertTrue(isinstance(generated_script, str))
        self.assertGreater(len(generated_script), 0)

    @patch('app.ppextractmodule.process')
    def test_extract_and_generate(self, mock_process):
        mock_process.return_value = 'S1: This is the first slide.\nS2: This is the second slide.'
        generated_script = extractAndGenerate()
        self.assertTrue(isinstance(generated_script, str))
        self.assertGreater(len(generated_script), 0)

    @patch('app.request')
    @patch('app.extractAndGenerate')
    def test_scripts(self, mock_extractAndGenerate, mock_request):
        mock_request.urlretrieve.return_value = None
        mock_extractAndGenerate.return_value = 'This is the generated script.'
        headers = {'link': 'http://example.com/presentation.pptx'}
        response = self.app.get('/scripts', headers=headers)
        self.assertEqual(response.status_code, 200)
        data = json.loads(response.get_data())
        self.assertTrue(isinstance(data, dict))
        self.assertTrue('script' in data)
        self.assertEqual(data['script'], 'This is the generated script.')
