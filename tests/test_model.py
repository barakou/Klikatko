import json
import tempfile
import unittest
from pathlib import Path

from klikatko.model import Store, catalog, flatten, safe_name, selected_nodes


class CatalogTests(unittest.TestCase):
    def test_catalog_ids_are_unique_and_have_expected_size_and_depth(self):
        data = catalog()
        nodes = list(flatten())
        ids = [node['id'] for node in nodes]
        self.assertEqual(len(nodes), 87)
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual([node['id'] for node in data['nodes']], ['A', 'B'])

        def depths(nodes, level=1):
            for node in nodes:
                yield level
                yield from depths(node.get('children', []), level + 1)

        self.assertEqual(max(depths(data['nodes'])), 5)
        self.assertIn('B3.4.SO101.a', ids)
        self.assertIn('B3.4.SO401.b', ids)

    def test_selected_child_keeps_ancestors_and_excludes_siblings(self):
        result = selected_nodes(['B3.4.SO101.b'])
        self.assertEqual(len(result), 1)
        section_b = result[0]
        self.assertEqual(section_b['id'], 'B')
        section_b3 = section_b['children'][0]
        self.assertEqual(section_b3['id'], 'B3')
        section_b34 = section_b3['children'][0]
        self.assertEqual(section_b34['id'], 'B3.4')
        so101 = section_b34['children'][0]
        self.assertEqual(so101['id'], 'B3.4.SO101')
        self.assertEqual([node['id'] for node in so101['children']], ['B3.4.SO101.b'])
        self.assertTrue(so101['children'][0]['explicit'])
        self.assertFalse(section_b['explicit'])

    def test_selected_nodes_rejects_unknown_ids(self):
        with self.assertRaises(ValueError):
            selected_nodes(['B99'])


class StoreTests(unittest.TestCase):
    def test_project_and_preset_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            store = Store(directory)
            project = {
                'name': 'Ukázková stavba',
                'number': '2026-01',
                'location': 'Praha',
                'investor': 'Investor',
                'author': 'Projektant',
                'date': '2026-10-27',
                'stage': 'Povolení',
                'output_dir': '/tmp/vystupy',
                'filename': 'zprava.docx',
                'cover_image': '/tmp/obrazek.png',
                'selected': ['B3.4.SO101.b'],
            }
            saved_project = store.save('projects', project)
            listed_projects = store.list('projects')
            self.assertEqual(len(listed_projects), 1)
            self.assertEqual(listed_projects[0], saved_project)
            self.assertEqual(listed_projects[0]['name'], project['name'])
            self.assertEqual(listed_projects[0]['selected'], project['selected'])

            preset = {
                'name': 'Dopravní část',
                'selected': ['B3.4.SO101.b'],
                'investor': 'Soukromý investor',
                'author': 'Osobní údaj',
                'location': 'Soukromé místo',
                'filename': 'osobni.docx',
            }
            saved_preset = store.save('presets', preset)
            listed_presets = store.list('presets')
            self.assertEqual(len(listed_presets), 1)
            self.assertEqual(listed_presets[0], saved_preset)
            self.assertEqual(listed_presets[0]['name'], preset['name'])
            self.assertEqual(listed_presets[0]['selected'], preset['selected'])
            self.assertEqual(set(listed_presets[0]), {'name', 'selected', 'id', 'schema_version'})

    def test_invalid_selected_ids_are_rejected_when_saving_project_and_preset(self):
        with tempfile.TemporaryDirectory() as directory:
            store = Store(directory)
            project = {'name': 'Projekt', 'selected': ['unknown']}
            preset = {'name': 'Předvolba', 'selected': ['unknown']}
            with self.assertRaises(ValueError):
                store.save('projects', project)
            with self.assertRaises(ValueError):
                store.save('presets', preset)


class SafeNameTests(unittest.TestCase):
    def test_windows_reserved_names_are_prefixed(self):
        self.assertEqual(safe_name('CON'), '_CON')
        self.assertEqual(safe_name('nul.txt'), '_nul.txt')
        self.assertEqual(safe_name('LPT1.log'), '_LPT1.log')

    def test_path_and_windows_filename_characters_are_sanitized(self):
        result = safe_name('C:\\Users/A:B*?"<name>|.docx')
        self.assertEqual(result, 'C--Users-A-B----name--.docx')
        self.assertFalse(any(char in result for char in '<>:"/\\|?*'))
        self.assertFalse(any(ord(char) < 32 for char in result))
        self.assertTrue(safe_name('...').strip(' .'))


if __name__ == '__main__':
    unittest.main()
