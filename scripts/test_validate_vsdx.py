"""Static-check regressions using synthetic OPC fixtures, not application-valid Visio files."""
import tempfile
import unittest
import zipfile
from pathlib import Path

from validate_vsdx import V, R, REL, inspect


def parts():
    v, r, rel = V[1:-1], R[1:-1], REL[1:-1]
    return {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        '_rels/.rels': f'<Relationships xmlns="{rel}"><Relationship Id="d" Target="visio/document.xml"/></Relationships>',
        'visio/document.xml': f'<VisioDocument xmlns="{v}"/>',
        'visio/pages/pages.xml': f'<Pages xmlns="{v}" xmlns:r="{r}"><Page ID="0"><PageSheet>'
            '<Cell N="PageWidth" V="3"/><Cell N="PageHeight" V="2"/></PageSheet><Rel r:id="p"/></Page></Pages>',
        'visio/pages/_rels/pages.xml.rels': f'<Relationships xmlns="{rel}"><Relationship Id="p" Target="page1.xml"/></Relationships>',
        'visio/pages/page1.xml': f'<PageContents xmlns="{v}"><Shapes><Shape ID="1" Type="Shape">'
            '<Section N="Character"><Row IX="0"><Cell N="Size" V="0.0833333333333"/>'
            '<Cell N="Font" V="Times New Roman"/></Row></Section><Text>Model</Text></Shape></Shapes></PageContents>',
    }


PROFILE = {'max_width_mm': 80, 'max_height_mm': 60, 'font_pt': 6,
           'western_font': 'Times New Roman', 'cjk_font': 'SimSun'}


class ValidationTests(unittest.TestCase):
    def check(self, entries=None, **kwargs):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / 'diagram.vsdx'
            with zipfile.ZipFile(path, 'w') as z:
                for name, value in (parts() if entries is None else entries).items():
                    z.writestr(name, value)
            return inspect(path, profile=PROFILE, **kwargs)

    def test_static_success_is_not_application_verification(self):
        result = self.check()
        self.assertTrue(result['static_ok'])
        self.assertEqual(result['shapes'], 1)
        self.assertTrue(result['unverified'])

    def test_missing_relationship_target(self):
        p = parts(); del p['visio/pages/page1.xml']
        self.assertFalse(self.check(p)['static_ok'])

    def test_unresolved_page_relationship(self):
        p = parts(); p['visio/pages/pages.xml'] = p['visio/pages/pages.xml'].replace('r:id="p"', 'r:id="absent"')
        self.assertFalse(self.check(p)['static_ok'])

    def test_root_escape(self):
        p = parts(); p['_rels/.rels'] = p['_rels/.rels'].replace('visio/document.xml', '../outside.xml')
        self.assertFalse(self.check(p)['static_ok'])

    def test_malformed_xml(self):
        p = parts(); p['visio/pages/page1.xml'] = '<broken>'
        self.assertFalse(self.check(p)['static_ok'])

    def test_duplicate_shape_id(self):
        p = parts(); p['visio/pages/page1.xml'] = p['visio/pages/page1.xml'].replace('</Shapes>', '<Shape ID="1"/></Shapes>')
        self.assertFalse(self.check(p)['static_ok'])

    def test_unconnected_target(self):
        p = parts(); p['visio/pages/page1.xml'] = p['visio/pages/page1.xml'].replace('</PageContents>',
            '<Connects><Connect FromSheet="1" ToSheet="77"/></Connects></PageContents>')
        self.assertFalse(self.check(p)['static_ok'])

    def test_media_requires_explicit_override(self):
        p = parts(); p['visio/media/image1.png'] = b'not-a-real-image'
        self.assertFalse(self.check(p)['static_ok'])
        self.assertTrue(self.check(p, allow_media=True)['static_ok'])

    def test_profile_limits(self):
        p = parts(); p['visio/pages/pages.xml'] = p['visio/pages/pages.xml'].replace('V="3"', 'V="4"')
        self.assertFalse(self.check(p)['static_ok'])

    def test_font_size(self):
        p = parts(); p['visio/pages/page1.xml'] = p['visio/pages/page1.xml'].replace('0.0833333333333', '0.2')
        self.assertFalse(self.check(p)['static_ok'])

    def test_fixed_geometry_warns(self):
        p = parts(); p['visio/pages/page1.xml'] = p['visio/pages/page1.xml'].replace('</Shape>',
            '<Section N="Geometry"><Row T="LineTo"><Cell N="X" V="1"/></Row></Section></Shape>')
        result = self.check(p)
        self.assertTrue(result['static_ok'])
        self.assertTrue(result['warnings'])

    def test_inherited_formatting_is_not_claimed_verified(self):
        p = parts(); p['visio/pages/page1.xml'] = p['visio/pages/page1.xml'].replace('N="Character"', 'N="Other"')
        self.assertTrue(self.check(p)['warnings'])

    def test_non_zip(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / 'bad.vsdx'
            path.write_bytes(b'not a zip')
            self.assertFalse(inspect(path)['static_ok'])


if __name__ == '__main__':
    unittest.main()
