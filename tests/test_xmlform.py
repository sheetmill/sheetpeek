"""The generators' XML edits hold whichever writer openpyxl used (lxml when it is installed)."""

import zipfile

from capture.screen.run import patch_part
from capture.xmlform import serialized, standard_form


def test_empty_elements_close_as_the_standard_library_closes_them():
    lxml = '<fonts><font><sz val="11"/><b/></font></fonts><c r="A1"><f>1</f><v></v></c>'
    assert standard_form(lxml) == (
        '<fonts><font><sz val="11" /><b /></font></fonts><c r="A1"><f>1</f><v /></c>'
    )
    assert standard_form('<sz val="11" />') == '<sz val="11" />'


def test_an_edit_sees_its_part_in_the_standard_form(tmp_path):
    path = tmp_path / "book.xlsx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/styles.xml", '<fills><fill><fgColor theme="5"/></fill></fills>')
    seen = []
    patch_part(path, {"xl/styles.xml": lambda text: seen.append(text) or text})
    assert seen == ['<fills><fill><fgColor theme="5" /></fill></fills>']


def test_the_theme_openpyxl_writes_as_a_fixed_text_stays_as_it_is():
    assert serialized("xl/styles.xml")
    assert serialized("xl/worksheets/_rels/sheet1.xml.rels")
    assert not serialized("xl/theme/theme1.xml")
    assert not serialized("xl/media/image1.png")
