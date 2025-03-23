import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from lxml.etree import Element
from openpyxl.workbook.workbook import Workbook

main_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def get_rel_max_id(el: Element) -> int:
    '''ある xml 要素の下で、Relationship 要素の最大 Id を取得する。
    '''
    ids = [
        int(re.search(r'\d+', rel.get("Id", ""))[0])  # 数値部分を抽出
        for rel in el.xpath(".//Relationship")
        if re.search(r'\d+', rel.get("Id", ""))
    ]
    max_id = max(ids) if ids else 0
    return max_id


def restore_folder(
        before_dir: Path, after_dir: Path, folder2restore: str,
        delete_first=False):
    '''folder2restore 引数で指定されたフォルダを復元する。
    '''
    src = before_dir / folder2restore
    dest = after_dir / folder2restore

    if not src.exists():
        return

    if delete_first:
        shutil.rmtree(dest)

    if not os.path.exists(dest):
        shutil.copytree(src, dest)


def restore_xl_drawings(before_dir: Path, after_dir: Path):
    '''xl/drawings/ フォルダ内の _rels/ フォルダと *.xml ファイルを復元する。
    '''
    src = before_dir / 'xl/drawings/'
    dest = after_dir / 'xl/drawings/'

    if not src.exists():
        return

    for f in src.iterdir():
        if f.name == '_rels':
            # _rels/ フォルダ内には drawing*.xml のリレーションしかないので、
            # openpyxl が *.vml を書き換えていたとしても、丸ごと上書きで OK。
            shutil.copytree(f, dest / '_rels')
        elif f.suffix == '.xml':
            shutil.copy2(f, dest)


def add_ns(root: Element, key: str, ns: str) -> Element:
    '''root に xmlns:key="ns" を追加して、新しい root を作る。
    '''
    nsmap = root.nsmap.copy() if root.nsmap else {}
    nsmap[key] = ns

    # 新しい root を作成し直す（既存の要素を移植）
    new_root = etree.Element(root.tag, nsmap=nsmap)
    new_root.extend(root)
    return new_root


def restor_xl_worksheets(before_dir: Path, after_dir: Path):
    '''xl/worksheets/ フォルダ内の _rels/ フォルダと *.xml ファイルを復元する。
    '''
    # Target="../drawings/drawing*.xml" を探すための正規表現パターン。
    target_ptn = re.compile(r'\.\./drawings/drawing[0-9]+.xml')

    src = before_dir / 'xl/worksheets/'
    dest = after_dir / 'xl/worksheets/'

    src_rels = src / '_rels/'
    dest_rels = dest / '_rels/'

    # 保存前の _rels/ フォルダ内の *.xml.rels ファイルについて処理する。
    for f in src_rels.iterdir():
        if not f.name.endswith('.xml.rels'):
            continue

        # 保存前の xml の root を取得する。
        before_tree = etree.parse(f)
        before_root = before_tree.getroot()

        # Target="../drawings/drawing*.xml" の Relationship を取得する。
        namespaces = {'ns': before_root.nsmap[None]}
        rels = before_root.xpath('ns:Relationship', namespaces=namespaces)
        existings = []
        for rel in rels:
            target = rel.get('Target')
            if target and target_ptn.fullmatch(target):
                existings.append(rel)

        if not existings:
            continue

        # 保存後の xml の root を取得または作成する。
        after_path = dest_rels / f.name
        if after_path.exists():
            after_tree = etree.parse(after_path)
            after_root = after_tree.getroot()
        else:
            namespace = \
                "http://schemas.openxmlformats.org/package/2006/relationships"
            after_root = etree.Element("Relationships", xmlns=namespace)

        # sheet*.xml.rels に対応する sheet*.xml ファイル。
        xml_path = dest / f.name[:-5]
        xml_tree = etree.parse(xml_path)
        xml_root = xml_tree.getroot()
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/"\
            "relationships"
        xml_root = add_ns(xml_root, 'r', rel_ns)

        # sheet*.xml ファイルに drawing を足す場合の場所を求めておく。
        # legacyDrawing 要素よりも前に挿入する必要があるらしい。
        legacy = xml_root.find(f'.//{{{main_ns}}}legacyDrawing')
        drw_index = xml_root.index(legacy) if legacy is not None else -1

        # 保存後の xml の root に復元した Relationship を足していく。
        max_id = get_rel_max_id(after_root)
        for rel in existings:
            max_id += 1
            rel.set('Id', f'rId{max_id}')
            after_root.append(rel)

            # 対応する drawing 要素を sheet*.xml に足す。
            drw = etree.Element('drawing')
            drw.set(f'{{{rel_ns}}}id', f'rId{max_id}')
            if drw_index >= 0:
                xml_root.insert(drw_index, drw)
                drw_index += 1
            else:
                xml_root.append(drw)

        # sheet*.xml を保存する。
        xml_tree = etree.ElementTree(xml_root)
        xml_tree.write(xml_path, encoding='utf-8')

        # 保存する。
        after_tree = etree.ElementTree(after_root)
        after_tree.write(after_path, encoding='utf-8')


diagram_ctype_base = \
    'application/vnd.openxmlformats-officedocument.drawingml.{}+xml'
diagram_ctype_map = {
    'colors': diagram_ctype_base.format('diagramColors'),
    'data': diagram_ctype_base.format('diagramData'),
    'layout': diagram_ctype_base.format('diagramLayout'),
    'quickStyle': diagram_ctype_base.format('diagramStyle'),
    'drawing': "application/vnd.ms-office.drawingml.diagramDrawing+xml",
}
drawing_ctype = 'application/vnd.openxmlformats-officedocument.drawing+xml'


def adjust_content_types(after_dir: Path):
    '''[Content_Types].xml の内容を調整する。
    '''
    file_path = after_dir / '[Content_Types].xml'
    tree = etree.parse(file_path)
    root = tree.getroot()

    # 画像の拡張子一覧
    img_exts = {"png", "jpg", "jpeg", "gif", "bmp", "tiff", "tif"}

    # 既存の Default 要素の拡張子を取得する。
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    namespaces = {'ns': ct_ns}
    defaults = root.xpath('ns:Default', namespaces=namespaces)
    def_exts = {elem.get("Extension") for elem in defaults}

    exts = set()
    for folder in ["xl/diagrams/", "xl/media/", "xl/drawings/"]:
        dir_path = after_dir / folder
        if not dir_path.exists():
            continue
        for file in dir_path.iterdir():
            ext = file.suffix.strip('.')
            if ext:
                exts.add(ext)

    # 既存にない拡張子を Default 要素として追加する。
    for ext in exts - def_exts:
        if ext in img_exts:
            content_type = f"image/{ext}"
        elif ext == "emf":
            content_type = "image/x-emf"
        else:
            content_type = f"application/{ext}"

        elem = etree.Element(
            "Default", Extension=ext, ContentType=content_type)
        root.append(elem)

    # xl/diagrams/ フォルダ内のファイルに対する Overrice 要素を追加
    dir_path = after_dir / 'xl/diagrams'
    for file in dir_path.iterdir():
        if file.suffix == ".xml":
            part_name = f"/xl/diagrams/{file.name}"
            ctype = diagram_ctype_map[re.sub(r'\d+$', '', file.stem)]
            override_elem = etree.Element(
                "Override", PartName=part_name, ContentType=ctype)
            root.append(override_elem)

    # xl/drawings/ フォルダ内のファイルに対する Overrice 要素を追加
    dir_path = after_dir / 'xl/drawings'
    for file in dir_path.iterdir():
        if file.suffix == ".xml":
            part_name = f"/xl/drawings/{file.name}"
            override_elem = etree.Element(
                "Override", PartName=part_name, ContentType=drawing_ctype)
            root.append(override_elem)

    # 保存する。
    tree = etree.ElementTree(root)
    tree.write(file_path, encoding='utf-8')


def restore_ext_lst(before_dir: Path, after_dir: Path):
    '''xl/worksheets/*.xml ファイル内の <extLst> を復元する。
    '''
    src = before_dir / 'xl/worksheets/'
    dest = after_dir / 'xl/worksheets/'

    # 保存前の worksheets/ フォルダ内の *.xml ファイルについて処理する。
    for f in src.iterdir():
        if not f.name.endswith('.xml'):
            continue

        # 保存前の xml の root を取得する。
        before_tree = etree.parse(f)
        before_root = before_tree.getroot()

        # 保存前の xml から <extLst> を取得する。
        extLsts = before_root.findall(f".//{{{main_ns}}}extLst")
        if not extLsts:
            continue

        # 保存後の xml の root を取得する。
        after_tree = etree.parse(dest / f.name)
        after_root = after_tree.getroot()

        # すべての <extLst> を復元
        for extLst in extLsts:
            parent = extLst.getparent()
            parent_tag = parent.tag.split('}')[-1]

            # 保存前と同じ親を探し、その下に復元する。
            after_parent = after_root.find(f".//{{{main_ns}}}{parent_tag}")
            if after_parent:
                after_parent.append(extLst)
            else:
                # 同じ親がなければ root 直下に復元する。
                after_root.append(extLst)

        # 保存する。
        after_tree = etree.ElementTree(after_root)
        after_tree.write(dest / f.name, encoding='utf-8')


def save_with_drawings(
        wb: Workbook, src: Path, dest: Path, temp_dir_args=None):
    '''図形や画像を復元しつつ Workbook を保存する。

    Parameters
    ----------
    wb : Workbook
        保存する Workbook。
    src : Path
        復元する図形や画像の元となるブックファイルのパス。
    dest : Path
        Workbook の保存先となるブックファイルのパス。
        上書き保存する場合は src と同じにする。
    temp_dir_args : dict | None, default None
        TemporaryDirectory を作る時のパラメータ。
    '''
    if temp_dir_args is None:
        temp_dir_args = {}

    with tempfile.TemporaryDirectory(**temp_dir_args) as temp_dir:
        before_dir = Path(temp_dir) / 'before'
        after_dir = Path(temp_dir) / 'after'

        # src を before_dir に解凍する。
        with zipfile.ZipFile(src, 'r') as zf:
            zf.extractall(str(before_dir))

        # wb を dest に保存する。
        wb.save(dest)

        # dest を after_dir に解凍する。
        with zipfile.ZipFile(dest, 'r') as zf:
            zf.extractall(str(after_dir))

        # xl/diagrams/, xl/media/ フォルダを復元する。
        restore_folder(before_dir, after_dir, 'xl/diagrams/')
        restore_folder(before_dir, after_dir, 'xl/media/')

        # xl/drawings/ フォルダ内のコンテンツを復元する。
        restore_xl_drawings(before_dir, after_dir)

        # xl/worksheets/ フォルダ内のコンテンツを復元する。
        restor_xl_worksheets(before_dir, after_dir)

        # [Content_Types].xml の内容を調整する。
        adjust_content_types(after_dir)

        # 数式を使った「データの入力規則」を復元する。
        restore_ext_lst(before_dir, after_dir)

        # dest に圧縮しなおす。
        with zipfile.ZipFile(dest, 'w') as zf:
            for root, _, files in os.walk(after_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, after_dir)
                    zf.write(file_path, arcname)
