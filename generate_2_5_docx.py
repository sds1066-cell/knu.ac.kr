#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the verified Korean translation of thesis section 2.5 as DOCX and ZIP.

The DOCX is assembled directly as Office Open XML. Footnotes 61–71 are true
Word footnotes; body typography follows the supplied thesis-writing example:
Batang 10.5 pt, justified paragraphs, first-line indentation, A4 page.
"""

from __future__ import annotations

import hashlib
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

OUT_DIR = Path("generated")
DOCX_PATH = OUT_DIR / "section_2_5_korean_translation.docx"
VERIFY_PATH = OUT_DIR / "section_2_5_verification.txt"
ZIP_PATH = OUT_DIR / "section_2_5_korean_translation.zip"

TITLE = "2.5 협정 형성의 가능성 조건"

PARAGRAPHS = [
    "한중 양국이 최종 경계에 관한 돌파구를 마련하지 못한 상황에서 어업협정을 체결하였다는 사실은 양국이 모든 분쟁이 해결된 뒤에야 협력하려고 한 것이 아님을 보여 준다. 협상 과정에서 분리된 것은 두 유형의 문제였다. 하나는 공간, 자원 및 관할권의 최종적 배분에 관한 문제이고, 다른 하나는 장기간 유보할 수 없는 어업 관리의 문제이다. 협정의 성립은 이러한 분리가 법적으로 가능하고 현실적으로 필요하며 정치적으로 수용될 수 있다는 데 의존하였다. 인접 해역에 이미 존재하던 어업 약정은 이러한 선택을 보다 쉽게 식별할 수 있게 하였지만, 양국 자체의 협상을 대체하지는 않았다.",
    "UNCLOS 제74조 제3항과 제83조 제3항은 관련국이 최종 경계획정을 해하지 않는 범위에서 실질적인 잠정약정을 체결할 수 있도록 하고, 최종 합의를 위태롭게 하지 않도록 요구한다.[[61]] 한중 협정 제14조는 해양법 문제에 관한 양국의 입장을 유보한다.[[62]] 협정의 나머지 조항은 규율 대상을 어업 관계에 집중시키고, 입어, 자원 보존, 조업 질서, 잠정조치수역 및 공동위원회를 중심으로 규칙을 설정한다.[[63]] 중국 측도 협정 서명 후 공개한 설명에서 이 협정을 배타적경제수역 경계획정 이전의 잠정적인 어업 약정으로 이해하였다.[[64]] 이러한 자료는 제한적 협력을 위해 어느 일방도 최종적인 권리 주장을 먼저 포기할 필요가 없었음을 보여 준다. 국제법이 협정을 직접 만들어 낸 것은 아니지만, ‘먼저 경계를 획정해야만 협력할 수 있다’는 제도적 장애를 제거하였다.",
    "한국의 공식 통계는 2000년 전후의 어가인구, 어가 및 등록어선 현황을 기록하고 있으며, 중국 국가통계국 역시 1995년부터 1999년까지의 해양 어획생산량과 황해 구역별 자료를 연속적으로 발표하였다.[[65]][[66]] 두 통계는 단위와 범위가 서로 다르므로 양측의 압력을 비교하는 데 사용할 수 없고, 자원이 이미 고갈되었다는 점을 입증할 수도 없다. 이들 자료가 뒷받침할 수 있는 판단은 어업활동이 협정 형성 전후에 줄곧 현실적인 관리대상이었다는 것이다. 한국 외교통상부 장관은 비준 절차에서 해양경계가 획정되지 않은 상태에서 일방적으로 관할권을 행사할 경우 충돌이 발생할 수 있으며, 한국 정부는 협정을 통하여 해양생물자원을 보존하고 어업질서를 확립하고자 한다고 지적하였다.[[67]] 이 발언은 한국 측의 입장만을 나타내지만, 공동 문서에 규정된 자원 보존 및 정상적인 조업질서의 목표와 상응한다. 최종 경계획정 협상은 계속될 수 있었지만 일상적인 조업, 법집행 및 자원관리는 장기간 기다릴 수 없었으며, 이는 제한적 약정의 현실적 필요성을 구성하였다.",
    "박재영과 최종화는 2000년의 연구에서 1997년 중일 어업 약정 이후의 방안 변화, 한중 어업협상 및 해양경계획정 회담의 분리 진행을 서로 연계하였다.[[68]] 이는 현실적인 어업 거버넌스를 최종 경계획정과 분리하여 처리하는 방식이 당시 이미 관찰 가능한 지역적 경험이었음을 보여 준다. 그러나 참조할 수 있는 모델 자체가 협정의 수용 가능성을 결정하는 것은 아니다. Putnam은 국제협상의 결과가 여전히 국내적으로 수용 가능한 범위에 들어가야 하며, 수용 가능한 방안의 교집합이 작을수록 협정의 형성이 어려워진다고 지적한다.[[69]] 최종 경계는 공간, 자원 및 관할권의 장기적인 배분과 관련되는 반면, 제한적인 어업협정은 최종 입장을 유보한 채 현실적인 어업 사안만을 처리하므로 요구되는 정치적 약속이 상대적으로 제한적이다. 지역적 경험은 분리 처리 방안을 식별 가능한 선택으로 만들었고, 낮은 수준의 정치적 약속은 그 방안이 수용될 가능성을 높였다.",
    "제2조부터 제9조까지, 제13조 및 부속서 I은 양국의 합의를 차별화된 수역 관리, 연례 입어 조건, 공동위원회의 협의 및 양국 주무기관의 개별적 집행으로 구체화하였다.[[70]] 수역별 관리는 하나의 문서 안에서 최종 경계를 처리하는 것을 피하게 하고, 연례 조건은 관련 수치를 조정할 수 있는 여지를 남기며, 국내의 개별 집행은 각국의 행정적 통제를 유지한다. 여기서는 형성 조건과 제도적 결과를 구분할 필요가 있다. 앞서 살펴본 법적·거버넌스적·정치적 조건은 협정이 왜 형성될 수 있었는지를 설명하고, 이들 조항은 협정이 성립한 이후 어떠한 구조를 채택하였는지를 보여 준다. 따라서 제도 설계 자체를 거꾸로 형성 원인으로 간주해서는 안 된다.",
    "협정 제3조, 제8조, 제13조, 제14조 및 제16조는 UNCLOS의 잠정약정 규칙과 함께 연례 협의, 과도기간, 입장 유보 및 발효기간을 하나의 제도 안에 배치한다.[[71]] 이에 따라 양국은 최종적인 해양 권리에 관하여 양보하지 않으면서도 최소한의 어업관리 질서를 확보하였다. 이러한 결과는 단계적 균형으로 이해할 수 있다. 그 균형이 유지될 수 있었던 것은 어업 의제가 연례 협의를 통하여 계속 조정될 수 있었기 때문이며, 그 균형의 취약성 역시 동일한 조건에서 비롯된다. 최종 경계는 여전히 해결되지 않았고, 협력 대상은 어업에 집중되어 있으며, 고정시설, 일반적인 해양과학조사, 통항 및 진입 통제 등의 문제에는 동일하게 명확한 규칙과 권한이 부여되지 않았다. 현재의 자료는 이러한 구조적 경계만을 지적할 수 있을 뿐, 이를 근거로 제도가 이미 기능을 상실하였다고 단정할 수는 없다. 새로운 의제가 등장할 때 제도가 정체되는지 여부는 실제 운영과 국가 간 상호작용을 결합하여 판단할 필요가 있다.",
]

FOOTNOTES = {
    61: "United Nations Convention on the Law of the Sea, 1982, Arts. 74(3), 83(3); 《联合国海洋法公约》, 제74조 제3항 및 제83조 제3항.",
    62: "《中华人民共和国政府和大韩民国政府渔业协定》, 2000년 8월 3일, 제14조; 대한민국 정부와 중화인민공화국 정부간의 어업에 관한 협정, 제14조(한국어 정본은 동일 조항의 정본 용어를 대조하는 데에만 사용함).",
    63: "《中华人民共和国政府和大韩民国政府渔业协定》, 전문, 제1—3조 및 제7—14조.",
    64: "Ministry of Foreign Affairs of the People’s Republic of China, “China and ROK Signed A Fishery Agreement,” updated 15 November 2000, p. 1(중국 측의 서명 후 일방적 설명).",
    65: "호남지방통계청 목포사무소, 「지난 50년간(1970년~2019년) 전남 어업구조 변화상」, 2020, 5, 9쪽.",
    66: "国家统计局, 『海洋水产生产情况（1999年）』, 1999年12月15日.",
    67: "대한민국 국회, 제218회 국회 통일외교통상위원회 제1차 회의록, 2001년 2월 16일, 42쪽.",
    68: "박재영·최종화, 「한·중어업협정의 평가 및 향후과제」, 『수산경영론집』, 제31권 제2호, 2000, 69—70, 76—77, 79—83쪽.",
    69: "Robert D. Putnam, “Diplomacy and Domestic Politics: The Logic of Two-Level Games,” International Organization, Vol. 42, No. 3, 1988, pp. 435—438.",
    70: "《中华人民共和国政府和大韩民国政府渔业协定》, 제2—9조, 제13조 및 부속서 I; 대한민국 정부와 중화인민공화국 정부간의 어업에 관한 협정, 제2—9조, 제13조 및 부속서 I(한국어 정본은 독립적인 사건 가중치를 추가하지 않음).",
    71: "《中华人民共和国政府和大韩民国政府渔业协定》, 제3조, 제8조, 제13조, 제14조 및 제16조; United Nations Convention on the Law of the Sea, Arts. 74(3), 83(3).",
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def xml_text(text: str) -> str:
    return escape(text, {'"': '&quot;'})


def run_text(text: str, *, size: int | None = None, bold: bool = False) -> str:
    props = []
    if bold:
        props.append("<w:b/><w:bCs/>")
    if size is not None:
        props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xml_text(text)}</w:t></w:r>'


def footnote_ref(internal_id: int) -> str:
    return (
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
        f'<w:footnoteReference w:id="{internal_id}"/></w:r>'
    )


def body_runs(text: str) -> str:
    pieces = re.split(r"\[\[(\d+)\]\]", text)
    out: list[str] = []
    for idx, piece in enumerate(pieces):
        if idx % 2 == 0:
            if piece:
                out.append(run_text(piece))
        else:
            number = int(piece)
            if number not in FOOTNOTES:
                raise ValueError(f"Unknown footnote marker: {number}")
            out.append(footnote_ref(number - 60))
    return "".join(out)


def title_paragraph() -> str:
    return (
        '<w:p><w:pPr><w:keepNext/><w:keepLines/>'
        '<w:spacing w:before="0" w:after="240" w:line="384" w:lineRule="auto"/>'
        '<w:jc w:val="left"/></w:pPr>'
        f'{run_text(TITLE, size=24, bold=True)}</w:p>'
    )


def body_paragraph(text: str) -> str:
    return (
        '<w:p><w:pPr><w:widowControl/><w:jc w:val="both"/>'
        '<w:ind w:firstLine="420" w:firstLineChars="200"/>'
        '<w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/>'
        '</w:pPr>'
        f'{body_runs(text)}</w:p>'
    )


def make_document_xml() -> str:
    body = title_paragraph() + "".join(body_paragraph(p) for p in PARAGRAPHS)
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}">
  <w:body>
    {body}
    <w:sectPr>
      <w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="61"/><w:numRestart w:val="continuous"/><w:pos w:val="pageBottom"/></w:footnotePr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="312"/>
    </w:sectPr>
  </w:body>
</w:document>'''


def footnote_paragraph(internal_id: int, text: str) -> str:
    return (
        f'<w:footnote w:id="{internal_id}">'
        '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/><w:jc w:val="both"/>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
        f'{run_text(" " + text, size=18)}</w:p></w:footnote>'
    )


def make_footnotes_xml() -> str:
    notes = [
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for number in range(61, 72):
        notes.append(footnote_paragraph(number - 60, FOOTNOTES[number]))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W_NS}">{''.join(notes)}</w:footnotes>'''


CONTENT_TYPES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

PACKAGE_RELS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}">
  <Relationship Id="rId1" Type="{OFFICE_REL_NS}/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="{OFFICE_REL_NS}/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOCUMENT_RELS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL_NS}">
  <Relationship Id="rId1" Type="{OFFICE_REL_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="{OFFICE_REL_NS}/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="{OFFICE_REL_NS}/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="{OFFICE_REL_NS}/footnotes" Target="footnotes.xml"/>
</Relationships>'''

STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>
      <w:sz w:val="21"/><w:szCs w:val="21"/>
      <w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:line="384" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
    <w:pPr><w:widowControl/><w:jc w:val="both"/><w:spacing w:line="384" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
</w:styles>'''

SETTINGS_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W_NS}">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="61"/><w:numRestart w:val="continuous"/><w:pos w:val="pageBottom"/></w:footnotePr>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
  <w:doNotTrackMoves/><w:doNotTrackFormatting/>
</w:settings>'''

FONT_TABLE_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="{W_NS}">
  <w:font w:name="Batang"><w:charset w:val="81"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
  <w:font w:name="바탕"><w:charset w:val="81"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
</w:fonts>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion>
</Properties>'''


def make_core_xml() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{xml_text(TITLE)}</dc:title><dc:subject>Master's thesis Korean translation</dc:subject><dc:creator>OpenAI</dc:creator><cp:keywords>2.5; Korean translation; thesis</cp:keywords><dc:description>Faithful Korean translation of thesis section 2.5</dc:description><cp:lastModifiedBy>OpenAI</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def write_docx() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    members = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": PACKAGE_RELS,
        "docProps/core.xml": make_core_xml(),
        "docProps/app.xml": APP_XML,
        "word/document.xml": make_document_xml(),
        "word/styles.xml": STYLES_XML,
        "word/settings.xml": SETTINGS_XML,
        "word/fontTable.xml": FONT_TABLE_XML,
        "word/footnotes.xml": make_footnotes_xml(),
        "word/_rels/document.xml.rels": DOCUMENT_RELS,
    }
    with zipfile.ZipFile(DOCX_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in members.items():
            zf.writestr(name, content.encode("utf-8"))


def verify_docx() -> str:
    required = {
        "[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml",
        "word/document.xml", "word/styles.xml", "word/settings.xml", "word/fontTable.xml",
        "word/footnotes.xml", "word/_rels/document.xml.rels",
    }
    with zipfile.ZipFile(DOCX_PATH, "r") as zf:
        bad = zf.testzip()
        if bad is not None:
            raise RuntimeError(f"Corrupt ZIP member: {bad}")
        names = set(zf.namelist())
        missing = required - names
        if missing:
            raise RuntimeError(f"Missing DOCX members: {sorted(missing)}")
        document_bytes = zf.read("word/document.xml")
        footnotes_bytes = zf.read("word/footnotes.xml")
        settings_bytes = zf.read("word/settings.xml")

    ET.fromstring(document_bytes)
    ET.fromstring(footnotes_bytes)
    ET.fromstring(settings_bytes)

    ns = {"w": W_NS}
    document_root = ET.fromstring(document_bytes)
    footnotes_root = ET.fromstring(footnotes_bytes)
    paragraph_count = len(document_root.findall(".//w:body/w:p", ns))
    refs = [int(el.attrib[f"{{{W_NS}}}id"]) for el in document_root.findall(".//w:footnoteReference", ns)]
    actual_notes = [
        int(el.attrib[f"{{{W_NS}}}id"])
        for el in footnotes_root.findall("w:footnote", ns)
        if int(el.attrib[f"{{{W_NS}}}id"]) > 0
    ]
    doc_text = "".join(document_root.itertext())

    checks = {
        "paragraph_count_is_7": paragraph_count == 7,
        "footnote_reference_count_is_11": len(refs) == 11,
        "footnote_reference_ids_are_1_to_11": refs == list(range(1, 12)),
        "footnote_nodes_are_1_to_11": actual_notes == list(range(1, 12)),
        "title_present": TITLE in doc_text,
        "all_markers_resolved": "[[" not in doc_text and "]]" not in doc_text,
        "source_structure_preserved": len(PARAGRAPHS) == 6,
        "footnote_numbers_preserved": sorted(FOOTNOTES) == list(range(61, 72)),
    }
    failed = [name for name, ok in checks.items() if not ok]
    if failed:
        raise RuntimeError("Verification failed: " + ", ".join(failed))

    sha = hashlib.sha256(DOCX_PATH.read_bytes()).hexdigest()
    lines = [
        "RESULT=PASS",
        f"DOCX={DOCX_PATH.name}",
        f"DOCX_BYTES={DOCX_PATH.stat().st_size}",
        f"DOCX_SHA256={sha}",
        f"PARAGRAPHS={paragraph_count}",
        f"BODY_PARAGRAPHS={len(PARAGRAPHS)}",
        f"FOOTNOTE_REFERENCES={len(refs)}",
        "DISPLAY_FOOTNOTE_RANGE=61-71",
        "FORMAT=A4; Batang; body 10.5pt; justified; first-line indent; 1.6 line spacing",
    ]
    lines.extend(f"CHECK_{name}=PASS" for name in checks)
    return "\n".join(lines) + "\n"


def write_zip() -> None:
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(DOCX_PATH, arcname=DOCX_PATH.name)
        zf.write(VERIFY_PATH, arcname=VERIFY_PATH.name)
    with zipfile.ZipFile(ZIP_PATH, "r") as zf:
        bad = zf.testzip()
        if bad is not None:
            raise RuntimeError(f"Corrupt delivery ZIP member: {bad}")


def main() -> None:
    write_docx()
    verification = verify_docx()
    VERIFY_PATH.write_text(verification, encoding="utf-8")
    write_zip()
    zip_sha = hashlib.sha256(ZIP_PATH.read_bytes()).hexdigest()
    with VERIFY_PATH.open("a", encoding="utf-8") as fh:
        fh.write(f"ZIP={ZIP_PATH.name}\nZIP_BYTES={ZIP_PATH.stat().st_size}\nZIP_SHA256={zip_sha}\n")
    # Rebuild the ZIP so the verification file inside includes the ZIP metadata except its final self-hash.
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(DOCX_PATH, arcname=DOCX_PATH.name)
        zf.write(VERIFY_PATH, arcname=VERIFY_PATH.name)
    print(VERIFY_PATH.read_text(encoding="utf-8"), end="")


if __name__ == "__main__":
    main()
