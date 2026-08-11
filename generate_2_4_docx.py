#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate a Word-compatible DOCX for the faithful Korean translation of thesis section 2.4.

The package is created directly with Office Open XML so that footnotes 51–60
remain real Word footnotes and the thesis-template typography is preserved.
"""

from __future__ import annotations

import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

OUTPUT = Path("第二章_2.4_韩语忠实翻译稿.docx")
VERIFY = Path("verification_2_4.txt")

BLOCKS = [
    ("section", "2.4 한중 어업협정의 제도적 설계"),
    ("body", "2000년 협정은 모든 해역 사안을 하나의 제도에 포괄하려 하지 않았다. 협정은 수역을 구분하여 서로 다른 관리관계를 배치하고, 공동위원회를 통해 연례 협의를 유지한 뒤 양국 주무기관이 각각 허가와 행정관리를 실시하도록 하였다.[[51]] 이러한 설계는 양측이 이미 형성한 제한적 합의를 어업 분야에 고정하였다. 분석의 초점은 공간, 권한 및 절차가 어떻게 서로 연결되는지, 그리고 이러한 연결이 어떠한 사안에서 멈추는지에 있다."),
    ("subsection", "2.4.1 세 가지 수역의 법리적 구조 : 협정 수역, 잠정조치 수역, 과도 수역"),
    ("body", "제1조는 중국의 배타적경제수역과 한국의 배타적경제수역을 ‘협정수역’으로 규정하고, 제6조는 제7조, 제8조 및 제9조에서 정한 수역을 제외한 나머지 부분에 제2조부터 제5조까지의 일반 입어 규칙을 적용한다. 잠정조치수역과 양측 각자의 과도수역은 전체 협정수역 안에 위치하지만 특별규칙의 적용을 받으며, 제9조는 그 밖의 관련 수역에 대해서도 양측이 별도로 합의하지 않는 동안 현행 어업활동을 유지하도록 정한다.[[52]] 이러한 공간적 배치가 제도적 구분을 이루는 이유는 관리관계가 서로 다르기 때문이다. 일반수역에서는 연안측이 상대방의 국민과 어선에 입어를 허가하고, 잠정조치수역에서는 어업공동위원회의 결정에 따라 공동 보존조치와 수량관리를 시행하는 동시에 양측이 각각 자국의 국민과 어선을 관리한다. 과도수역에서는 자국의 허가증 발급, 명부 교환, 공동 보존 및 가능한 공동 감시·감독을 통해 기존 조업을 일반 입어 규칙에 단계적으로 편입한다.[[53]] 협정수역은 전체 범위를 구성하고, 잠정조치수역, 과도수역 및 제9조 관련 수역에는 각각 서로 다른 권한과 책임이 배치된다. 따라서 세 가지 주요 수역은 공식 설계를 이해하는 기본 단위이지만 모든 공간 조항을 망라하지는 않는다."),
    ("body", "제8조는 과도수역의 특별규칙을 협정 발효 후 4년으로 한정하고, 그 기간이 만료된 후에는 제2조부터 제5조까지를 적용하도록 규정한다. 제9조 관련 수역에서는 양측이 별도로 합의할 때까지 특정한 약정을 계속 유지한다.[[54]] 4년의 기간은 문서가 규칙의 전환을 예정하였음을 보여 줄 뿐, 기간 만료 후 실제 조업, 허가 또는 단속의 상태를 입증하지 않는다. 제14조는 동시에 해양법 문제에 관한 양측의 입장을 유보한다. 따라서 이들 수역과 좌표는 협정 내부의 어업관리상 의미만을 가지며, 최종 배타적경제수역 경계선, 해양경계 또는 주권선을 구성하지 않는다. 공간 구분의 법적 의미는 시간적 배치와 입장 유보에 의해 함께 한정된다."),
    ("subsection", "2.4.2 한중 어업공동위원회와 연례 입어 조건 메커니즘"),
    ("body", "협정의 관리절차에 관한 배치는 수역 구분에 그치지 않는다. 제2조부터 제5조까지는 일반 입어의 허가와 법규 준수 관계를 규정하고, 잠정조치수역에서는 각자가 자국 어선을 관리하고 상호 통보하는 점을 강조하며, 과도수역에서는 자국의 허가증 발급, 명부 교환 및 가능한 공동 감시·감독을 추가한다. 부속서 I은 서면 신청, 허가증 발급, 어선 표지, 조업일지 및 통계자료를 일반 입어절차에 포함한다.[[55]] 이러한 양자 규칙은 다시 국내 행정체계에 들어가야 한다. 중국 농업농촌부의 시행방법은 서로 다른 층위의 주무기관에 신청 심사, 자격 심사, 허가증, 명부, 일지, 통계 및 선박 표지 등의 직무를 배분한다. 한국 해양수산부의 2025년도 어업인 참고자료는 한국 어선이 중국의 배타적경제수역에 입어할 때 신청의 전달, 수역 입·출역 통보, 일지의 취합, 표지 및 준수사항에 관한 안내를 설명한다.[[56]] 두 자료의 대상과 시기는 서로 대칭적이지 않으므로, 양측이 공동규칙을 국내 행정체계에서 어떻게 구현하는지를 각각 설명할 수 있을 뿐이다. 이를 동일한 하나의 제도적 사실로 합칠 수 없으며, 관련 절차가 실제로 집행되었음을 입증할 수도 없다."),
    ("body", "어업공동위원회는 양자 협의와 국가 집행을 연결하는 상설 조정기관이다. 제13조는 위원회가 양측 대표와 위원으로 구성되며, 필요한 경우 전문가분과위원회를 설치할 수 있다고 규정한다. 위원회는 어획가능어종, 할당량, 조업조건, 조업질서, 자원 보존 및 어업협력에 관하여 협의하고 권고하며, 제7조 및 제8조 수역에 관한 사항인 경우에만 협의하고 결정한다. 모든 권고와 결정은 양측 대표의 합의에 의해서만 이루어져야 한다.[[57]] 제3조는 연간 어획가능어종, 할당량, 조업기간, 조업수역 및 그 밖의 조건에 관한 결정과 통보를 각 당사자에게 맡기며, 각 당사자가 결정을 내릴 때 관련 요소를 고려하고 위원회의 협의 결과를 존중하도록 한다.[[58]] 따라서 위원회는 양국 주무기관을 대체하지 않는다. 위원회는 양측의 연례 협의를 유지하고, 이어 각국 기관이 협의 결과를 자국의 허가와 조업규모 배분으로 전환하도록 한다. 서명이 완비된 회의록과 서로 대응하는 연도별 문서는 아직 확보되지 않았기 때문에, 특정 연도의 공동 선박 수, 할당량 및 협의과정을 완전하게 복원하기 어렵다. 공식적 배치에 따르면 위원회의 협의 결과는 여전히 각 당사자의 연례 통보와 국내 허가절차를 거쳐야 실제 관리에 편입될 수 있다."),
    ("subsection", "2.4.3 어업 규칙의 정밀성과 비어업 공간 이용의 규칙 공백"),
    ("body", "제1조부터 부속서 I에 이르기까지 협정은 어업관리의 내용을 적용수역, 국민과 어선, 허가 주체, 어획가능어종, 할당량, 조업기간과 조업수역, 자원 보존, 조업질서, 억류 후의 통보와 석방, 그리고 어업공동위원회의 절차로 점차 구체화한다. 서면 신청, 허가증 발급, 통계자료, 어선 표지 및 조업일지는 이러한 요구를 다시 운영 단계에 구현한다.[[59]] 관리대상, 권한, 조건, 절차 및 기간은 서로 연계될 수 있으므로, 공식문서는 누가 어떠한 수역에서 어떠한 조건에 따라 어떠한 기관을 통하여 어업활동을 하는지에 답할 수 있다. 여기서 구체성은 규칙을 식별할 수 있음을 의미할 뿐, 허가, 일지, 통계 및 단속이 이미 효과적으로 운영되고 있음을 입증하지 않는다."),
    ("body", "제10조는 항행 및 조업의 안전, 해상 조업질서와 사고 처리에 관한 사항을 다루고, 제11조는 해난구조, 긴급피난 및 통보를 규정한다. 부속서 II는 연락부서, 연락방법 및 선박정보를 열거하며, 제12조는 양측이 해양생물자원의 보존과 합리적 이용을 위한 과학연구 협력을 강화하도록 요구한다.[[60]] 이 조항들은 안전, 피난 및 생물자원에 관한 과학연구를 위해 제한적인 연결 지점을 남겨 둔다. 그러나 고정시설의 설치, 일반적인 해양과학조사, 항행 또는 출입 통제에는 어업 허가, 할당량 및 자원 보존과 같은 수준의 구체적인 규칙이 마련되어 있지 않다. 따라서 협정은 명확한 사안의 경계를 형성하였다. 협정은 어업과 제한적인 인접 사안에 비교적 구체적인 규칙을 배치하였지만, 이를 토대로 일반적인 해역 거버넌스 제도를 수립하지는 않았다. 이 경계는 협정 형성 당시 이미 이루어진 선택이며, 그 이유는 당시의 법적 공간, 거버넌스 압력 및 정치적 수용 조건과 결부하여 이해해야 한다."),
]

FOOTNOTES = {
    51: "《中华人民共和国政府和大韩民国政府渔业协定》, 2000년 8월 3일, 전문, 제1조, 제14—16조 및 서명란.",
    52: "《中华人民共和国政府和大韩民国政府渔业协定》, 제1조, 제6—9조, 제14조; 대한민국 정부와 중화인민공화국 정부간의 어업에 관한 협정, 제1조, 제6—9조, 제14조(한국어 자료에서 확인 가능한 조문에 한하여 용어를 대조함).",
    53: "《中华人民共和国政府和大韩民国政府渔业协定》, 제2—8조.",
    54: "《中华人民共和国政府和大韩民国政府渔业协定》, 제8조 제1항·제5항, 제9조, 제14조.",
    55: "《中华人民共和国政府和大韩民国政府渔业协定》, 제2—8조 및 부속서 I.",
    56: "中华人民共和国农业农村部, 《中韩渔业协定暂定措施水域和过渡水域管理办法》(2001년 2월 16일 농업부령 제47호로 공포, 2022년 1월 7일 농업농촌부령 2022년 제1호로 개정), 제2—3조, 제5—12조; 대한민국 해양수산부 편, 『중화인민공화국 배타적경제수역에서 대한민국 어선의 조업조건 및 입어절차』(어업인 참고용), 2025, Ⅱ 「중화인민공화국 배타적경제수역에서의 대한민국 어선의 입어에 관한 절차규칙」 제1·2·6—8·10조 및 Ⅲ 참고자료 2.",
    57: "《中华人民共和国政府和大韩民国政府渔业协定》, 제13조; 대한민국 정부와 중화인민공화국 정부간의 어업에 관한 협정, 제13조(한국어 자료에서 확인 가능한 조문에 한하여 용어를 대조함).",
    58: "《中华人民共和国政府和大韩民国政府渔业协定》, 제3조, 제13조 및 부속서 I; 中华人民共和国农业农村部, 《中韩渔业协定暂定措施水域和过渡水域管理办法》, 제4조; 대한민국 해양수산부 편, 『중화인민공화국 배타적경제수역에서 대한민국 어선의 조업조건 및 입어절차』(어업인 참고용), 2025, Ⅰ 및 Ⅲ 참고자료 1.",
    59: "《中华人民共和国政府和大韩民国政府渔业协定》, 전문, 제1—13조 및 부속서 I.",
    60: "《中华人民共和国政府和大韩民国政府渔业协定》, 제10—12조 및 부속서 II; 대한민국 정부와 중화인민공화국 정부간의 어업에 관한 협정, 제10—12조 및 부속서 II(한국어 자료에서 확인 가능한 문구에 한하여 용어를 대조하였으며, 결손된 부분은 보충하지 않음).",
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MARKER_RE = re.compile(r"\[\[(\d+)\]\]")


def text_run(text: str, size: int = 21, bold: bool = False) -> str:
    if not text:
        return ""
    b = "<w:b/><w:bCs/>" if bold else ""
    return (
        "<w:r><w:rPr>"
        '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/>'
        f"{b}<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>"
        '<w:lang w:val="en-US" w:eastAsia="ko-KR"/>'
        "</w:rPr>"
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
    )


def footnote_reference(fid: int) -> str:
    return (
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
        f'<w:footnoteReference w:id="{fid}"/></w:r>'
    )


def runs_with_footnotes(text: str) -> str:
    parts: list[str] = []
    pos = 0
    for match in MARKER_RE.finditer(text):
        parts.append(text_run(text[pos:match.start()]))
        parts.append(footnote_reference(int(match.group(1))))
        pos = match.end()
    parts.append(text_run(text[pos:]))
    return "".join(parts)


def make_document_xml() -> str:
    body_parts: list[str] = []
    for kind, text in BLOCKS:
        if kind == "section":
            style, size, bold = "SectionHeading", 26, True
        elif kind == "subsection":
            style, size, bold = "SubsectionHeading", 23, True
        else:
            style, size, bold = "BodyTextKorean", 21, False
        body_parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
            + (runs_with_footnotes(text) if kind == "body" else text_run(text, size=size, bold=bold))
            + "</w:p>"
        )

    sect_pr = """
<w:sectPr>
  <w:footnotePr>
    <w:numFmt w:val="decimal"/>
    <w:numStart w:val="51"/>
    <w:numRestart w:val="continuous"/>
    <w:pos w:val="pageBottom"/>
  </w:footnotePr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1440" w:right="1701" w:bottom="1440" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>
  <w:cols w:space="720"/>
  <w:docGrid w:linePitch="360"/>
</w:sectPr>
"""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>'
        + "".join(body_parts)
        + sect_pr
        + "</w:body></w:document>"
    )


def make_footnotes_xml() -> str:
    items = [
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for fid in range(51, 61):
        items.append(
            f'<w:footnote w:id="{fid}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            '<w:r><w:tab/></w:r>'
            + text_run(FOOTNOTES[fid], size=18)
            + "</w:p></w:footnote>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{W_NS}">' + "".join(items) + "</w:footnotes>"
    )


def make_styles_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/>
      <w:sz w:val="21"/><w:szCs w:val="21"/>
      <w:lang w:val="en-US" w:eastAsia="ko-KR"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/><w:widowControl/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="SectionHeading">
    <w:name w:val="Section Heading"/><w:basedOn w:val="Normal"/><w:next w:val="BodyTextKorean"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="0" w:after="240" w:line="360" w:lineRule="auto"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/><w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="SubsectionHeading">
    <w:name w:val="Subsection Heading"/><w:basedOn w:val="Normal"/><w:next w:val="BodyTextKorean"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="left"/><w:outlineLvl w:val="2"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/><w:b/><w:bCs/><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="BodyTextKorean">
    <w:name w:val="Korean Thesis Body"/><w:basedOn w:val="Normal"/><w:next w:val="BodyTextKorean"/><w:qFormat/>
    <w:pPr><w:widowControl/><w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/><w:ind w:firstLine="420" w:firstLineChars="200"/><w:jc w:val="both"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:next w:val="FootnoteText"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
    <w:pPr><w:spacing w:before="0" w:after="0" w:line="288" w:lineRule="auto"/><w:ind w:left="360" w:hanging="360"/><w:jc w:val="both"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang" w:cs="Batang"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>
  </w:style>
</w:styles>'''


def make_settings_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W_NS}">
  <w:zoom w:percent="100"/>
  <w:proofState w:spelling="clean" w:grammar="clean"/>
  <w:defaultTabStop w:val="420"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="51"/><w:numRestart w:val="continuous"/></w:footnotePr>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''


def make_font_table_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="{W_NS}"><w:font w:name="Batang"><w:altName w:val="바탕"/><w:family w:val="roman"/><w:charset w:val="81"/><w:pitch w:val="variable"/></w:font></w:fonts>'''


def make_core_xml() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    title = BLOCKS[0][1]
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{escape(title)}</dc:title><dc:subject>석사학위논문 제2장 2.4절 한국어 충실 번역문</dc:subject>
  <dc:creator>OpenAI</dc:creator><cp:lastModifiedBy>OpenAI</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
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

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

DOCUMENT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>'''

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion>
  <Pages>0</Pages><Words>0</Words><Characters>0</Characters><Lines>0</Lines><Paragraphs>11</Paragraphs>
  <Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>
</Properties>'''


def write_docx() -> None:
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": make_core_xml(),
        "docProps/app.xml": APP_XML,
        "word/document.xml": make_document_xml(),
        "word/styles.xml": make_styles_xml(),
        "word/settings.xml": make_settings_xml(),
        "word/fontTable.xml": make_font_table_xml(),
        "word/footnotes.xml": make_footnotes_xml(),
        "word/_rels/document.xml.rels": DOCUMENT_RELS,
    }
    with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value.encode("utf-8"))


def validate_docx() -> str:
    expected_parts = {
        "[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml",
        "word/document.xml", "word/styles.xml", "word/settings.xml", "word/fontTable.xml",
        "word/footnotes.xml", "word/_rels/document.xml.rels",
    }
    with zipfile.ZipFile(OUTPUT, "r") as archive:
        names = set(archive.namelist())
        missing = expected_parts - names
        if missing:
            raise RuntimeError(f"Missing DOCX parts: {sorted(missing)}")
        bad = archive.testzip()
        if bad:
            raise RuntimeError(f"Corrupt ZIP member: {bad}")
        for name in expected_parts:
            ET.fromstring(archive.read(name))
        document = ET.fromstring(archive.read("word/document.xml"))
        footnotes = ET.fromstring(archive.read("word/footnotes.xml"))

    ns = {"w": W_NS}
    refs = [int(node.attrib[f"{{{W_NS}}}id"]) for node in document.findall(".//w:footnoteReference", ns)]
    actual = [int(node.attrib[f"{{{W_NS}}}id"]) for node in footnotes.findall("w:footnote", ns) if int(node.attrib[f"{{{W_NS}}}id"]) > 0]
    paragraphs = document.findall(".//w:body/w:p", ns)
    text = "".join((node.text or "") for node in document.findall(".//w:t", ns))

    if refs != list(range(51, 61)):
        raise RuntimeError(f"Unexpected footnote references: {refs}")
    if actual != list(range(51, 61)):
        raise RuntimeError(f"Unexpected footnote definitions: {actual}")
    if len(paragraphs) != 11:
        raise RuntimeError(f"Unexpected paragraph count: {len(paragraphs)}")
    if "[[" in text or "]]" in text:
        raise RuntimeError("Unresolved footnote marker found")
    for heading in (BLOCKS[0][1], BLOCKS[2][1], BLOCKS[5][1], BLOCKS[8][1]):
        if heading not in text:
            raise RuntimeError(f"Heading missing: {heading}")

    return (
        "VALIDATION=PASS\n"
        f"FILE={OUTPUT.name}\n"
        f"SIZE_BYTES={OUTPUT.stat().st_size}\n"
        "STRUCTURE=1_MAIN_HEADING_PLUS_3_SUBHEADINGS_PLUS_7_BODY_PARAGRAPHS\n"
        "FOOTNOTES=10_REAL_WORD_FOOTNOTES_NUMBERED_51_TO_60\n"
        "PAGE=A4_PORTRAIT\n"
        "BODY_FONT=BATANG_10.5PT\n"
        "BODY_ALIGNMENT=JUSTIFIED\n"
        "FIRST_LINE_INDENT=2_CHARACTERS\n"
        "LINE_SPACING=1.6\n"
    )


if __name__ == "__main__":
    write_docx()
    report = validate_docx()
    VERIFY.write_text(report, encoding="utf-8")
    print(report)
