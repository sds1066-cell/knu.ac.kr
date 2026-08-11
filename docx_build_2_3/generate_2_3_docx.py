from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile
import xml.etree.ElementTree as ET

OUTPUT = Path("output/第二章_2.3_韩语忠实翻译稿.docx")
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

# Each paragraph is a sequence of text strings and integer footnote references.
TITLE = "2.3 한중 해양 경계 미획정의 역사적 맥락과 협정 협상 배경"

PARAGRAPHS: list[list[str | int]] = [
    [
        "황해의 면적은 약 75만 제곱킬로미터이며, 가장 넓은 곳은 약 300해리, 가장 좁은 곳은 약 104해리이다.", 26,
        " UNCLOS는 연안국이 200해리를 초과하지 않는 배타적 경제수역을 주장할 수 있도록 허용하고, 대륙붕에 대한 권리의 기본 범위를 규정한다.", 27,
        " 한중 양국의 해안은 서로 마주 보고 있으므로, 이러한 공간적 규모에서 양측이 배타적 경제수역과 대륙붕에 관한 주장을 제기할 경우 구조적 중첩이 발생한다. 이러한 중첩은 협의를 필요하게 하지만, 경계의 위치를 미리 결정하지 않으며 어떠한 사안을 우선적으로 처리해야 하는지도 규정하지 않는다."
    ],
    [
        "한국 외교부는 한중 어업협정 협상이 1993년 12월에 시작된 것으로 기록하고 있으며, 이후 2000년 협정 서명, 2001년 한국의 비준 및 발효 준비로 이어졌다.", 28,
        " 양국이 UNCLOS 당사국이 된 것과 각각 배타적 경제수역 관련 법률을 공포한 것은 모두 협상이 시작된 이후였다.", 29, 30, 31,
        " 이러한 시간적 순서는 법적 변화가 협상을 직접 개시하였다는 설명을 배제한다. 그보다는 이러한 변화가 기존 협상의 표현 방식을 바꾸어, 양측이 배타적 경제수역의 권리, 관할권 및 경계획정 등의 개념을 통해 기존의 이견을 새롭게 표현할 수 있도록 하였을 가능성이 더 크다."
    ],
    [
        "법률상의 표현이 더욱 명확해졌지만, 경계획정 원칙에 관한 양측의 이견이 해소된 것은 아니었다. 중국 외교부는 중국 측 입장을 형평의 원칙에 따라 모든 관련 사정을 고려하여 경계획정 문제를 해결해야 한다는 것으로 요약하였다. 동시에 한국 측이 중간선을 기초로 해야 한다고 주장하였으며, 양측이 단계별·해역별 추진 방식도 논의하였다고 기록하였다.", 32,
        " 한국의 「배타적 경제수역법」은 관련 경계가 국제법에 따라 합의로 획정되어야 하며, 합의가 없는 경우 한국은 중간선 너머에서 배타적 경제수역에 관한 권리를 행사하지 않는다고 규정한다.", 33,
        " 반면 중국의 법률은 형평의 원칙과 국제법에 기초하여 합의를 통해 중첩수역의 경계를 획정할 것을 요구한다.", 34,
        " 이러한 법률상 공식은 양측의 출발점을 설명할 수 있을 뿐, 최종 경계를 직접 도출할 수는 없다. Suk Kyoon Kim은 더 나아가 등거리 방식과 자연연장론을 둘러싼 이견이 직선기선, 이어도, 어업 및 광물 이익과도 상호 연계되어 있다고 지적한다.", 35,
        " 이석룡 역시 양국이 등거리, 자연연장 및 형평한 결과를 실현하는 방식에 대해 서로 다르게 이해하고 있음을 논의하였다.", 36,
        " 이에 따라 최종 경계획정은 여러 법적 쟁점과 이익 문제를 함께 다루는 협상이 되었으며, 단일한 공식으로 신속하게 해결하기 어렵게 되었다."
    ],
    [
        "이에 따라 최종 경계획정과 어업협정은 서로 다른 속도로 진행되었다. 2000년에 이르러서도 양측은 해양경계획정을 추진하는 방식에 관하여 계속 논의하고 있었다.", 37,
        " 반면 1993년에 시작된 어업협정 협상은 이미 서명, 비준 및 발효 준비 단계에 들어서 있었다.", 38,
        " 이는 특정한 한 차례의 경계획정 난항이 특정 어업규칙을 직접 만들어 냈음을 의미하지 않는다. 보다 합리적인 설명은 최종 경계가 장기적인 권리 배분과 관련되기 때문에 협상은 지속될 수 있었던 반면, 어업활동은 허가, 단속 및 자원관리 문제를 계속 발생시켰다는 것이다. 두 사안이 요구하는 시간은 서로 달랐으며, 그 결과 제한적인 어업약정이 먼저 추진될 수 있었다."
    ],
    [
        "한국의 공식 통계자료에 따르면 전국 어선 수는 1995년 76,801척, 2000년 95,890척이었다(통계 범위는 한국 전국이며 황해 조업어선 수가 아니다).", 39,
        " 중국 국가통계국이 공표한 황해의 해양어획량은 1995년 1,706,250톤, 1999년 3,477,667톤이었다(통계 대상은 해양어획량이다).", 40,
        " 통계 지역, 연도, 대상 및 단위가 서로 다르므로 두 자료는 한중 양국의 어업 규모를 비교하는 데 사용하는 것이 적절하지 않으며, 협상이 왜 시작되었는지를 단독으로 설명할 수도 없다. 다만 두 자료는 모두 1990년대 중·후반에도 어업활동이 지속되었으며, 관련 관리가 경계미획정 때문에 중단되지 않았음을 보여 준다. 나포, 단속 과정에서의 접촉 및 법률 적용을 둘러싼 분쟁이 계속 발생하면서, 지속적인 조업은 점차 생산활동에서 정부 간 관리의 문제로 전환되었다. Sun Pyo Kim은 한중 어선의 나포, 단속 과정에서의 접촉, 한국의 배타적 경제수역 어업법 적용을 둘러싼 분쟁 및 단속 방안에 관한 논의를 기록하였다.", 41,
        " 협정이 서명된 뒤 한국 국회는 순찰 및 단속 배치에 관하여 해양경찰청에 질의하기도 하였다.", 42,
        " Sun Pyo Kim의 서술은 협상과 관련된 시기의 관리 압력을 반영하며, 한국 국회의 질의는 협정 발효에도 행정적 집행 역량이 필요하였음을 보여 준다. 협상 기간의 관리 압력과 협정 발효 이후의 집행 수요는 모두 어업 사안을 최종 경계획정이 완료될 때까지 장기간 미루어 둘 수 없었음을 보여 준다."
    ],
    [
        "최종적으로 제한적인 어업방안이 형성된 이유를 설명하려면 정치적 수용 조건도 고려해야 한다. Putnam의 양면게임(two-level game) 이론은 국제협상의 결과가 국내적으로 수용 가능한 범위 안에 들어가야 하며, 수용 가능한 방안의 교집합이 작을수록 협정이 형성되기 어렵다고 지적한다.", 43,
        " 최종 경계는 공간, 자원 및 관할권의 장기적 배분과 관련되므로, 의제를 한정하고 최종 입장을 유보하는 어업약정보다 일반적으로 더 높은 수준의 정치적 약속을 요구한다. 양국 정부는 최종적으로 공동문서를 마련하였고, 한국 국회는 그 뒤 비준 절차를 완료하였다.", 44, 45,
        " 이러한 사실은 정부 간 차원과 한국의 공식 절차 차원에서 수용이 이루어졌음을 뒷받침하지만, 양국의 국내 협의과정 전체를 복원하기에는 충분하지 않다."
    ],
    [
        "협정 자체는 제한적인 방안의 수용 가능성을 더욱 분명히 보여 준다. 공동문서는 어업관계만을 규율하며, 양측 각자의 해양법상 입장에 영향을 미치지 않는다고 명시한다.", 46,
        " 한국의 공식자료에는 협정의 서명, 비준 및 발효 준비가 기록되어 있다.", 47,
        " 반면 중국 측 공개자료는 어업협정과 계속 진행 중이던 해양경계획정 협의를 나란히 기록하고 있다.", 48,
        " 협정 제1조, 제14조, 제16조 및 서명면은 어업약정, 입장 유보 및 발효 절차를 동일한 문서 구조 안에 배치한다.", 49, 50,
        " 양측은 동일한 협정에서 최종 권리 배분까지 완료할 필요가 없었기 때문에, 제한적인 합의는 먼저 어업관리 분야에서 형성될 수 있었다."
    ],
    [
        "최종 경계획정 협상이 느리게 진행된 것은 방법, 지리 및 이익 배분 등의 문제가 서로 얽혀 있었기 때문이다. 반면 어업활동은 관리와 단속에 대한 수요를 계속 발생시켰다. 서로 다른 두 가지 속도가 병행되면서 양측은 어업을 최종 권리 배분으로부터 잠정적으로 분리하였다. 협정은 먼저 협력 대상과 입장 유보 문제를 해결하였다. 다만 제한적인 합의가 구체적인 수역 구분, 관리 권한 및 협의 절차로 구체화될 때에만 집행 가능한 제도적 장치가 될 수 있다."
    ],
]

FOOTNOTES: dict[int, str] = {
    26: 'Seokwoo Lee and Clive Schofield, "China and South Korea\'s Maritime Boundary in the Yellow Sea: A Preliminary Analysis," Korean Journal of International and Comparative Law, Vol. 13, 2025, pp. 6–7.',
    27: 'United Nations Convention on the Law of the Sea, Arts. 57 and 76(1).',
    28: 'Ministry of Foreign Affairs of the Republic of Korea, "The Korea-China Fisheries Agreement Will Come into Force on June 30, 2001," Press Release, 29 June 2001, paras. 1–4 (비공식 영문 번역본).',
    29: 'United Nations Treaty Collection, "United Nations Convention on the Law of the Sea," participant status records for the Republic of Korea and China; United Nations Division for Ocean Affairs and the Law of the Sea, "Chronological Lists of Ratifications of, Accessions and Successions to the Convention and the Related Agreements."',
    30: 'United Nations, Law of the Sea Bulletin, No. 33, Republic of Korea, Exclusive Economic Zone Act, Act No. 5151 of 8 August 1996, Arts. 1, 2 and 5, pp. 52–53.',
    31: '中华人民共和国专属经济区和大陆架法, 1998년 6월 26일, 제1—2조.',
    32: 'Ministry of Foreign Affairs of the People\'s Republic of China, "III. China\'s Maritime Demarcation and Bilateral Fishery Affairs," 9 July 2001, sec. III(3).',
    33: 'United Nations, Law of the Sea Bulletin, No. 33, Republic of Korea, Exclusive Economic Zone Act, Act No. 5151 of 8 August 1996, Arts. 1, 2 and 5, pp. 52–53.',
    34: '中华人民共和国专属经济区和大陆架法, 1998년 6월 26일, 제1—2조.',
    35: 'Suk Kyoon Kim, "Maritime Boundary Negotiations between China and Korea: The Factors at Stake," The International Journal of Marine and Coastal Law, Vol. 32, No. 1, 2017, PDF pp. 1–8.',
    36: '이석룡, 「우리나라와 중국간 해양경계획정」, 『국제법학회논총』, 제52권 제2호, 2007, 262—263쪽.',
    37: 'Ministry of Foreign Affairs of the People\'s Republic of China, "III. China\'s Maritime Demarcation and Bilateral Fishery Affairs," 9 July 2001, sec. III(3).',
    38: 'Ministry of Foreign Affairs of the Republic of Korea, "The Korea-China Fisheries Agreement Will Come into Force on June 30, 2001," Press Release, 29 June 2001, paras. 1–4 (비공식 영문 번역본).',
    39: '호남지방통계청 목포사무소, 「지난 50년간(1970년~2019년) 전남 어업구조 변화상」, 2020, 9쪽.',
    40: '国家统计局, 「海洋水产生产情况（1999年）」, 1999년 12월 15일.',
    41: 'Sun Pyo Kim, "The UN Convention on the Law of the Sea and New Fisheries Agreements in North East Asia," Marine Policy, Vol. 27, 2003, pp. 101–102.',
    42: '대한민국 국회, 제218회 농림해양수산위원회회의록 제3호, 2001년 2월 20일, 52—53쪽.',
    43: 'Robert D. Putnam, "Diplomacy and Domestic Politics: The Logic of Two-Level Games," International Organization, Vol. 42, No. 3, 1988, pp. 435–438.',
    44: '《中华人民共和国政府和大韩民国政府渔业协定》, 2000년 8월 3일, 제1조, 제14조, 제16조 및 서명면.',
    45: '대한민국 국회, 제218회 통일외교통상위원회회의록 제3호, 2001년 2월 27일, 29—30쪽; 대한민국 국회, 제219회 국회본회의회의록 제6호, 2001년 2월 28일, 25쪽.',
    46: '《中华人民共和国政府和大韩民国政府渔业协定》, 2000년 8월 3일, 제1조, 제14조, 제16조 및 서명면.',
    47: 'Ministry of Foreign Affairs of the Republic of Korea, "The Korea-China Fisheries Agreement Will Come into Force on June 30, 2001," Press Release, 29 June 2001, paras. 1–4 (비공식 영문 번역본).',
    48: 'Ministry of Foreign Affairs of the People\'s Republic of China, "III. China\'s Maritime Demarcation and Bilateral Fishery Affairs," 9 July 2001, sec. III(3).',
    49: '《中华人民共和国政府和大韩民国政府渔业协定》, 2000년 8월 3일, 제1조, 제14조, 제16조 및 서명면.',
    50: '《中华人民共和国政府和大韩民国政府渔业协定》, 2000년 8월 3일, 제1조, 제14조, 제16조 및 서명면.',
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def xml_decl(body: str) -> str:
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body


def text_run(text: str) -> str:
    return (
        '<w:r><w:rPr>'
        '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
        '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>'
        '</w:rPr><w:t xml:space="preserve">' + escape(text) + '</w:t></w:r>'
    )


def footnote_ref_run(fid: int) -> str:
    return (
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
        f'<w:footnoteReference w:id="{fid}"/></w:r>'
    )


def body_paragraph(parts: list[str | int]) -> str:
    runs: list[str] = []
    for part in parts:
        if isinstance(part, int):
            runs.append(footnote_ref_run(part))
        else:
            runs.append(text_run(part))
    return (
        '<w:p><w:pPr><w:pStyle w:val="BodyTextKorean"/></w:pPr>'
        + ''.join(runs)
        + '</w:p>'
    )


def heading_paragraph(text: str) -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="SectionHeading"/></w:pPr>'
        '<w:r><w:rPr>'
        '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
        '<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>'
        '</w:rPr><w:t xml:space="preserve">' + escape(text) + '</w:t></w:r></w:p>'
    )


def footnote_paragraph(text: str) -> str:
    return (
        '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
        '<w:r><w:tab/></w:r>'
        '<w:r><w:rPr>'
        '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
        '<w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>'
        '</w:rPr><w:t xml:space="preserve">' + escape(text) + '</w:t></w:r></w:p>'
    )


document_xml = xml_decl(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:xml="http://www.w3.org/XML/1998/namespace">'
    '<w:body>'
    + heading_paragraph(TITLE)
    + ''.join(body_paragraph(p) for p in PARAGRAPHS)
    + '<w:sectPr>'
      '<w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="26"/><w:numRestart w:val="continuous"/></w:footnotePr>'
      '<w:pgSz w:w="11906" w:h="16838"/>'
      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
      '<w:cols w:space="425"/>'
    '</w:sectPr>'
    '</w:body></w:document>'
)

styles_xml = xml_decl(
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults>'
      '<w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
        '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>'
      '</w:rPr></w:rPrDefault>'
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="336" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    '</w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
      '<w:name w:val="Normal"/><w:qFormat/>'
      '<w:pPr><w:spacing w:after="0" w:line="336" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr>'
      '<w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
      '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="BodyTextKorean">'
      '<w:name w:val="Korean thesis body"/><w:basedOn w:val="Normal"/><w:next w:val="BodyTextKorean"/><w:qFormat/>'
      '<w:pPr><w:widowControl/><w:spacing w:after="0" w:line="336" w:lineRule="auto"/>'
      '<w:ind w:firstLine="420" w:firstLineChars="200"/><w:jc w:val="both"/></w:pPr>'
      '<w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
      '<w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="SectionHeading">'
      '<w:name w:val="Section heading"/><w:basedOn w:val="Normal"/><w:next w:val="BodyTextKorean"/><w:qFormat/>'
      '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="0" w:after="180"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>'
      '<w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
      '<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="FootnoteText">'
      '<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:next w:val="FootnoteText"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
      '<w:pPr><w:tabs><w:tab w:val="left" w:pos="360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
      '<w:ind w:left="360" w:hanging="360"/><w:jc w:val="both"/></w:pPr>'
      '<w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>'
      '<w:sz w:val="18"/><w:szCs w:val="18"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="character" w:styleId="FootnoteReference">'
      '<w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
      '<w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
    '</w:style>'
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
    '</w:styles>'
)

settings_xml = xml_decl(
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
      '<w:zoom w:percent="100"/>'
      '<w:defaultTabStop w:val="720"/>'
      '<w:characterSpacingControl w:val="doNotCompress"/>'
      '<w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="26"/><w:numRestart w:val="continuous"/></w:footnotePr>'
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    '</w:settings>'
)

footnote_items = [
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
]
for fid in range(26, 51):
    footnote_items.append(f'<w:footnote w:id="{fid}">' + footnote_paragraph(FOOTNOTES[fid]) + '</w:footnote>')

footnotes_xml = xml_decl(
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + ''.join(footnote_items)
    + '</w:footnotes>'
)

content_types_xml = xml_decl(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '</Types>'
)

root_rels_xml = xml_decl(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)

document_rels_xml = xml_decl(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
    '</Relationships>'
)

now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
core_xml = xml_decl(
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:title>제2장 2.3 한국어 번역문</dc:title>'
    '<dc:subject>한중 해양 경계 미획정의 역사적 맥락과 협정 협상 배경</dc:subject>'
    '<dc:creator>OpenAI</dc:creator><cp:lastModifiedBy>OpenAI</cp:lastModifiedBy>'
    f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
    f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
    '</cp:coreProperties>'
)

app_xml = xml_decl(
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
    '<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion>'
    '</Properties>'
)

parts = {
    '[Content_Types].xml': content_types_xml,
    '_rels/.rels': root_rels_xml,
    'word/document.xml': document_xml,
    'word/styles.xml': styles_xml,
    'word/settings.xml': settings_xml,
    'word/footnotes.xml': footnotes_xml,
    'word/_rels/document.xml.rels': document_rels_xml,
    'docProps/core.xml': core_xml,
    'docProps/app.xml': app_xml,
}

with zipfile.ZipFile(OUTPUT, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    for name, data in parts.items():
        zf.writestr(name, data.encode('utf-8'))

# Mechanical validation: ZIP integrity, required parts, XML well-formedness,
# exact body paragraph and footnote reference counts, and contiguous note IDs.
with zipfile.ZipFile(OUTPUT, 'r') as zf:
    assert zf.testzip() is None
    required = set(parts)
    assert required.issubset(set(zf.namelist()))
    for name in required:
        if name.endswith('.xml') or name.endswith('.rels'):
            ET.fromstring(zf.read(name))
    doc = zf.read('word/document.xml').decode('utf-8')
    notes = zf.read('word/footnotes.xml').decode('utf-8')
    assert doc.count('<w:p>') == 9  # one section heading + eight source paragraphs
    assert doc.count('<w:footnoteReference ') == 25
    for fid in range(26, 51):
        assert f'w:id="{fid}"' in notes
        assert doc.count(f'<w:footnoteReference w:id="{fid}"/>') == 1

print(f'Created: {OUTPUT}')
print(f'Bytes: {OUTPUT.stat().st_size}')
print('Paragraphs: 8 body paragraphs; footnotes: 26–50 (25 notes)')
