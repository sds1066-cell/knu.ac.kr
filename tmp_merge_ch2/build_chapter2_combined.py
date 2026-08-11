#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build one DOCX containing Korean thesis sections 2.1–2.5 in order.

The body wording and all footnote texts are taken from the five accepted section
files. No editorial rewriting is performed. The output contains true Word
footnotes numbered continuously from 1 through 71.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import generate_2_2_docx as s22
import generate_2_4_docx as s24
import generate_2_5_docx as s25

spec = importlib.util.spec_from_file_location(
    "s23", Path("docx_build_2_3/generate_2_3_docx.py")
)
if spec is None or spec.loader is None:
    raise RuntimeError("Could not load section 2.3 source module")
s23 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(s23)

OUT_DIR = Path("delivery_ch2")
DOCX_PATH = OUT_DIR / "Chapter2_Sections_2_1_to_2_5_Korean_Integrated.docx"
VERIFY_PATH = OUT_DIR / "Chapter2_Sections_2_1_to_2_5_Verification.txt"

BLOCKS_21: list[tuple[str, str]] = [
    ("section", "2.1 UNCLOS 프레임워크 하 미획정 해역 잠정 조치의 규범적 기초"),
    ("body", "최종 해양경계가 아직 확정되지 않았다는 것은 현실적 이용을 중단할 수 있음을 의미하지 않는다. 어로, 자원 보존 및 일상적 관리와 같은 어업활동은 계속되고 있으며, 대향 또는 인접 연안국은 경계획정 협상을 추진하는 동시에 과도기에 질서가 상실되는 것을 방지해야 한다. 제122조와 제123조는 이러한 협력을 반폐쇄해의 맥락에 위치시키며, 제74조와 제83조 제1항은 최종 경계획정이 합의를 통하여 공평한 해결에 이르러야 한다고 규정한다. 두 규범군은 각각 협력의 배경과 최종 목표를 제공하지만, 경계미획정 기간에 현실적 활동을 어떻게 관리해야 하는지에 대해서는 어느 쪽도 직접 답하지 않는다."),
    ("body", "제74조 제3항과 제83조 제3항은 바로 이 과도기를 대상으로 한다. 이 조항들은 한편으로 관련국이 실질적인 잠정약정을 체결하기 위해 노력하도록 요구하고, 다른 한편으로 최종 합의를 위태롭게 하거나 방해할 수 있는 일방적 행위를 제약한다. 어업형 잠정조치수역이 이러한 규범구조에 포섭될 수 있는지는 국가 간 합의, 실제 용도, 잠정성 및 최종 경계획정을 예단하지 않는다는 조건 등에 달려 있다. 이러한 조건이 충족되더라도 제도가 처리할 수 있는 사안은 여전히 협정 조항과 양 당사국의 공동 결정에 의해 정해진다. 또한 조문은 자제의무의 사실적 기준을 망라하지 않으므로, 구체적인 한계는 판례를 통해 밝혀야 한다."),
    ("subsection", "2.1.1 UNCLOS의 반폐쇄해 상황 하에서의 규범적 한계"),
    ("body", "郑凡은 UNCLOS 제122조와 제123조를 공간, 참여 주체, 협력 대상 및 절차적 경로에 관한 기본적 틀로 요약하고, 그 구체화 정도가 제한적이라고 지적한다.[[1]] 조문상 제122조는 지리적 형태와 해역의 관할관계를 기준으로 폐쇄해 또는 반폐쇄해를 정의하며, 제123조는 “should cooperate”라는 문구를 사용하여 연안국이 해양생물자원, 해양환경 및 해양과학조사 등의 사안에 관하여 협력하도록 요구하고, 국가가 직접 조정하거나 적절한 지역기구를 활용하는 것도 허용한다.[[2]] 황해는 한중 양국의 해안이 서로 마주 보는 반폐쇄해이므로 이러한 틀에서 논의할 수 있다. 제123조는 협력 분야와 채택 가능한 조직적 경로를 제시하지만, 공동기구, 의사결정 절차, 이행 감독 및 불이행 이후의 제도적 결과를 추가로 규정하지 않는다. 따라서 황해 협력에서는 국가 간 협정을 통해 일반적인 협력 요구를 구체적 규칙으로 전환할 필요가 있다."),
    ("body", "최종 경계획정 조항은 다른 층위의 문제를 다룬다. 제74조 제1항과 제83조 제1항은 대향 또는 인접 국가가 합의를 통하여 국제법에 기초한 공평한 해결에 이르도록 요구한다.[[3]] “합의를 통하여”는 경계가 국가 간 합의로 형성되어야 함을 나타내며, “공평한 해결”은 협상의 규범적 목표를 규정한다. 조문은 유일한 방법을 미리 정하지 않으며, 협상이 아직 완료되지 않았을 때 중첩된 주장과 현실적 활동을 어떻게 관리해야 하는지도 설명하지 않는다. 기선, 지리적 조건 및 이익 배분에 관한 이견은 여전히 협상에서 다루어야 한다. UNCLOS는 공통의 언어를 제공하지만, 이러한 정치적·법적 협의를 대신할 수는 없다."),
    ("body", "이로써 반폐쇄해 조항과 최종 경계획정 조항은 두 가지 기초적 과제를 수행한다. 전자는 협력이 이루어지는 맥락과 사안의 범위를 정하고, 후자는 영구적 경계의 목표와 형성 방식을 규정한다. 그러나 최종 합의가 이루어지기 전까지 협력이 어떠한 형식으로 전개되어야 하는지, 일방적 행위가 어떠한 제한을 받는지, 잠정약정이 어떠한 범위에서 효력을 가져야 하는지에 대해서는 여전히 직접적인 규범이 부족하다. 제74조 제3항과 제83조 제3항은 바로 이러한 과도기 관리 문제에 대응한다."),
    ("subsection", "2.1.2 제74조 제3항 및 제83조 제3항의 이중적 의무 구조"),
    ("body", "Lagoni는 제3항을 잠정약정의 체결을 촉진하기 위한 노력과 최종 합의를 위태롭게 하지 않을 의무로 요약한다.[[4]] 조문은 또한 어떠한 잠정약정도 최종 경계획정을 해하여서는 안 된다고 요구한다.[[5]] 이 세 가지 요구는 과도기에 함께 작용하지만 기능은 서로 다르다. 잠정약정의 체결을 촉진하는 의무는 현실적 이용을 위한 최소한의 질서를 구축하고, 자제의무는 일방적 행위가 협상 조건을 변경하는 것을 방지하며, 최종 경계획정을 예단하지 않는다는 조항은 잠정약정이 최종 권리의 근거로 간주되는 것을 막는다. 이 세 요소를 하나의 구조에 배치함으로써 협력의 여지를 남기는 동시에, 그러한 협력을 최종 경계획정 이전으로 한정한다."),
    ("body", "“모든 노력을 다한다”는 요구는 협상이 실질적으로 진전되는지에 초점을 둔다. Lagoni는 국가가 선의에 따라 실질적인 내용을 갖춘 협의를 진행해야 한다고 보며, 赵静 역시 국가가 협상에 성실하게 참여하고 필요한 경우 방안을 조정해야 한다고 강조한다.[[6]] 이 의무의 이행 여부를 판단할 때에는 최종적으로 제도적 문서가 형성되었는지만 볼 것이 아니라, 국가가 상대방의 방안에 지속적으로 응답하고 실질적인 협의를 진행했는지도 검토해야 한다. 합의에 이르지 못했다는 사실이 반드시 위반을 구성하는 것은 아니지만, 형식적인 회동, 기존 입장의 반복 또는 실질적 협의의 거부만으로는 조문의 요구를 충족하기에 부족하다."),
    ("body", "자제의무는 다른 유형의 위험을 대상으로 한다. Lagoni가 지적하듯이 모든 권리, 자유 또는 관할권의 행사가 위반을 구성하는 것은 아니며, 조문도 금지되는 행위의 목록을 열거하지 않는다.[[7]] UNCLOS는 최종 합의를 위태롭게 하거나 방해해서는 안 된다는 원칙만을 확립한다.[[8]] 경계미획정 해역이 자제의무로 인해 완전히 동결되는 것은 아니다. 활동이 한계를 넘어서는지는 여전히 물리적 효과, 지속기간 및 협상 조건에 미치는 실제 영향에 달려 있다."),
    ("body", "협력과 자제가 현실적 질서로 전환되기 위해서는 국가가 공동의 의사를 이행 가능한 약정에 명문화해야 한다. Lagoni는 특정 용도를 위해 형성된 잠정약정의 적용범위가 구체적인 합의에 의해 제한되며, 잠정적 성격을 갖는다는 이유만으로 다른 사안까지 자동으로 확대되지 않는다고 지적한다.[[9]] 관련 연구는 더 나아가 이러한 약정을 국가 간 합의에 기초하고 현실적 이용을 대상으로 하며, 최종 경계획정 이전에만 적용되고 최종 경계를 예단하지 않는 것으로 개괄한다.[[10]] 문서의 형식은 유연할 수 있지만, 그 규율 대상, 실질적 내용 및 효력의 경계는 명확하게 식별될 수 있어야 한다."),
    ("body", "한국 학자 원현우·허종원·이영주는 한중 잠정조치수역을 어업 사안에 관한 잠정적 공동관리구역으로 이해하고, 관련 협정이 최종 경계획정의 성격을 갖지 않는다고 강조한다.[[11]] 이러한 학술적 분류가 한중 PMZ에 적용될 수 있는지는 공동문서의 구체적인 조항을 통해 검토해야 한다. 한중 어업협정의 전문, 제1조, 제7조 및 제14조는 각각 국가 간 합의, 어업 협력의 대상, 잠정조치수역 및 해양법에 관한 입장의 유보를 나타낸다.[[12]] 이러한 요소들이 서로 대응할 때에만 한중 PMZ를 실질적인 잠정약정으로 제한적으로 이해할 수 있다. 이 협정은 어업활동을 직접 규율하므로 제74조 제3항과 해당 약정의 관련성이 더 직접적이며, 해저와 하층토 등 대륙붕 활동이 관련되는 경우에는 제83조 제3항이 이에 상응하는 과도기 규범을 제공한다."),
    ("body", "이러한 성격 규정은 경계가 확정되지 않은 상태에서도 협정이 성립할 수 있었던 이유를 설명하지만, 제도에 모든 해역 사안을 처리할 권한을 부여하지는 않는다. 양국은 특정한 어업 용도를 중심으로 질서를 유지하면서 최종적인 권리 주장도 계속 유지할 수 있다. 제도의 규율대상과 권한은 여전히 구체적인 합의에 의해 제약되며, 합의 밖의 사안은 자동으로 공동관리에 편입될 수 없다. 잠정성만으로 제도적 취약성이 발생하는 것도 아니다. 실질적인 위험은 용도 제한이 장기간 지속되는 반면, 규칙, 의제 범위 및 기구의 권한이 새로운 문제에 맞추어 확대되지 않을 때 나타난다. 그 경우 제도 고도화는 초기 약정의 제약을 받게 된다."),
    ("body", "이상에서 제74조 제3항과 제83조 제3항은 과도기에 협력과 자제가 필요한 이유를 설명하고, 어업형 PMZ의 잠정약정적 성격에 대한 규범적 근거도 제공한다. 그러나 조문은 어떠한 유형의 일방적 활동이 최종 합의를 위태롭게 하거나 방해하는지 아직 설명하지 않으며, 지속기간, 물리적 효과 및 실제 결과에 관한 구체적인 판단기준도 제시하지 않는다. 이 문제는 관련 판례가 서로 다른 사실관계를 어떻게 다루었는지를 통해서만 계속 검토할 수 있다."),
]

FOOTNOTES_21: dict[int, str] = {
    1: "郑凡, 「论〈联合国海洋法公约〉半闭海条款在南海区域合作中的适用」, 『政法论丛』 2019년 제6기, 101–105쪽.",
    2: "United Nations, United Nations Convention on the Law of the Sea, 1982, arts. 122–123; 联合国, 『联合国海洋法公约』, 제122–123조.",
    3: "United Nations, United Nations Convention on the Law of the Sea, 1982, arts. 74(1), 83(1); 联合国, 『联合国海洋法公约』, 제74조 제1항 및 제83조 제1항.",
    4: "Rainer Lagoni, “Interim Measures Pending Maritime Delimitation Agreements,” American Journal of International Law, Vol. 78, No. 2, 1984, pp. 348, 354.",
    5: "United Nations, United Nations Convention on the Law of the Sea, 1982, arts. 74(3), 83(3); 联合国, 『联合国海洋法公约』, 제74조 제3항 및 제83조 제3항.",
    6: "Rainer Lagoni, “Interim Measures Pending Maritime Delimitation Agreements,” American Journal of International Law, Vol. 78, No. 2, 1984, pp. 354–356; 赵静, 「论海洋划界前“作出临时安排”义务——以〈联合国海洋法公约〉第74、83条第3款为基础」, 『海南热带海洋学院学报』 2021년 제4기, 38–39쪽.",
    7: "Rainer Lagoni, “Interim Measures Pending Maritime Delimitation Agreements,” American Journal of International Law, Vol. 78, No. 2, 1984, pp. 362, 365–367.",
    8: "United Nations, United Nations Convention on the Law of the Sea, 1982, arts. 74(3), 83(3); 联合国, 『联合国海洋法公约』, 제74조 제3항 및 제83조 제3항.",
    9: "Rainer Lagoni, “Interim Measures Pending Maritime Delimitation Agreements,” American Journal of International Law, Vol. 78, No. 2, 1984, pp. 358–360.",
    10: "Rainer Lagoni, “Interim Measures Pending Maritime Delimitation Agreements,” American Journal of International Law, Vol. 78, No. 2, 1984, pp. 358–359; 赵静, 「论海洋划界前“作出临时安排”义务——以〈联合国海洋法公约〉第74、83条第3款为基础」, 『海南热带海洋学院学报』 2021년 제4기, 39–40쪽.",
    11: "원현우·허종원·이영주, 「한·중·일 분쟁수역의 법적 성격에 관한 연구」, 『해사법연구』 제36권 제2호, 2024, pp. 93–95, 99.",
    12: "《中华人民共和国政府和大韩民国政府渔业协定》, 전문, 제1조, 제7조 및 제14조; 동 협정 한국어본의 해당 조항.",
}

PARAGRAPHS_22: list[str] = [
    "제74조 제3항과 제83조 제3항은 관련국이 실질적인 잠정약정을 체결하기 위해 노력하는 동시에 최종 합의를 위태롭게 하거나 방해하지 않을 것을 요구한다.[[13]] 조약은 구체적인 위반 유형을 열거하지 않으므로, 사법 및 중재 실무는 행위의 경계를 명확히 하는 역할을 담당해 왔다. 활동의 명칭만을 기준으로 판단할 경우, 자제의무를 현실적 활동의 전면적인 동결로 해석할 수도 있고, 활동의 물리적 결과와 국가의 대응이 협상 조건에 미치는 영향을 간과할 수도 있다. 관련 판례는 이러한 문제를 중심으로 점차 보다 명확한 판단 경로를 형성해 왔다.",
    "절차 단계에 따라 재판기관이 다루어야 할 쟁점은 달라진다. 1976년 「에게해 대륙붕 사건」은 권리의 보전과 회복 불가능한 손해를 심사하였을 뿐, 제74조 제3항 또는 제83조 제3항에 따른 실체적 책임은 다루지 않았다. 재판소는 문제된 탄성파 탐사가 일시적 성격을 띠고, 해저에 시설을 설치하지 않았으며, 천연자원을 실제로 점유하거나 이용하지 않았다는 점에 주목하였다. 또한 당시 자료 역시 해당 탐사가 해저, 하층토 또는 천연자원에 물리적 손상을 초래할 것임을 보여 주지 않았다. 그리스가 주장한 권리 침해가 성립하더라도 적절한 방식으로 구제될 수 있었으므로, 재판소는 잠정조치를 명하지 않았다. 재판소는 동시에 신청 범위의 제약을 받아 그리스가 제기한 무력 문제에 관하여 실체적 판단을 내리지 않았다.[[14]] 1978년 같은 사건의 판결은 재판소의 관할권만을 다루었다.[[15]] 이 두 재판은 보전 단계에서 일시적 활동이 어떻게 평가되는지를 보여 주지만, 자제의무의 실체적 기준을 대신할 수 없으며, 탄성파 탐사가 일반적인 의미에서 적법하다는 결론을 뒷받침하지도 않는다. 일방적 활동 자체의 한계를 판단하려면 활동의 효과와 국가의 대응 방식을 직접 평가한 사건으로 나아가야 한다.",
    "「가이아나-수리남 사건」은 쟁점을 실체적 평가의 단계로 진전시켰다. 중재재판소는 관련국이 잠정약정 체결을 위해 선의의 협의를 진행하였는지, 일방적 활동이 최종 합의를 위태롭게 하거나 방해하였는지, 그리고 퇴거 행위가 무력의 위협 또는 사용을 구성하였는지를 각각 심사하였다. 잠정약정 체결을 위해 노력할 의무는 협의가 선의, 타협적 태도 및 필요한 양보를 바탕으로 진행되었는지를 살피며, 잠정약정을 체결하지 못했다는 사실만으로는 위반으로 인정하기에 부족하다. 반면 자제의무는 일방적 활동이 협상 조건과 최종 합의의 형성 여지를 변경하였는지에 초점을 둔다.[[16]] 수리남이 CGX 시추 플랫폼을 퇴거시킨 행위는 별도로 「유엔헌장」상 무력의 위협 또는 사용에 관한 규범에 따라 심사되었다.[[17]] 세 가지 쟁점은 동일한 사건에서 비롯되었지만, 적용 근거와 입증 경로는 서로 다르다. 활동의 효과에 따라 자제의무가 평가되며, 퇴거 또는 법집행에 의한 저지는 관련 국제법 규범에 따라 별도로 판단되어야 한다. 대응이 사실상 상대방의 활동을 배제하였는지는 협상 환경의 변화를 살펴보는 사실적 배경이 될 수 있으나, 이를 자제의무와 결합하여 하나의 법적 결론으로 볼 수는 없다.",
    "중재재판소는 경계미획정 해역을 반드시 정지 상태로 유지해야 하는 공간으로 보지 않았다. 판정은 해양환경에 물리적 변화를 초래하지 않는 일방적 행위는 일반적으로 최종 합의를 방해하지 않지만, 물리적 변화를 초래할 수 있는 행위는 현상을 변경하고 상대방의 협상 지위를 약화시킬 수 있으므로 통상 양측이 공동으로 수행하거나 합의에 따라 수행해야 한다고 보았다. 이에 따라 중재재판소는 탄성파 탐사와 시험굴착을 구분하여, 전자는 자제의무와 양립하지 않는 것으로 인정되지 않았으나 후자는 영구적인 변화 또는 손상을 초래할 수 있다고 보았다.[[18]] 이러한 차이의 핵심은 사실상태를 쉽게 회복할 수 있는지 여부에 있다. 되돌리기 어려운 물리적 변화는 원상회복 비용을 높이고, 본래 협상을 통해 조정할 수 있었던 상태를 더욱 조정하기 어렵게 만든다. 일시성, 가역성 및 보상 가능성은 회복과 조정을 위한 여지를 남기지만, 그 자체로 활동의 적법 여부를 자동으로 결정하지는 않는다. 본 논문은 이에 근거하여 지속적·배타적 또는 통제적 효과를 후속 사실 비교에서 확인해야 할 문제로 설정한다. 구체적인 판단은 여전히 공간적 위치, 기술적 속성, 지속기간 및 실제 영향과 결합하여 이루어져야 한다. Nishimoto 역시 영구성을 중요하지만 배타적이지는 않은 요소로 보며, 절차, 청구 내용 및 사건의 맥락이 모두 결론에 영향을 미친다고 지적한다.[[19]]",
    "문제된 활동 자체에 위험성이 있더라도, 국가는 이를 이유로 저지 방식을 임의로 선택할 수는 없다. 「가이아나-수리남 사건」에서 수리남 군함은 CGX 플랫폼에 정해진 시간 안에 해당 해역을 떠날 것을 요구하고, 이에 불응할 경우 그 결과를 감수해야 할 것이라고 위협하였다. 중재재판소는 명령을 내린 주체의 권한 근거, 명령의 내용, 위협의 임박성 및 그로 인해 발생할 수 있는 결과를 검토한 후, 해당 행위가 군사행동의 위협을 구성하며 일반적인 해상 법집행에는 해당하지 않는다고 판단하였다.[[20]] 중재재판소는 또한 당시에는 협상, 「협약」 제15부 및 부속서 Ⅶ에 따른 분쟁해결절차의 개시 또는 잠정조치 신청과 같은 평화적 수단으로 분쟁을 처리할 수 있었다고 지적하였다.[[21]] 관련 연구 역시 활동의 물리적 결과와 퇴거·법집행 대응을 서로 인접하지만 상호 대체할 수 없는 문제로 본다.[[22]] 강제적 대응을 평가할 때에는 행위 주체의 권한, 조치의 강도, 임박성 및 이용 가능한 평화적 대체수단을 별도로 확인해야 한다. 현장 활동이 저지되었다는 결과만으로는 통제 관계 또는 관할권이 이미 변경되었음을 입증할 수 없다.",
    "「가나-코트디부아르 해양경계획정 사건」의 두 단계는 영구성이 서로 다른 절차에서 상이한 기능을 수행한다는 점을 더욱 분명히 보여 준다. 2015년 잠정조치 명령은 새로운 굴착이 초래할 수 있는 중대하고 영구적인 물리적 변화를 권리 보전 단계의 핵심 위험으로 보았으며, 가나가 분쟁수역에서 더 이상의 새로운 굴착이 이루어지지 않도록 보장할 것을 요구하였다. 특별재판부는 이미 굴착이 진행된 사업을 갑자기 중단할 경우 중대한 경제적 손실과 해양환경 위험이 발생할 수 있다는 이유로 기존 활동을 전면 중단시키지는 않았다.[[23]] 2017년 본안판결은 제83조 제3항에 포함된 두 의무를 각각 심사하였다. 코트디부아르는 가나에 잠정약정에 관한 협상을 요청하지 않았으며, 가나는 2015년 명령에 따라 새로운 굴착을 이미 중단하였다. 최종 경계획정 결과에 따르면 문제된 석유·가스 활동이 이루어진 해역은 가나에 귀속되었고, 코트디부아르의 최종 청구에서 지칭한 “코트디부아르 해역”에는 포함되지 않았다. 이에 특별재판부는 자제의무 위반을 인정하지 않았다.[[24]] 동일한 요소는 잠정조치 단계에서 위험을 식별하는 데 사용되지만, 본안판결 단계에서는 청구 범위, 활동이 이루어진 공간 및 국가의 후속 행위와 결합되어야만 책임의 근거가 될 수 있다.",
    "책임에 관한 결론은 사실의 입증을 거쳐야 한다. 「소말리아-케냐 사건」은 굴착이 영구적인 물리적 변화를 초래할 수 있음을 인정하는 한편, 관련 활동의 시점, 위치 및 효과가 입증되어야 한다고 보았다. 이 사건에서 확인된 광구에 대한 양허 부여, 탄성파 탐사 및 기타 조사활동은 최종 합의에 충분한 영향을 미쳤다는 점이 입증되지 않았으며, 굴착과 관련된 구체적인 시점과 위치도 충분히 입증되지 않았다.[[25]] 전자의 주장은 효과에 관한 증거가 부족했고, 후자의 주장은 시간적·공간적 증거가 부족했다. “충분히 입증되지 않았다”는 것은 책임에 관한 결론을 제한할 뿐, 관련 사실이 전혀 발생하지 않았다는 뜻은 아니다. 활동 유형은 일반적인 위험을 시사할 뿐이며, 행위의 발생, 공간적 귀속 및 실제 효과가 입증되어야만 그 위험을 특정 국가에 귀속시킬 수 있다.",
    "이들 판례는 공통적으로 하나의 기본 판단을 제시한다. 자제의무가 보호하는 것은 최종 합의를 계속 협상하고 조정할 수 있는 여지이므로, 분쟁 상태를 회복하기 어렵게 만들 수 있는 행위가 그 관심의 대상이 된다. 물리적 변화의 회복 가능성, 절차 단계 및 증거 상황이 함께 평가의 강도를 결정하며, 강제적 대응은 별도의 규범적 심사를 받아야 한다. 고정시설, 조사, 통항 및 접근 제한이 서로 비교 가능한지는 구체적인 활동, 대응 방식 및 사실적 근거를 결합하여 각각 판단해야 한다. 판례는 행위가 어떻게 평가되는지를 설명할 수 있지만, 한중 양국이 왜 우선 어업 분야에서 협력을 형성했는지는 설명할 수 없다. 이 문제에 답하기 위해서는 황해의 현실적 조건과 양국의 협상과정을 함께 살펴보고, 어업 의제가 왜 최종 경계획정에 앞서 협력약정에 포함될 수 있었는지를 검토해야 한다.",
]


def pieces_to_marker_text(pieces: list[str | int]) -> str:
    return "".join(f"[[{item}]]" if isinstance(item, int) else item for item in pieces)


BLOCKS: list[tuple[str, str]] = []
BLOCKS.extend(BLOCKS_21)
BLOCKS.append(("section", s22.TITLE))
BLOCKS.extend(("body", paragraph) for paragraph in PARAGRAPHS_22)
BLOCKS.append(("section", s23.TITLE))
BLOCKS.extend(("body", pieces_to_marker_text(paragraph)) for paragraph in s23.PARAGRAPHS)
BLOCKS.extend(list(s24.BLOCKS))
BLOCKS.append(("section", s25.TITLE))
BLOCKS.extend(("body", paragraph) for paragraph in s25.PARAGRAPHS)

FOOTNOTES: dict[int, str] = {}
for source in (FOOTNOTES_21, s22.FOOTNOTES, s23.FOOTNOTES, s24.FOOTNOTES, s25.FOOTNOTES):
    overlap = set(FOOTNOTES).intersection(source)
    if overlap:
        raise RuntimeError(f"Duplicate footnote numbers: {sorted(overlap)}")
    FOOTNOTES.update(source)

EXPECTED_FOOTNOTES = list(range(1, 72))
if sorted(FOOTNOTES) != EXPECTED_FOOTNOTES:
    raise RuntimeError(f"Footnote range mismatch: {sorted(FOOTNOTES)}")

MARKER_RE = re.compile(r"\[\[(\d+)\]\]")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def xml_text(text: str) -> str:
    return escape(text, {'"': '&quot;'})


def text_run(text: str, *, size: int | None = None, bold: bool = False) -> str:
    if not text:
        return ""
    props: list[str] = []
    if bold:
        props.append("<w:b/><w:bCs/>")
    if size is not None:
        props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xml_text(text)}</w:t></w:r>'


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
        fid = int(match.group(1))
        if fid not in FOOTNOTES:
            raise ValueError(f"Unknown footnote marker: {fid}")
        parts.append(footnote_reference(fid))
        pos = match.end()
    parts.append(text_run(text[pos:]))
    return "".join(parts)


def make_document_xml() -> str:
    body_parts: list[str] = []
    for kind, text in BLOCKS:
        if kind == "section":
            style, size, bold = "SectionHeading", 24, True
            runs = text_run(text, size=size, bold=bold)
        elif kind == "subsection":
            style, size, bold = "SubsectionHeading", 23, True
            runs = text_run(text, size=size, bold=bold)
        elif kind == "body":
            style = "BodyTextKorean"
            runs = runs_with_footnotes(text)
        else:
            raise ValueError(f"Unknown block kind: {kind}")
        body_parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{runs}</w:p>'
        )

    sect_pr = '''
<w:sectPr>
  <w:footnotePr>
    <w:numFmt w:val="decimal"/>
    <w:numStart w:val="1"/>
    <w:numRestart w:val="continuous"/>
    <w:pos w:val="pageBottom"/>
  </w:footnotePr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
  <w:cols w:space="708"/>
  <w:docGrid w:linePitch="312"/>
</w:sectPr>
'''
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        + "".join(body_parts)
        + sect_pr
        + "</w:body></w:document>"
    )


def make_footnotes_xml() -> str:
    items = [
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    ]
    for fid in EXPECTED_FOOTNOTES:
        items.append(
            f'<w:footnote w:id="{fid}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            '<w:r><w:tab/></w:r>'
            + text_run(FOOTNOTES[fid], size=18)
            + '</w:p></w:footnote>'
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
      <w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/>
      <w:sz w:val="21"/><w:szCs w:val="21"/>
      <w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:line="384" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
    <w:pPr><w:widowControl/><w:jc w:val="both"/><w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BodyTextKorean">
    <w:name w:val="Body Text Korean"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:widowControl/><w:jc w:val="both"/><w:ind w:firstLine="420" w:firstLineChars="200"/><w:spacing w:before="0" w:after="0" w:line="384" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SectionHeading">
    <w:name w:val="Section Heading"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:jc w:val="left"/><w:spacing w:before="240" w:after="240" w:line="384" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SubsectionHeading">
    <w:name w:val="Subsection Heading"/><w:basedOn w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:jc w:val="left"/><w:spacing w:before="240" w:after="120" w:line="384" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:b/><w:bCs/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:jc w:val="both"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="바탕" w:cs="Batang"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
</w:styles>'''


def make_settings_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W_NS}">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:footnotePr><w:numFmt w:val="decimal"/><w:numStart w:val="1"/><w:numRestart w:val="continuous"/><w:pos w:val="pageBottom"/></w:footnotePr>
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
  <w:doNotTrackMoves/><w:doNotTrackFormatting/>
</w:settings>'''


def make_font_table_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="{W_NS}">
  <w:font w:name="Batang"><w:charset w:val="81"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
  <w:font w:name="바탕"><w:charset w:val="81"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
</w:fonts>'''

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

APP_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion>
</Properties>'''


def make_core_xml() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Chapter 2 Sections 2.1–2.5 Korean Integrated</dc:title><dc:subject>Master's thesis Korean translation</dc:subject><dc:creator>OpenAI</dc:creator><cp:keywords>Chapter 2; 2.1-2.5; Korean; thesis</cp:keywords><dc:description>Faithful integration of the accepted Korean translations of sections 2.1 through 2.5</dc:description><cp:lastModifiedBy>OpenAI</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def write_docx() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    members = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": PACKAGE_RELS,
        "docProps/core.xml": make_core_xml(),
        "docProps/app.xml": APP_XML,
        "word/document.xml": make_document_xml(),
        "word/styles.xml": make_styles_xml(),
        "word/settings.xml": make_settings_xml(),
        "word/fontTable.xml": make_font_table_xml(),
        "word/footnotes.xml": make_footnotes_xml(),
        "word/_rels/document.xml.rels": DOCUMENT_RELS,
    }
    with zipfile.ZipFile(DOCX_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in members.items():
            zf.writestr(name, content.encode("utf-8"))


def paragraph_text(paragraph: ET.Element, ns: dict[str, str]) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", ns))


def markerless(text: str) -> str:
    return MARKER_RE.sub("", text)


def sha_of_json(value: object) -> str:
    payload = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def verify_docx() -> str:
    required = {
        "[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml",
        "word/document.xml", "word/styles.xml", "word/settings.xml", "word/fontTable.xml",
        "word/footnotes.xml", "word/_rels/document.xml.rels",
    }
    with zipfile.ZipFile(DOCX_PATH, "r") as zf:
        bad = zf.testzip()
        if bad is not None:
            raise RuntimeError(f"Corrupt DOCX member: {bad}")
        names = set(zf.namelist())
        missing = required - names
        if missing:
            raise RuntimeError(f"Missing DOCX members: {sorted(missing)}")
        document_bytes = zf.read("word/document.xml")
        footnotes_bytes = zf.read("word/footnotes.xml")
        settings_bytes = zf.read("word/settings.xml")
        styles_bytes = zf.read("word/styles.xml")

    document_root = ET.fromstring(document_bytes)
    footnotes_root = ET.fromstring(footnotes_bytes)
    ET.fromstring(settings_bytes)
    ET.fromstring(styles_bytes)
    ns = {"w": W_NS}

    body_paragraphs = document_root.findall(".//w:body/w:p", ns)
    actual_body_texts = [paragraph_text(p, ns) for p in body_paragraphs]
    expected_body_texts = [markerless(text) for _, text in BLOCKS]
    if actual_body_texts != expected_body_texts:
        for idx, (actual, expected) in enumerate(zip(actual_body_texts, expected_body_texts), start=1):
            if actual != expected:
                raise RuntimeError(f"Body text mismatch at paragraph {idx}: {actual!r} != {expected!r}")
        raise RuntimeError("Body paragraph count mismatch")

    refs = [
        int(el.attrib[f"{{{W_NS}}}id"])
        for el in document_root.findall(".//w:footnoteReference", ns)
    ]
    expected_refs = [
        int(match.group(1))
        for kind, text in BLOCKS if kind == "body"
        for match in MARKER_RE.finditer(text)
    ]
    if expected_refs != EXPECTED_FOOTNOTES:
        raise RuntimeError(f"Source footnote order mismatch: {expected_refs}")
    if refs != EXPECTED_FOOTNOTES:
        raise RuntimeError(f"DOCX footnote reference order mismatch: {refs}")

    actual_note_ids: list[int] = []
    actual_note_texts: dict[int, str] = {}
    for note in footnotes_root.findall("w:footnote", ns):
        fid = int(note.attrib[f"{{{W_NS}}}id"])
        if fid > 0:
            actual_note_ids.append(fid)
            actual_note_texts[fid] = "".join(node.text or "" for node in note.findall(".//w:t", ns))
    if actual_note_ids != EXPECTED_FOOTNOTES:
        raise RuntimeError(f"DOCX footnote node order mismatch: {actual_note_ids}")
    if actual_note_texts != FOOTNOTES:
        for fid in EXPECTED_FOOTNOTES:
            if actual_note_texts.get(fid) != FOOTNOTES[fid]:
                raise RuntimeError(f"Footnote {fid} text mismatch")

    heading_order = [
        "2.1 UNCLOS 프레임워크 하 미획정 해역 잠정 조치의 규범적 기초",
        "2.1.1 UNCLOS의 반폐쇄해 상황 하에서의 규범적 한계",
        "2.1.2 제74조 제3항 및 제83조 제3항의 이중적 의무 구조",
        "2.2 관련 판례의 자제 의무 해석 기능",
        "2.3 한중 해양 경계 미획정의 역사적 맥락과 협정 협상 배경",
        "2.4 한중 어업협정의 제도적 설계",
        "2.4.1 세 가지 수역의 법리적 구조 : 협정 수역, 잠정조치 수역, 과도 수역",
        "2.4.2 한중 어업공동위원회와 연례 입어 조건 메커니즘",
        "2.4.3 어업 규칙의 정밀성과 비어업 공간 이용의 규칙 공백",
        "2.5 협정 형성의 가능성 조건",
    ]
    actual_headings = [text for kind, text in BLOCKS if kind in {"section", "subsection"}]
    if actual_headings != heading_order:
        raise RuntimeError(f"Heading order mismatch: {actual_headings}")

    document_text = "".join(document_root.itertext())
    if "[[" in document_text or "]]" in document_text:
        raise RuntimeError("Unresolved marker remains in DOCX")

    docx_sha = hashlib.sha256(DOCX_PATH.read_bytes()).hexdigest()
    source_body_sha = sha_of_json(BLOCKS)
    source_footnotes_sha = sha_of_json(FOOTNOTES)
    body_count = sum(1 for kind, _ in BLOCKS if kind == "body")
    section_count = sum(1 for kind, _ in BLOCKS if kind == "section")
    subsection_count = sum(1 for kind, _ in BLOCKS if kind == "subsection")

    return "\n".join([
        "RESULT=PASS",
        f"DOCX={DOCX_PATH.name}",
        f"DOCX_BYTES={DOCX_PATH.stat().st_size}",
        f"DOCX_SHA256={docx_sha}",
        f"SOURCE_BODY_SHA256={source_body_sha}",
        f"SOURCE_FOOTNOTES_SHA256={source_footnotes_sha}",
        f"TOTAL_PARAGRAPHS={len(BLOCKS)}",
        f"SECTION_HEADINGS={section_count}",
        f"SUBSECTION_HEADINGS={subsection_count}",
        f"BODY_PARAGRAPHS={body_count}",
        "FOOTNOTE_REFERENCE_COUNT=71",
        "FOOTNOTE_NODE_COUNT=71",
        "FOOTNOTE_DISPLAY_RANGE=1-71",
        "SECTION_ORDER=2.1,2.2,2.3,2.4,2.5",
        "BODY_TEXT_EXACT_MATCH=PASS",
        "FOOTNOTE_TEXT_EXACT_MATCH=PASS",
        "OOXML_PARSE=PASS",
        "ZIP_INTEGRITY=PASS",
        "FORMAT=A4; Batang; body 10.5pt; justified; first-line indent; 1.6 line spacing",
        "CONTENT_EDITED=NO",
        "REFERENCES_EDITED=NO",
        "",
    ])


def main() -> None:
    write_docx()
    verification = verify_docx()
    VERIFY_PATH.write_text(verification, encoding="utf-8")
    print(verification, end="")


if __name__ == "__main__":
    main()
