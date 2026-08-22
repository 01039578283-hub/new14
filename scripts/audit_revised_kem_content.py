from __future__ import annotations

"""Strict read-only content and provenance audit for revised K/E/M detail pages.

The generator is deliberately imported only to obtain an in-memory projection.
This program never calls its apply/freeze entry points and never writes inside
the repository.  The checked release contract is 1,114 product documents
(1,113 detail pages plus sitemap.xml) and the three named release scripts.
"""

import argparse
import copy
import hashlib
import html
import importlib.util
import json
import re
import statistics
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
from typing import Iterable
from urllib.parse import unquote


ROOT = Path(__file__).resolve().parents[1]
SUBJECT_ROOT = ROOT / "과목별학원"
DEFAULT_GENERATOR = ROOT / "scripts" / "generate_revised_kem_pages.py"
DEFAULT_ARCHIVE = Path.home() / "Desktop" / "1.zip"
DEFAULT_COMMON = ROOT.parent / "참고자료" / "공통자료"
SITE_ORIGIN = "https://xn--3e0bz50b1zcyxat54c.com"
REVISION_DATE = "2026-08-22"
REVISION_MARKER = "composite-2026-08-22"
EXPECTED_BASE_COMMIT = "9e58f271f6126db72d4eb10a363c9d3b4d163779"
EXPECTED_GENERATOR_SHA256 = "f145adec84c78a61bfcf30b8a137e5d741e0eb1f1ba59a4370fc2409533b296e"
EXPECTED_CANDIDATE_SHA256 = "33f212961e34e6978d9dfa5b0eeab9cad7916a0da6626afccadcd67a5f17b9e6"
EXPECTED_GENERATOR_AFTER_MANIFEST_SHA256 = "081de645e104568f2b63e019093c5656ba585dc4204947b92098d938ee240cc9"
KNOWN_SUPERSEDED_AFTER_MANIFEST_SHA256 = "f8828e61788be26d9d72e5ed619b450d626f5f0964052686730c4b6f50c8f451"
EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS = 290
EXPECTED_DETAILS_PER_PROFILE = 371
EXPECTED_DETAIL_DOCUMENTS = 1113
EXPECTED_PRODUCT_DOCUMENTS = 1114
EXPECTED_RELEASE_SCOPE = 1117
PROFILE_SPECS = (
    ("elementary", "초등학생학원 원고.xlsx", "초등학생국영수학원"),
    ("middle", "중학생학원 원고.xlsx", "중학생국영수학원"),
    ("high", "고등학생학원 원고.xlsx", "고등학생국영수학원"),
)
RELEASE_SCRIPTS = (
    ROOT / "scripts" / "generate_revised_kem_pages.py",
    ROOT / "scripts" / "audit_revised_kem_content.py",
    ROOT / "scripts" / "audit_revised_kem_technical.py",
)
REPRESENTATIVE_TARGETS = (
    ("elementary", "가경동"),
    ("elementary", "명일동"),
    ("elementary", "개운동"),
    ("middle", "명일동"),
    ("middle", "송강동"),
    ("middle", "좌동"),
    ("middle", "개운동"),
    ("high", "명일동"),
    ("high", "첨단"),
    ("high", "연동"),
)
LOCALITY_ALIASES = {
    "신불당": {"불당동"},
    "운정": {"야당동"},
    "운정신도시": {"야당동"},
    "별내동": {"별내중앙"},
}

KNOWN_AWKWARD_COPY = (
    "고이 포함됩니다", "고등학교이 포함됩니다", "기록라는 표현",
    "에 대한 설명는", "중학교이 포함됩니다", "초이 포함됩니다",
    "초등학교이 포함됩니다", "관리을", "관리이라는",
    "학부모님이 체감하는 현실은 학부모님이", "관리 기준 방문 전에는",
    "요일를", "시간를", "범위을", "이해도을", "결과을", "진도을",
    "분류을", "학원로", "와와학교 일정 점검학원", "_x000D_",
    "영약", "유형 학습- 문제풀이", "교과 평가과",
    "문장력 향상으로 연결합니다", "유형 전환 능력을 오답으로 관리",
    "문제 해결력 기반으로 보완", "학습 변화’으로",
    "시간은 했는데 학습 변화를 확인하는 상황", "유형 적응을 빠르게 합니다",
    "국어·학습 코칭", "영수(영수)",
)
AWKWARD_PATTERNS = {
    "duplicate_level_fit": re.compile(r"목표에\s*맞(?:춘|는)\s*현재\s*수준에\s*맞춘"),
    "provided_internal_exam": re.compile(r"(?:학교\s+)?내신\s+제공된\s+시험지의\s+문항\s+유형"),
    "provided_exam": re.compile(r"(?<!내신\s)시험\s+제공된\s+시험지의\s+문항\s+유형"),
    "subject_column_duplicate": re.compile(r"(?:국어(?:와|/)\s*국어·영어·수학|영어/수학/국어/영어·수학)"),
    "subject_particle": re.compile(r"(?:영어·수학|국어·영어·수학|수학)\s*를(?![가-힣])"),
    "object_rewrite_collision": re.compile(
        r"(?:을|를)\s+(?:점검\s+순서를\s+정리|보완\s+순서를\s+정|설계\s+기준을\s+정리|단계를\s+나눠)"
    ),
    "elementary_token_corruption": re.compile(r"별교과\s*평가도시"),
    "missing_outcome_object": re.compile(
        r"학습\s+과정에서는\s+[^.!?]{1,60}(?:을|를)\s+바탕으로\s+"
        r"(?:다음\s+보완\s+순서에\s+반영|복습\s+시점을\s+정하는\s+데\s+활용|"
        r"변화를\s+확인하는\s+기준으로\s+삼|다음\s+학습\s+계획에\s+기록|"
        r"반복\s+오류를\s+점검하는\s+데\s+활용|설명\s+가능한\s+범위를\s+확인하는\s+데\s+씀|"
        r"보완할\s+단원을\s+구분하는\s+데\s+활용|상담에서\s+확인할\s+기준으로\s+정리)"
    ),
    "double_topic_consultation": re.compile(
        r"(?:실제[^.!?]{0,60}때는|상담\s+기록에서는\s+이와\s+함께|"
        r"과목별[^.!?]{0,60}대조하면)[^.!?]{0,80}상담은"
    ),
    "stacked_lead_in": re.compile(
        r"(?:과목별\s+기록을\s+대조하면|실제\s+학습\s+흐름을\s+확인할\s+때는|"
        r"기존\s+점검\s+기준과\s+함께\s+보면),?\s*최근\s+교재와\s+오답을\s+대조하면,"
    ),
    "duplicate_current_level_fit": re.compile(
        r"현재\s+수준에\s+맞춘[^.!?]{0,60}현재\s+수준에\s+맞춘"
    ),
    "duplicate_dot_token": re.compile(r"(?<![가-힣])([가-힣]{2,12})·\1(?![가-힣])"),
    "duplicate_compound_suffix": re.compile(
        r"(?:영수·영수학습|서술·서술형(?:/유형별)?|독해·독해형)"
    ),
    "outcome_service_fragment": re.compile(
        r"(?:성적|점수|실력|성취|성과|결과|성장|독해력|문해력|사고력|응용력|"
        r"해석력|이해력|대응력|적응력|풀이력|역량|기본기|실전\s+감각)"
        r"[^.!?]{0,45}?(?:완성|연결|강화|확장|상승|향상|성장|코칭|관리|지도|설계|대비)\."
        r"(?=\s|$)"
    ),
}
READER_META_TERMS = (
    "검색엔진", "상위노출", "SEO", "LOCAL ACADEMY GUIDE",
    "지역명을 바꾼 홍보 문구", "정보성 페이지 형태", "검색어",
)
FORBIDDEN_SOURCE_MARKUP = (
    "article-section", "feature-grid", "article-closing", "<h1", "_x000D_",
)
PROMOTIONAL_PATTERNS = {
    "guarantee": re.compile(r"(?:성적|점수|합격).{0,16}(?:보장|약속)"),
    "certain_improvement": re.compile(r"(?:성적|점수|실력).{0,14}(?:반드시|확실히).{0,10}(?:오르|향상|상승)"),
    "raise_score": re.compile(r"(?:점수|성적|실력).{0,16}(?:끌어올|올려\s*드립니다|올립니다)"),
    "superlative": re.compile(r"(?:최고의?|최적의|유일한|완벽한|업계\s*1위|지역\s*1위)"),
    "absolute": re.compile(r"(?:무조건|100\s*%|반드시\s*성공)"),
    "score_outcome": re.compile(r"점수(?:를|가)?[^.!?]{0,24}(?:확보합니다|확보하도록|만듭니다)"),
    "score_conversion": re.compile(
        r"(?:오답|약점|실수|오류|학습\s*결과)(?:이|가|을|를)[^.!?]{0,45}?"
        r"[‘“\"']?(?:다음\s+)?(?:시험\s+)?점수[’”\"']?\s*로[^.!?]{0,10}?"
        r"(?:바꿉니다|바꾸(?:는|도록|게|어|며|고)|바꿔(?:집니다|지도록|주는|줍니다)?|"
        r"바뀝니다|바뀌(?:는|도록|게|어|며|고)|바뀐)"
        r"(?:\s+(?:방법|기준|과정|전략|질문|표현|피드백|관리))?"
    ),
    "outcome_transform": re.compile(
        r"(?:꾸준함|오답|약점|성적)(?:이|가|을|를|은|는)?[^.!?]{0,35}?"
        r"(?:성적|실력|점수)(?:이|가|을|를|은|는|으?로)?[^.!?]{0,12}?"
        r"(?:바꿉니다|바꾸(?:는|도록|게|어|며|고)|바뀝니다|바뀌(?:는|도록|게|어)|"
        r"전환(?:합니다|됩니다|시키|하도록|되도록|하는))"
    ),
    "optimized": re.compile(r"최적화된"),
    "outcome_link": re.compile(
        r"(?:성적|점수|실력|정확도|성취도|성취|성과|결과|성장|학습\s+변화|점수\s+변화)"
        r"(?:이|가|을|를|은|는|도|으?로|까지)?\s*(?:바로\s+|곧\s+|다음\s+)?"
        r"(?:이어지|연결)"
    ),
    "outcome_transition": re.compile(
        r"(?:성적|점수|실력|정확도|성취도|성취|성과|결과|학습\s+결과|학습\s+변화|점수\s+변화)"
        r"[^.!?]{0,30}?(?:향상|상승|성장)[^.!?]{0,16}?(?:이어지|연결)"
        r"(?:도록|되도록|됩니다|합니다|하는)"
    ),
    "outcome_growth": re.compile(
        r"(?:성장을\s*(?:돕|지원|이끌|관리|지도|코칭|목표|약속)|"
        r"성장(?:하도록|할\s+수\s+있도록|하게|되게|합니다|됩니다|시킵니다|시켜|"
        r"하는\s+방향|으로\s+연결|이\s+이어지)|"
        r"성장\s+로드맵[^.!?]{0,30}(?:제공|운영|설계))"
    ),
    "outcome_make": re.compile(
        r"(?:성적|점수|실력|정확도|성취도|성취|성과|결과)"
        r"(?:이|가|을|를|은|는|도|으?로)?\s*"
        r"(?:바로\s+|다음\s+|꾸준히\s+|안정적인\s+|안정적으로\s+|결국\s+|함께\s+)?"
        r"(?:만들|만듭|완성|확보|달성|나오|나타나)"
    ),
    "outcome_raise": re.compile(
        r"(?:성적|점수|실력|정확도|성취도|성취|성과|결과|효과|수업(?:의)?\s+효과|학습(?:의)?\s+효과)"
        r"(?:이|가|을|를|은|는|도|의|과|와|으?로)?[^.!?]{0,35}?"
        r"(?:향상|상승|오르|올라|올리|올립|올려|끌어올|높이|높입|높여|키우|키웁|"
        r"개선|확장|강화|극대화)(?:합니다|됩니다|시킵니다|시키는|시키며|하도록|"
        r"되도록|되게|할\s+수|해|하는|되는)"
    ),
    "outcome_direct_link": re.compile(
        r"(?:성적|점수|성과|결과|실력)[^.!?]{0,30}?직결"
    ),
    "outcome_stack": re.compile(
        r"(?:성적|점수|실력|정확도|성취도|성취|성과|결과|성장|학습\s+변화|점수\s+변화)"
        r"(?:이|가|을|를|은|는|도|의|과|와|으?로|까지)?[^.!?]{0,20}?"
        r"(?:쌓|누적)(?:습니다|합니다|됩니다|시킵니다|시키는|시키며|되도록|되게|"
        r"되는|하도록|하게|을\s+수|아가|이게|이도록)"
    ),
    "outcome_stable": re.compile(
        r"(?:성적|점수|실력|성과|성취도|학습\s+변화)"
        r"(?:이|가|을|를|은|는|도|의|과|와|으?로|까지)?[^.!?]{0,30}?"
        r"(?:안정(?:됩니다|되도록|되게|시키|화)|유지(?:됩니다|되도록|하도록|하게|합니다))"
    ),
    "outcome_decline": re.compile(
        r"(?:성적|점수|실력)[^.!?]{0,35}?(?:하락(?:을)?\s*(?:방지|막)|"
        r"떨어지지\s*않|낮아지지\s*않|유지(?:되도록|하게|합니다|됩니다))"
    ),
    "outcome_achievement": re.compile(
        r"(?:성취도|성취)(?:이|가|은|는|을|를|과|와|도|의)?[^.!?]{0,60}?"
        r"(?:향상|상승|높입|높이|올립|올리|끌어올|개선|확보|강화|확장|성장|쌓|누적|"
        r"만듭|달성)"
    ),
    "outcome_ability": re.compile(
        r"(?:독해력|문해력|사고력|응용력|문제\s*해결력|계산력|적용력|학습력|집중력|"
        r"어휘력|표현력|평가\s*대응력|실전\s*감각|기초\s*체력|자신감|이해력|해석력|"
        r"판단력|서술력|논리력|추론력|풀이력|기본기|기초력|학습\s*역량|과목\s*역량|"
        r"국어\s*역량|영어\s*역량|수학\s*역량|학업\s*역량|평가\s*역량|대응력|적응력|풀이\s*역량|"
        r"서술\s*역량|읽기\s*역량|해석\s*능력|풀이\s*능력|학습\s*능력)"
        r"(?:이|가|을|를|은|는|도|의|과|와)?[^.!?]{0,25}?"
        r"(?:향상|상승|오르|올라|올리|올립|올려|끌어올|높이|높입|높여|키우|키웁|"
        r"개선|확보|만듭|쌓|누적|성장|안정|확장|강화|완성|극대화)"
    ),
    "outcome_appearance": re.compile(
        r"(?:성적|점수|실력|성과|결과|학습\s+변화)(?:이|가|에|은|는|도)?[^.!?]{0,25}?"
        r"(?:나게|나도록|나오게|나오도록|따라옵니다|따라오(?:게|도록|는|며|고)|"
        r"반영되(?:게|도록|는|며|고|었습니다|ㅂ니다)|"
        r"회복(?:합니다|됩니다|시키|하게|하도록|되도록|하는)|"
        r"도달(?:합니다|하게|하도록|하는))"
    ),
    "outcome_destination_link": re.compile(
        r"(?:(?:학습\s+변화|성적|점수\s+변화)(?:에|으?로)[^.!?]{0,8}?(?:연결|이어지)|"
        r"(?:오답|약점|실수|오류)[^.!?]{0,35}?[‘“\"']?(?:다음\s+)?(?:시험\s+)?"
        r"(?:점수|성과)[’”\"']?(?:에|으?로)[^.!?]{0,8}?(?:연결|이어지))"
    ),
    "outcome_minimize": re.compile(
        r"(?:성적|점수|실력)(?:의|에서)?\s*(?:격차|차이|편차|변동|변화|하락|손실|흔들림)"
        r"(?:을|를|은|는)?[^.!?]{0,8}?최소화(?:합니다|됩니다|시키|하도록|되도록|하는)"
    ),
    "outcome_ability_finite": re.compile(
        r"(?:스스로\s+공부하는\s+힘|독해력|문해력|사고력|응용력|문제\s*해결력|계산력|"
        r"적용력|학습력|집중력|어휘력|표현력|평가\s*대응력|실전\s*감각|기초\s*체력|"
        r"자신감|이해력|해석력|판단력|서술력|논리력|추론력|풀이력|기본기|기초력|"
        r"학습\s*역량|과목\s*역량|국어\s*역량|영어\s*역량|수학\s*역량|학업\s*역량|"
        r"평가\s*역량|대응력|적응력|풀이\s*역량|서술\s*역량|읽기\s*역량|해석\s*능력|"
        r"풀이\s*능력|학습\s*능력)(?:이|가|을|를|은|는|도|과|와)?[^.!?]{0,18}?"
        r"(?:길러(?:냅니다|줍니다|지도록|가는|가도록)|기릅니다|키웁니다|키우(?:도록|게|는|며|고)|"
        r"회복(?:합니다|됩니다|시키|하도록|되도록|하는)|보완(?:합니다|되도록|하는)|"
        r"다집니다|다지(?:도록|게|는|며|고)|잡습니다|잡아(?:줍니다|가도록|주는)|잡도록)"
    ),
    "outcome_growth_flow_make": re.compile(
        r"(?:성장|학습\s+변화)[^.!?]{0,16}?(?:루틴|흐름)(?:이|가|을|를|은|는)?"
        r"[^.!?]{0,8}?(?:만듭니다|만들(?:어|도록|게|며|고|는))"
    ),
    "outcome_change_make": re.compile(
        r"(?:성적|점수|실력|성과|결과|학습\s+변화|점수\s+변화)"
        r"(?:이|가|을|를|은|는|도|으?로)?[^.!?]{0,20}?"
        r"(?:만듭니다|만들(?:어|도록|게|며|고|는))"
    ),
    "outcome_skill_build": re.compile(
        r"실력(?:이|가|을|를|은|는|도)?[^.!?]{0,14}?"
        r"(?:잡습니다|잡아(?:줍니다|가도록|주는)|잡도록|다집니다|다지(?:도록|게|는|며|고)|"
        r"쌓습니다|쌓이(?:도록|게|는)|쌓(?:도록|게|는))"
    ),
    "outcome_achievement_maintain": re.compile(
        r"(?:성취도|성취감|성취\s+경험|성취)(?:이|가|을|를|은|는|도)?[^.!?]{0,14}?"
        r"유지(?:합니다|됩니다|시키|하도록|되도록|하는)"
    ),
    "outcome_variance_reduce": re.compile(
        r"(?:격차|변동)(?:이|가|을|를|은|는|도)?[^.!?]{0,14}?"
        r"(?:줄입니다|줄이(?:도록|게|는|며|고)|줄어(?:들도록|듭니다|드는|들게))"
    ),
    "academic_omission_control": re.compile(
        r"(?:(?:개념|기초|학습|진도|내용)\s*누락[^.!?]{0,35}?"
        r"(?:최소화(?:합니다|됩니다|시키|하도록|되도록|하는)|"
        r"줄이(?:고|며|는|도록|게)|줄여|막(?:습니다|도록|는)|"
        r"방지(?:합니다|하도록|하는)|없애(?:도록|는|줍니다))|"
        r"(?:개념\s*누락|풀이\s*실수|독해\s*오류)[^.!?]{0,40}?재발[^.!?]{0,20}?"
        r"(?:막(?:습니다|도록|는)|정착(?:시킵니다|시키는|하도록)))"
    ),
}
OUTCOME_SAFE_NOMINAL_SPANS = re.compile(
    r"(?:목표\s+성적\s+달성을\s+위한|"
    r"(?:향상|상승|개선|강화|확장)(?:을|를|에|은|이|과|와|의|까지)?\s*"
    r"(?:위해|위한|목표로|목표|계획|전략|방법|기준|과정|필요|중요|로드맵|방향|피드백)|"
    r"훈련\s+강화|성장\s+(?:방향|목표)|성장을\s+목표로|성장하는\s+방향|"
    r"(?:지속\s+)?성장\s+로드맵|성취\s+목표(?:\([^)]*\))?|성취\s+중심|"
    r"쌓아가고\s+싶다면)"
)
PROMOTIONAL_FAMILY_EXCLUSIONS = {
    "score_conversion": re.compile(
        r"바꾸(?:는|기\s+위한)\s+(?:방법|기준|과정|전략|질문|표현)"
    ),
    "outcome_transform": re.compile(
        r"바꾸(?:는|기\s+위한)\s+(?:방법|기준|과정|전략|질문|표현)"
    ),
    # A later grammatical subject, rather than the earlier outcome noun, is
    # what accumulates: "실력 ... 과정이 꾸준히 쌓이도록".
    "outcome_stack": re.compile(
        r"(?:과정|기록|자료|내용|단계|흐름|습관|태도)(?:이|가)[^.!?]{0,12}(?:쌓|누적)"
    ),
    # These constructions explicitly reject achievement as the operative
    # subject, or attach the predicate to a concrete planning object.
    "outcome_achievement": re.compile(
        r"(?:성취도|성취)(?:만|를)?\s*(?:보지\s+않|통해|기준으로|바탕으로|근거로)|"
        r"(?:학습\s+)?(?:순서|계획|기준|범위|자료|기록|바탕)(?:을|를)[^.!?]{0,12}"
        r"(?:만들|확장|강화|완성)"
    ),
    # An ability used as evidence/criterion is not the object of a later
    # predicate; nor is a separately marked planning noun.
    "outcome_ability": re.compile(
        r"(?:력|역량|능력|감각|체력|자신감|기본기)(?:을|를)\s*"
        r"(?:통해|기준으로|바탕으로|근거로)|"
        r"(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:만들|확장|강화|완성)"
    ),
    "outcome_appearance": re.compile(
        r"(?:기록|계획|상담|다음\s+풀이|학습\s+순서|자료|피드백)(?:에|으로)"
        r"[^.!?]{0,4}반영되|"
        r"(?:나오|따라오|반영되|회복|도달)[^.!?]{0,12}(?:는지|여부|원인|기준|확인|점검|분석)"
    ),
    "outcome_destination_link": re.compile(
        r"(?:연결|이어지)[^.!?]{0,12}(?:는지|여부|원인|기준|확인|점검|분석)"
    ),
    "outcome_ability_finite": re.compile(
        r"(?:력|역량|능력|감각|체력|자신감|기본기)(?:을|를)\s*"
        r"(?:통해|기준으로|바탕으로|근거로)|"
        r"(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:만들|확장|강화|완성|잡|다지)"
    ),
    "outcome_change_make": re.compile(
        r"(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:만들|만듭|확장|강화|완성)"
    ),
    "outcome_skill_build": re.compile(
        r"(?:(?:학습\s+)?(?:바탕|계획|순서|기준|자료|기록|범위)(?:을|를)[^.!?]{0,12}"
        r"(?:잡|다지|쌓)|(?:쌓(?:는|을\s+수\s+있는)|다지는|잡는)\s+"
        r"(?:과정|연습|계획|전략|방법|기준|방향))"
    ),
    "academic_omission_control": re.compile(
        r"(?:줄여|줄이도록)[^.!?]{0,30}(?:학습\s+변화를\s+목표로|유형을\s+반복\s+점검)|"
        r"재발\s+방지\s+전략[^.!?]{0,18}세웁"
    ),
}
ASPIRATIONAL_CLAIM = re.compile(r"(?:학습\s+변화|점수\s+변화)[^.!?]{0,32}목표")
SOFTENING_TERMS = (
    "단정", "보장보다", "보장하지", "약속하지", "피해야", "아니",
    "기대하기보다", "보장한다는 말보다", "보장한다는 표현보다",
    "보장 문구보다", "확인해야", "확인하는 표현",
    "점수 약속보다", "성적 약속보다",
)
UNSUPPORTED_OPERATIONAL_PATTERNS = {
    "phone": re.compile(r"(?<!\d)01[016789]-?\d{3,4}-?\d{4}(?!\d)"),
    "registration": re.compile(r"(?:교육지원청.{0,24}(?:등록|제\s*\d+호)|등록번호)"),
    "fee_amount": re.compile(r"(?:교습비|수업료).{0,16}\d[\d,]*\s*원"),
    "opening_hours": re.compile(r"(?:운영시간|수업시간).{0,20}(?:오전|오후|\d{1,2}:\d{2})"),
    "parking_assertion": re.compile(r"주차.{0,10}(?:가능|제공|지원)"),
    "vehicle_assertion": re.compile(r"(?:차량|셔틀).{0,10}(?:운행|제공|지원|가능)"),
}
ADMIN_REGION_ALIASES = {
    "서울": "서울특별시", "서울특별시": "서울특별시",
    "부산": "부산광역시", "부산광역시": "부산광역시",
    "대구": "대구광역시", "대구광역시": "대구광역시",
    "인천": "인천광역시", "인천광역시": "인천광역시",
    "광주": "광주광역시", "광주광역시": "광주광역시",
    "대전": "대전광역시", "대전광역시": "대전광역시",
    "울산": "울산광역시", "울산광역시": "울산광역시",
    "세종": "세종특별자치시", "세종시": "세종특별자치시", "세종특별자치시": "세종특별자치시",
    "경기": "경기도", "경기도": "경기도",
    "강원": "강원특별자치도", "강원도": "강원특별자치도", "강원특별자치도": "강원특별자치도",
    "충북": "충청북도", "충청북도": "충청북도",
    "충남": "충청남도", "충청남도": "충청남도",
    "전북": "전북특별자치도", "전라북도": "전북특별자치도", "전북특별자치도": "전북특별자치도",
    "전남": "전라남도", "전라남도": "전라남도",
    "경북": "경상북도", "경상북도": "경상북도",
    "경남": "경상남도", "경상남도": "경상남도",
    "제주": "제주특별자치도", "제주도": "제주특별자치도", "제주특별자치도": "제주특별자치도",
}


class AuditFailure(RuntimeError):
    pass


class Findings:
    def __init__(self) -> None:
        self.counts: Counter[str] = Counter()
        self.samples: dict[str, list[str]] = defaultdict(list)
        self.warnings: Counter[str] = Counter()
        self.warning_samples: dict[str, list[str]] = defaultdict(list)

    def add(self, code: str, detail: str) -> None:
        self.counts[code] += 1
        if len(self.samples[code]) < 8:
            self.samples[code].append(detail)

    def warn(self, code: str, detail: str) -> None:
        self.warnings[code] += 1
        if len(self.warning_samples[code]) < 8:
            self.warning_samples[code].append(detail)

    @property
    def total(self) -> int:
        return sum(self.counts.values())

    def payload(self) -> dict[str, object]:
        return {
            "error_count": self.total,
            "error_codes": dict(sorted(self.counts.items())),
            "error_samples": dict(sorted(self.samples.items())),
            "warning_count": sum(self.warnings.values()),
            "warning_codes": dict(sorted(self.warnings.items())),
            "warning_samples": dict(sorted(self.warning_samples.items())),
        }


@dataclass(frozen=True)
class SourceRecord:
    profile: str
    workbook: str
    row: int
    locality: str
    raw_html: str
    raw_sha256: str
    body_text: str
    corrected_localities: tuple[str, ...]


@dataclass(frozen=True)
class Manuscript:
    opening_tag: str
    intro: str
    headings: tuple[str, ...]
    paragraphs: tuple[str, ...]
    html: str
    text: str


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    return sha256_bytes(path.read_bytes())


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", html.unescape(str(value or ""))).strip()


def punctuation_balance_issues(value: str) -> tuple[str, ...]:
    text = str(value or "")
    issues: list[str] = []
    for opening, closing, label in (("‘", "’", "single_quote"), ("“", "”", "double_quote"), ("(", ")", "parenthesis")):
        opened = text.count(opening)
        closed = text.count(closing)
        if opened != closed:
            issues.append(f"{label}:{opened}:{closed}")
    return tuple(issues)


def visible_text(source: str) -> str:
    source = re.sub(r"<(?:script|style)\b.*?</(?:script|style)>", " ", source, flags=re.I | re.S)
    return clean(re.sub(r"<[^>]+>", " ", source))


def extract_one(pattern: str, source: str) -> str:
    match = re.search(pattern, source, flags=re.I | re.S)
    return match.group(1).strip() if match else ""


def attributes(source: str) -> dict[str, str]:
    return {
        key.lower(): html.unescape(value)
        for key, value in re.findall(r"([:\w-]+)\s*=\s*['\"]([^'\"]*)['\"]", source)
    }


def distribution(values: Iterable[float | int]) -> dict[str, float | int]:
    ordered = sorted(values)
    if not ordered:
        return {"min": 0, "median": 0, "p90": 0, "max": 0}
    return {
        "min": ordered[0],
        "median": round(float(statistics.median(ordered)), 6),
        "p90": ordered[min(len(ordered) - 1, int(len(ordered) * .9))],
        "max": ordered[-1],
    }


def digest_rows(rows: Iterable[str]) -> str:
    return sha256_bytes(("\n".join(rows) + "\n").encode("utf-8"))


def relative(path: Path) -> str:
    return path.resolve().relative_to(ROOT.resolve()).as_posix()


def page_url_from_canonical(source: str) -> str:
    return html.unescape(extract_one(r'<link\s+rel="canonical"\s+href="([^"]+)"', source))


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise AuditFailure(f"cannot import module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def git_blobs(commit: str, rels: Iterable[str]) -> dict[str, bytes]:
    ordered = sorted(set(rels))
    request = "".join(f"{commit}:{rel}\n" for rel in ordered).encode("utf-8")
    completed = subprocess.run(
        ["git", "cat-file", "--batch"], cwd=ROOT, input=request,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True,
    )
    payload = completed.stdout
    offset = 0
    result: dict[str, bytes] = {}
    for rel in ordered:
        end = payload.find(b"\n", offset)
        if end < 0:
            raise AuditFailure(f"truncated git blob header: {rel}")
        header = payload[offset:end]
        offset = end + 1
        if header.endswith(b" missing"):
            raise AuditFailure(f"baseline blob missing: {commit}:{rel}")
        parts = header.rsplit(b" ", 2)
        if len(parts) != 3 or parts[1] != b"blob":
            raise AuditFailure(f"unexpected git blob response for {rel}: {header!r}")
        size = int(parts[2])
        body_end = offset + size
        if body_end >= len(payload) or payload[body_end:body_end + 1] != b"\n":
            raise AuditFailure(f"truncated git blob body: {rel}")
        result[rel] = payload[offset:body_end]
        offset = body_end + 1
    if offset != len(payload):
        raise AuditFailure("unexpected trailing bytes from git cat-file")
    return result


def file_manifest(paths: Iterable[Path]) -> dict[str, str]:
    return {relative(path): sha256_file(path) for path in sorted(set(paths)) if path.is_file()}


def parse_graph(source: str) -> list[dict]:
    matches = re.findall(r'<script\s+type="application/ld\+json"[^>]*>(.*?)</script>', source, flags=re.I | re.S)
    if len(matches) != 1:
        raise ValueError(f"JSON-LD block count={len(matches)}")
    payload = json.loads(html.unescape(matches[0]))
    graph = payload.get("@graph") if isinstance(payload, dict) else None
    if not isinstance(graph, list) or not all(isinstance(node, dict) for node in graph):
        raise ValueError("JSON-LD @graph is missing or invalid")
    return graph


def node_types(node: dict) -> set[str]:
    value = node.get("@type", [])
    if isinstance(value, str):
        return {value}
    return set(value) if isinstance(value, list) else set()


def first_node(nodes: list[dict], kind: str) -> dict:
    return next((node for node in nodes if kind in node_types(node)), {})


def manuscript_from_html(source: str) -> Manuscript:
    match = re.search(
        r'(<section\s+class="section manuscript-wrap"[^>]*>)\s*<article\b[^>]*>(.*?)</article>\s*</section>',
        source, flags=re.I | re.S,
    )
    if not match:
        raise ValueError("manuscript wrapper not found")
    opening, body = match.groups()
    intro_html = extract_one(r'<div\s+class="manuscript-intro"[^>]*>.*?<p>(.*?)</p>\s*</div>', body)
    intro = visible_text(intro_html)
    headings: list[str] = []
    paragraphs: list[str] = []
    for block in re.findall(r'<section\s+class="manuscript-section"[^>]*>(.*?)</section>', body, flags=re.I | re.S):
        heading = visible_text(extract_one(r'<h2\b[^>]*>(.*?)</h2>', block))
        if heading:
            headings.append(heading)
        paragraphs.extend(
            item for item in (visible_text(value) for value in re.findall(r'<p\b[^>]*>(.*?)</p>', block, flags=re.I | re.S))
            if item
        )
    text = clean(" ".join([intro, *headings, *paragraphs]))
    return Manuscript(opening, intro, tuple(headings), tuple(paragraphs), body, text)


def class_fragment(source: str, tag: str, class_name: str) -> str:
    return extract_one(
        rf'<{tag}\b[^>]*class="[^"]*\b{re.escape(class_name)}\b[^"]*"[^>]*>(.*?)</{tag}>',
        source,
    )


def visible_faq(source: str) -> tuple[tuple[str, str], ...]:
    values = re.findall(
        r'<details(?:\s+[^>]*)?>\s*<summary>(.*?)</summary>\s*<p>(.*?)</p>\s*</details>',
        source, flags=re.I | re.S,
    )
    return tuple((visible_text(question), visible_text(answer)) for question, answer in values)


def center_fact_contract(source: str) -> dict[str, object]:
    aside = class_fragment(source, "aside", "local-info-card")
    fields: dict[str, str] = {}
    for label, value in re.findall(r'<dt>(.*?)</dt>\s*<dd>(.*?)</dd>', aside, flags=re.I | re.S):
        fields[visible_text(label)] = visible_text(value)
    grades = tuple(
        (visible_text(subject), visible_text(values))
        for subject, values in re.findall(
            r'<li>\s*<strong>(.*?)</strong>\s*<span>(.*?)</span>\s*</li>', aside, flags=re.I | re.S,
        )
    )
    schools = tuple(
        sorted(value for label, value in fields.items() if "학교 참고" in label)
    )
    return {
        "all_fields": fields,
        "region": fields.get("지역", ""),
        "address": fields.get("제공 주소", ""),
        "registration": fields.get("등록 정보", ""),
        "schools": schools,
        "grades": grades,
        "visible_sha256": sha256_bytes(clean(aside).encode("utf-8")),
    }


def image_contract(source: str) -> dict[str, object]:
    images: list[tuple[str, str, str, str]] = []
    for tag in re.findall(r'<img\b[^>]*>', source, flags=re.I):
        attrs = attributes(tag)
        images.append((attrs.get("src", ""), attrs.get("width", ""), attrs.get("height", ""), attrs.get("alt", "")))
    meta: list[tuple[str, str]] = []
    for tag in re.findall(r'<meta\b[^>]*>', source, flags=re.I):
        attrs = attributes(tag)
        key = attrs.get("property") or attrs.get("name") or ""
        if "image" in key.lower():
            meta.append((key, attrs.get("content", "")))
    return {"img": tuple(images), "meta": tuple(meta)}


def tuition_urls(source: str, graph: list[dict]) -> tuple[str, ...]:
    found = set(re.findall(r'https://drive\.google\.com/[^"<\s]+', html.unescape(source)))

    def walk(value: object) -> None:
        if isinstance(value, dict):
            name = clean(value.get("name"))
            url = clean(value.get("url"))
            if "교습" in name and url:
                found.add(url)
            for nested in value.values():
                walk(nested)
        elif isinstance(value, list):
            for nested in value:
                walk(nested)

    walk(graph)
    return tuple(sorted(found))


def immutable_graph(nodes: list[dict]) -> list[dict]:
    result = copy.deepcopy(nodes)
    for node in result:
        if "Article" in node_types(node):
            node.pop("dateModified", None)
            node.pop("hasPart", None)
    return result


def fact_bundle(source: str) -> dict[str, object]:
    graph = parse_graph(source)
    article = first_node(graph, "Article")
    webpage = first_node(graph, "WebPage")
    faq_node = first_node(graph, "FAQPage")
    faq_schema = tuple(
        (clean(item.get("name")), clean((item.get("acceptedAnswer") or {}).get("text")))
        for item in faq_node.get("mainEntity", []) if isinstance(item, dict)
    )
    title = visible_text(extract_one(r'<title>(.*?)</title>', source))
    h1_values = re.findall(r'<h1\b[^>]*>(.*?)</h1>', source, flags=re.I | re.S)
    h1 = visible_text(h1_values[0]) if len(h1_values) == 1 else ""
    description = html.unescape(extract_one(r'<meta\s+name="description"\s+content="([^"]*)"', source))
    quick = class_fragment(source, "div", "local-answer-grid")
    media = extract_one(r'<section\s+class="local-media-section"[^>]*>(.*?)</section>', source)
    center = center_fact_contract(source)
    return {
        "canonical": page_url_from_canonical(source),
        "og_url": html.unescape(extract_one(r'<meta\s+property="og:url"\s+content="([^"]+)"', source)),
        "title": title,
        "h1": h1,
        "description": description,
        "quick_answer": visible_text(quick),
        "center": center,
        "images": image_contract(source),
        "media_visible": visible_text(media),
        "faq_visible": visible_faq(source),
        "faq_schema": faq_schema,
        "tuition_urls": tuition_urls(source, graph),
        "graph": graph,
        "graph_immutable_sha256": sha256_bytes(
            json.dumps(immutable_graph(graph), ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
        ),
        "article_date_published": clean(article.get("datePublished")),
        "article_date_modified": clean(article.get("dateModified")),
        "article_has_part": tuple(
            clean(item.get("name")) for item in article.get("hasPart", []) if isinstance(item, dict)
        ),
        "article_headline": clean(article.get("headline")),
        "article_description": clean(article.get("description")),
        "webpage_description": clean(webpage.get("description")),
    }


def source_records(
    generator, plan, archive: Path, common_dir: Path, findings: Findings,
) -> tuple[dict[tuple[str, str], SourceRecord], dict[str, set[str]]]:
    workbooks, _ = generator.read_archive(archive)
    result: dict[tuple[str, str], SourceRecord] = {}
    locality_sets: dict[str, set[str]] = {}
    detail_docs = [doc for doc in plan.documents if getattr(doc, "profile", "") in {item[0] for item in PROFILE_SPECS}]
    docs_by_key: dict[tuple[str, str], object] = {}
    for doc in detail_docs:
        key = (clean(doc.profile), clean(doc.locality))
        if key in docs_by_key:
            findings.add("plan_locality_duplicate", f"{key[0]}/{key[1]}")
        docs_by_key[key] = doc
    for profile, workbook, category in PROFILE_SPECS:
        docs = [doc for doc in detail_docs if clean(doc.profile) == profile]
        rows = workbooks.get(workbook, [])
        if len(docs) != EXPECTED_DETAILS_PER_PROFILE:
            findings.add("source_profile_document_count", f"{profile}: {len(docs)}")
        if len(rows) != EXPECTED_DETAILS_PER_PROFILE:
            findings.add("source_profile_row_count", f"{workbook}: {len(rows)}")
        base = generator.load_base(profile, common_dir)
        locality_order = [clean(value) for value in base.load_centers()]
        if len(locality_order) != EXPECTED_DETAILS_PER_PROFILE or len(set(locality_order)) != len(locality_order):
            findings.add(
                "source_csv_locality_order", f"{profile}: rows={len(locality_order)}, unique={len(set(locality_order))}",
            )
        localities = set(locality_order)
        locality_sets[profile] = localities
        for offset in range(min(len(locality_order), len(rows))):
            row_number = offset + 2
            locality = locality_order[offset]
            raw = rows[offset]
            doc = docs_by_key.get((profile, locality))
            if doc is None:
                findings.add("plan_locality_missing", f"{profile}/{locality}: source row {row_number}")
                continue
            raw_sha = sha256_bytes(raw.encode("utf-8"))
            try:
                _, intro, sections = generator.source_fragment(raw)
            except Exception as exc:
                findings.add("source_fragment", f"{profile} row {row_number}: {exc}")
                intro, sections = "", []
            body_text = clean(" ".join([intro, *[paragraph for _, paragraphs in sections for paragraph in paragraphs]]))
            raw_without_h1 = re.sub(r"<h1\b[^>]*>.*?</h1>", "", raw, flags=re.I | re.S)
            raw_body = visible_text(raw_without_h1)
            foreign_counts = sorted(
                ((raw_body.count(item), item) for item in localities - {locality} if item and raw_body.count(item)),
                reverse=True,
            )
            corrected: tuple[str, ...] = ()
            if foreign_counts:
                top_count, top_locality = foreign_counts[0]
                intended = raw_body.count(locality)
                if intended == 0 or (top_count >= 3 and top_count >= intended * 2):
                    corrected = (top_locality,)
            record = SourceRecord(
                profile, workbook, row_number, locality, raw, raw_sha, body_text, corrected,
            )
            key = (profile, locality)
            if key in result:
                findings.add("source_mapping_duplicate", f"{profile}/{locality}")
            result[key] = record
            if doc.source_sha256 != raw_sha:
                findings.add("source_plan_sha_mismatch", f"{profile}/{locality}: row={row_number}")
            expected_slug = re.sub(r"\s+", "", locality)
            if doc.path.parent.name != expected_slug or doc.path.parent.parent.name != category:
                findings.add(
                    "source_route_mapping", f"{relative(doc.path)}: locality={locality!r}, category={category!r}",
                )
    if len(result) != EXPECTED_DETAIL_DOCUMENTS:
        findings.add("source_mapping_count", f"{len(result)} != {EXPECTED_DETAIL_DOCUMENTS}")
    return result, locality_sets


GENERIC_TOKENS = {
    "학습", "학생", "수업", "확인", "과목", "국어", "영어", "수학", "학원",
    "상담", "기준", "현재", "자료", "계획", "과정", "지역", "센터", "합니다",
    "있습니다", "됩니다", "초등학생", "중학생", "고등학생", "국영수학원",
}


def normalized_tokens(value: str, localities: Iterable[str]) -> list[str]:
    text = clean(value).lower()
    for locality in sorted(set(localities), key=len, reverse=True):
        text = text.replace(locality.lower(), " 지역 ")
    tokens = re.findall(r"[가-힣a-z0-9]{2,}", text)
    return [token for token in tokens if token not in GENERIC_TOKENS and not token.isdigit()]


def token_metrics(left: str, right: str, localities: Iterable[str]) -> tuple[float, float]:
    left_set = set(normalized_tokens(left, localities))
    right_set = set(normalized_tokens(right, localities))
    if not left_set and not right_set:
        return 1.0, 1.0
    intersection = len(left_set & right_set)
    jaccard = intersection / max(1, len(left_set | right_set))
    recall = intersection / max(1, len(left_set))
    return jaccard, recall


def corpus_normalized(value: str, localities: Iterable[str]) -> str:
    text = clean(value).lower()
    for locality in sorted(set(localities), key=len, reverse=True):
        text = text.replace(locality.lower(), "{지역}")
    text = re.sub(r"\d+(?:[.,]\d+)*", "{n}", text)
    return re.sub(r"\s+", " ", text).strip()


def word_shingles(value: str, size: int = 5) -> set[tuple[str, ...]]:
    words = re.findall(r"[가-힣a-z0-9{}]+", value.lower())
    if len(words) < size:
        return {tuple(words)} if words else set()
    return {tuple(words[index:index + size]) for index in range(len(words) - size + 1)}


def promotional_hits(value: str) -> tuple[list[str], int]:
    unsafe: list[str] = []
    softened = 0
    sentences = re.split(r"(?<=[.!?])\s+", clean(value))
    for sentence in sentences:
        # Preserve nominal/advisory language while retaining any later finite
        # outcome assertion in the same sentence for the high-recall gates.
        scan_sentence = OUTCOME_SAFE_NOMINAL_SPANS.sub(
            lambda match: " " * len(match.group(0)), sentence
        )
        for label, pattern in PROMOTIONAL_PATTERNS.items():
            matches = tuple(pattern.finditer(scan_sentence))
            if not matches:
                continue
            exclusion = PROMOTIONAL_FAMILY_EXCLUSIONS.get(label)
            if exclusion is not None and all(
                exclusion.search(scan_sentence[match.start():match.end() + 48])
                for match in matches
            ):
                continue
            if any(term in sentence for term in SOFTENING_TERMS):
                softened += 1
            else:
                unsafe.append(f"{label}: {sentence[:180]}")
    return unsafe, softened


def administrative_mentions(value: str, locality: str) -> set[str]:
    prefixes = "|".join(re.escape(item) for item in sorted(ADMIN_REGION_ALIASES, key=len, reverse=True))
    pattern = re.compile(
        rf"(?<![가-힣])({prefixes})(?:\s+[가-힣A-Za-z0-9]+(?:시|군|구)){{0,3}}\s+{re.escape(locality)}(?![가-힣])"
    )
    mentions: set[str] = set()
    for match in pattern.finditer(clean(value)):
        phrase = clean(match.group(0))
        alias = match.group(1)
        mentions.add(ADMIN_REGION_ALIASES[alias] + phrase[len(alias):])
    return mentions


def administrative_compatible(expected: str, actual: str, locality: str) -> bool:
    expected_tokens = [token for token in clean(expected).split() if token != locality]
    actual_tokens = [token for token in clean(actual).split() if token != locality]
    if not expected_tokens or not actual_tokens or expected_tokens[0] != actual_tokens[0]:
        return False
    expected_set = set(expected_tokens)
    actual_set = set(actual_tokens)
    return expected_set.issubset(actual_set) or actual_set.issubset(expected_set)


def locality_mentions(value: str, locality: str, candidates: Iterable[str]) -> list[str]:
    text = clean(value)
    target_compact = re.sub(r"\s+", "", locality)
    particles = r"(?:에서|으로|부터|까지|처럼|보다|은|는|이|가|을|를|의|에|와|과|도|만)?"
    found: list[str] = []
    for candidate in sorted(set(candidates) - {locality}, key=len, reverse=True):
        candidate_compact = re.sub(r"\s+", "", candidate)
        # Canonical/colloquial variants such as 대구유천동↔유천동 and
        # 분당 정자동↔정자동 are not foreign branches.
        if candidate_compact in target_compact or target_compact in candidate_compact:
            continue
        if (
            candidate in LOCALITY_ALIASES.get(locality, set())
            or locality in LOCALITY_ALIASES.get(candidate, set())
        ):
            continue
        if re.search(
            rf"(?<![가-힣A-Za-z0-9]){re.escape(candidate)}{particles}(?![가-힣A-Za-z0-9])",
            text,
        ):
            found.append(candidate)
    return found


def sentences_with(value: str, needle: str) -> list[str]:
    return [
        sentence for sentence in re.split(r"(?<=[.!?])\s+", clean(value))
        if needle in sentence
    ]


def embedded_target_corruptions(value: str, locality: str) -> list[str]:
    compact_locality = re.sub(r"\s+", "", locality)
    patterns = [
        re.compile(re.escape(compact_locality[:1] + compact_locality)) if compact_locality else re.compile(r"(?!)"),
        re.compile(rf"정\s*{re.escape(locality)}\s*게"),
        re.compile(rf"비\s*{re.escape(locality)}\s*기보다"),
    ]
    return [match.group(0) for pattern in patterns for match in pattern.finditer(clean(value))]


def audit_page(
    stage: str,
    path: Path,
    profile: str,
    locality: str,
    baseline_bytes: bytes,
    raw_worktree_before: bytes,
    candidate_bytes: bytes,
    source_record: SourceRecord,
    all_localities: set[str],
    findings: Findings,
) -> dict[str, object]:
    rel = relative(path)
    label = f"{stage}:{rel}"
    try:
        baseline = baseline_bytes.decode("utf-8")
        candidate = candidate_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        findings.add("encoding_utf8", f"{label}: {exc}")
        return {"path": rel, "profile": profile, "locality": locality, "invalid": True}
    if candidate.startswith("\ufeff") or "\x00" in candidate or "�" in candidate:
        findings.add("encoding_artifact", label)
    try:
        current_manuscript = manuscript_from_html(baseline)
        new_manuscript = manuscript_from_html(candidate)
    except ValueError as exc:
        findings.add("manuscript_parse", f"{label}: {exc}")
        return {"path": rel, "profile": profile, "locality": locality, "invalid": True}

    marker = attributes(new_manuscript.opening_tag)
    if marker.get("data-revision") != REVISION_MARKER:
        findings.add("provenance_revision_marker", f"{label}: {marker.get('data-revision')!r}")
    if marker.get("data-source-row") != str(source_record.row):
        findings.add(
            "provenance_source_row", f"{label}: {marker.get('data-source-row')!r} != {source_record.row}",
        )
    if marker.get("data-source-sha256") != source_record.raw_sha256:
        findings.add("provenance_source_sha", f"{label}: marker/source mismatch")

    try:
        before_facts = fact_bundle(baseline)
        after_facts = fact_bundle(candidate)
    except (ValueError, json.JSONDecodeError) as exc:
        findings.add("fact_schema_parse", f"{label}: {exc}")
        return {"path": rel, "profile": profile, "locality": locality, "invalid": True}

    simple_preservation = (
        ("canonical", "preserve_canonical"),
        ("og_url", "preserve_og_url"),
        ("title", "preserve_title"),
        ("h1", "preserve_h1"),
        ("description", "preserve_description"),
        ("quick_answer", "preserve_verified_facts"),
        ("media_visible", "preserve_media"),
        ("images", "preserve_images"),
        ("tuition_urls", "preserve_tuition"),
        ("faq_visible", "preserve_faq"),
        ("article_date_published", "preserve_date_published"),
        ("graph_immutable_sha256", "preserve_schema_facts"),
    )
    parity: dict[str, bool] = {}
    for field, code in simple_preservation:
        equal = before_facts[field] == after_facts[field]
        parity[field] = equal
        if not equal:
            findings.add(code, f"{label}: {field}")

    before_center = before_facts["center"]
    after_center = after_facts["center"]
    center_contract = (
        ("region", "preserve_region"),
        ("address", "preserve_address"),
        ("registration", "preserve_registration"),
        ("schools", "preserve_schools"),
        ("grades", "preserve_grades"),
        ("all_fields", "preserve_center_facts"),
    )
    for field, code in center_contract:
        equal = before_center[field] == after_center[field]
        parity[field] = equal
        if not equal:
            findings.add(code, f"{label}: {field}")

    if after_facts["faq_visible"] != after_facts["faq_schema"] or len(after_facts["faq_visible"]) != 4:
        findings.add(
            "visible_jsonld_faq_parity",
            f"{label}: visible={len(after_facts['faq_visible'])}, schema={len(after_facts['faq_schema'])}",
        )
    if after_facts["article_has_part"] != new_manuscript.headings:
        findings.add(
            "visible_jsonld_heading_parity",
            f"{label}: visible={new_manuscript.headings!r}, schema={after_facts['article_has_part']!r}",
        )
    if after_facts["article_headline"] != after_facts["h1"]:
        findings.add("visible_jsonld_headline_parity", label)
    if not (
        after_facts["description"] == after_facts["article_description"] == after_facts["webpage_description"]
    ):
        findings.add("visible_jsonld_description_parity", label)
    if after_facts["article_date_modified"] != REVISION_DATE:
        findings.add(
            "article_revision_date", f"{label}: {after_facts['article_date_modified']!r}",
        )

    expected_slug = re.sub(r"\s+", "", locality)
    if path.parent.name != expected_slug:
        findings.add("locality_route", f"{label}: slug={path.parent.name!r}, locality={locality!r}")
    if locality not in visible_text(candidate):
        findings.add("locality_page_absent", label)
    expected_area = clean(after_center.get("region"))
    expected_mentions = administrative_mentions(expected_area, locality)
    expected_normalized_area = next(iter(expected_mentions), expected_area)
    area_mentions = administrative_mentions(new_manuscript.text, locality)
    wrong_areas = sorted(
        item for item in area_mentions
        if not administrative_compatible(expected_normalized_area, item, locality)
    )
    if wrong_areas:
        findings.add(
            "locality_administrative_mismatch",
            f"{label}: expected={expected_normalized_area!r}, actual={wrong_areas[:6]!r}",
        )
    masked = new_manuscript.text.replace(locality, " {대상지역} ")
    detected_foreign = locality_mentions(masked, locality, all_localities)
    verified_address = clean(after_center.get("address"))
    verified_foreign_mentions: list[str] = []
    foreign_mentions: list[str] = []
    for foreign in detected_foreign:
        contexts = sentences_with(new_manuscript.text, foreign)
        if (
            verified_address and foreign in verified_address and contexts
            and all(verified_address in context for context in contexts)
        ):
            verified_foreign_mentions.append(foreign)
        else:
            foreign_mentions.append(foreign)
    if foreign_mentions:
        findings.add("locality_foreign_contamination", f"{label}: {foreign_mentions[:8]}")
    leaked_corrections = locality_mentions(masked, locality, source_record.corrected_localities)
    if leaked_corrections:
        findings.add("locality_known_contamination_not_corrected", f"{label}: {leaked_corrections}")
    if locality + locality in new_manuscript.text or f"{locality} {locality}" in new_manuscript.text:
        findings.add("locality_accidental_duplication", label)
    embedded_corruptions = embedded_target_corruptions(new_manuscript.text, locality)
    for corruption in embedded_corruptions:
        findings.add("locality_embedded_target_corruption", f"{label}: {corruption!r}")

    if not new_manuscript.intro:
        findings.add("natural_intro_missing", label)
    if not 3 <= len(new_manuscript.headings) <= 6:
        findings.add("natural_section_count", f"{label}: {len(new_manuscript.headings)}")
    if len(set(new_manuscript.headings)) != len(new_manuscript.headings):
        findings.add("natural_heading_duplicate", label)
    if not 900 <= len(new_manuscript.text) <= 7500:
        findings.add("natural_visible_length", f"{label}: {len(new_manuscript.text)}")
    all_paragraphs = (new_manuscript.intro, *new_manuscript.paragraphs)
    natural_blocks = tuple(block for block in (*all_paragraphs, *new_manuscript.headings) if block)
    normalized_within: set[str] = set()
    for index, paragraph in enumerate(all_paragraphs):
        norm = re.sub(r"\W+", "", paragraph)
        if norm in normalized_within:
            findings.add("natural_within_page_duplicate", f"{label}: paragraph {index}")
        normalized_within.add(norm)
        if not 35 <= len(paragraph) <= 850:
            findings.add("natural_paragraph_length", f"{label}: paragraph {index} length={len(paragraph)}")
        if paragraph.endswith(")"):
            findings.add(
                "natural_parenthetical_fragment_ending",
                f"{label}: paragraph {index} ends={paragraph[-40:]!r}",
            )
        elif paragraph and paragraph[-1] not in ".?!다요죠":
            findings.warn("natural_paragraph_ending", f"{label}: paragraph {index} ends={paragraph[-10:]!r}")
        for sentence in re.split(r"(?<=[.!?])\s+", paragraph):
            if len(sentence) > 340:
                findings.add("natural_sentence_too_long", f"{label}: {len(sentence)} chars")
    for phrase in KNOWN_AWKWARD_COPY:
        if any(phrase in block for block in natural_blocks):
            findings.add("natural_known_copy_error", f"{label}: {phrase!r}")
    for block_index, block in enumerate(natural_blocks):
        for issue in punctuation_balance_issues(block):
            findings.add(
                "natural_unbalanced_punctuation", f"{label}: block {block_index}: {issue}"
            )
    for family, pattern in AWKWARD_PATTERNS.items():
        for block_index, block in enumerate(natural_blocks):
            for match in pattern.finditer(block):
                findings.add(
                    f"natural_awkward_{family}",
                    f"{label}: block {block_index}: {match.group(0)!r}",
                )
    for phrase in READER_META_TERMS:
        if any(phrase in block for block in natural_blocks):
            findings.add("natural_reader_meta_copy", f"{label}: {phrase!r}")
    for phrase in FORBIDDEN_SOURCE_MARKUP:
        if phrase.lower() in new_manuscript.html.lower():
            findings.add("natural_source_artifact", f"{label}: {phrase!r}")
    for block_index, block in enumerate(natural_blocks):
        repeated = re.search(r"(?<![가-힣])([가-힣]{2,})(?:\s+\1){1,}", block)
        if repeated:
            findings.add(
                "natural_repeated_word", f"{label}: block {block_index}: {repeated.group(0)!r}"
            )
    question_count = new_manuscript.text.count("?")
    direct_question_allowed = profile == "middle" and locality == "풍동" and question_count == 1
    if question_count and not direct_question_allowed:
        findings.add("natural_question_mark_residue", f"{label}: {question_count}")

    unsafe_claims: list[str] = []
    softened_claims = 0
    for block in natural_blocks:
        block_unsafe, block_softened = promotional_hits(block)
        unsafe_claims.extend(block_unsafe)
        softened_claims += block_softened
    for claim in unsafe_claims:
        family, _, excerpt = claim.partition(": ")
        findings.add(f"unverified_promotional_claim_{family}", f"{label}: {excerpt or claim}")
    operational_hits: list[str] = []
    for claim_type, pattern in UNSUPPORTED_OPERATIONAL_PATTERNS.items():
        if any(pattern.search(block) for block in natural_blocks):
            operational_hits.append(claim_type)
            findings.add(f"unverified_operational_claim_{claim_type}", label)
    for block in natural_blocks:
        for match in ASPIRATIONAL_CLAIM.finditer(block):
            findings.warn("aspirational_change_target", f"{label}: {match.group(0)!r}")

    source_jaccard, source_recall = token_metrics(source_record.body_text, new_manuscript.text, all_localities)
    current_jaccard, current_recall = token_metrics(current_manuscript.text, new_manuscript.text, all_localities)
    if source_recall < .12:
        findings.add("attachment_use_too_weak", f"{label}: source token recall={source_recall:.3f}")
    elif source_recall < .18:
        findings.warn("attachment_use_low", f"{label}: source token recall={source_recall:.3f}")
    if new_manuscript.text == current_manuscript.text:
        findings.add("new_manuscript_unchanged", label)
    if clean(new_manuscript.text) == clean(source_record.body_text):
        findings.add("source_manuscript_verbatim", label)

    source_unsafe, source_softened = promotional_hits(source_record.body_text)
    norm_text = corpus_normalized(new_manuscript.text, all_localities)
    provenance = {
        "path": rel,
        "profile": profile,
        "locality": locality,
        "source_workbook": source_record.workbook,
        "source_row": source_record.row,
        "source_sha256": source_record.raw_sha256,
        "current_sha256": sha256_bytes(baseline_bytes),
        "current_raw_worktree_sha256": sha256_bytes(raw_worktree_before),
        "current_raw_normalizes_to_verified": raw_worktree_before.replace(b"\r\n", b"\n") == baseline_bytes,
        "new_sha256": sha256_bytes(candidate_bytes),
        "current_manuscript_sha256": sha256_bytes(current_manuscript.text.encode("utf-8")),
        "new_manuscript_sha256": sha256_bytes(new_manuscript.text.encode("utf-8")),
        "source_chars": len(source_record.body_text),
        "current_chars": len(current_manuscript.text),
        "new_chars": len(new_manuscript.text),
        "source_new_jaccard": round(source_jaccard, 6),
        "source_new_recall": round(source_recall, 6),
        "current_new_jaccard": round(current_jaccard, 6),
        "current_new_recall": round(current_recall, 6),
        "corrected_localities": list(source_record.corrected_localities),
        "foreign_mentions": foreign_mentions,
        "verified_address_locality_mentions": verified_foreign_mentions,
        "embedded_target_corruptions": embedded_corruptions,
        "source_unsafe_claims": len(source_unsafe),
        "source_softened_claims": source_softened,
        "new_unsafe_claims": len(unsafe_claims),
        "new_softened_claims": softened_claims,
        "operational_hits": operational_hits,
        "fact_parity": all(parity.values()),
    }
    return {
        **provenance,
        "paragraphs": list(all_paragraphs),
        "headings": list(new_manuscript.headings),
        "normalized_text": norm_text,
        "shingles": word_shingles(norm_text),
    }


def audit_corpus(observations: list[dict[str, object]], all_localities: set[str], findings: Findings) -> dict[str, object]:
    valid = [item for item in observations if not item.get("invalid")]
    if len(valid) != EXPECTED_DETAIL_DOCUMENTS:
        findings.add("corpus_valid_document_count", f"{len(valid)} != {EXPECTED_DETAIL_DOCUMENTS}")
    source_hashes = [str(item["source_sha256"]) for item in valid]
    manuscript_hashes = [str(item["new_manuscript_sha256"]) for item in valid]
    if len(set(source_hashes)) != len(source_hashes):
        findings.add("corpus_source_hash_duplicate", f"unique={len(set(source_hashes))}/{len(source_hashes)}")
    if len(set(manuscript_hashes)) != len(manuscript_hashes):
        findings.add("corpus_new_manuscript_duplicate", f"unique={len(set(manuscript_hashes))}/{len(manuscript_hashes)}")

    paragraph_df: dict[str, set[str]] = defaultdict(set)
    sentence_df: dict[str, set[str]] = defaultdict(set)
    heading_sequence_df: dict[str, set[str]] = defaultdict(set)
    for item in valid:
        rel = str(item["path"])
        per_page: set[str] = set()
        for paragraph in item["paragraphs"]:
            normalized = corpus_normalized(str(paragraph), all_localities)
            if len(normalized) >= 30:
                per_page.add(normalized)
            for sentence in re.split(r"(?<=[.!?])\s+", normalized):
                if len(sentence) >= 45:
                    sentence_df[sentence].add(rel)
        for paragraph in per_page:
            paragraph_df[paragraph].add(rel)
        heading_sequence = " | ".join(
            corpus_normalized(str(heading), all_localities) for heading in item["headings"]
        )
        heading_sequence_df[heading_sequence].add(rel)

    paragraph_top = sorted(
        ((len(paths), text, sorted(paths)[:5]) for text, paths in paragraph_df.items()), reverse=True,
    )[:10]
    sentence_top = sorted(
        ((len(paths), text, sorted(paths)[:5]) for text, paths in sentence_df.items()), reverse=True,
    )[:10]
    heading_top = sorted(
        ((len(paths), text, sorted(paths)[:5]) for text, paths in heading_sequence_df.items()), reverse=True,
    )[:10]
    paragraph_max = paragraph_top[0][0] if paragraph_top else 0
    sentence_max = sentence_top[0][0] if sentence_top else 0
    heading_max = heading_top[0][0] if heading_top else 0
    if paragraph_max > 80:
        findings.add("corpus_paragraph_repetition", f"max document frequency={paragraph_max}")
    if sentence_max > 100:
        findings.add("corpus_sentence_repetition", f"max document frequency={sentence_max}")
    if heading_max > 80:
        findings.add("corpus_heading_sequence_repetition", f"max document frequency={heading_max}")

    postings: dict[tuple[str, ...], list[int]] = defaultdict(list)
    for index, item in enumerate(valid):
        for shingle in item["shingles"]:
            postings[shingle].append(index)
    shared: Counter[tuple[int, int]] = Counter()
    for members in postings.values():
        if not 2 <= len(members) <= 30:
            continue
        for left_index in range(len(members)):
            for right_index in range(left_index + 1, len(members)):
                shared[(members[left_index], members[right_index])] += 1
    pair_results: list[tuple[float, str, str, int]] = []
    for (left_index, right_index), shared_count in shared.most_common(5000):
        if shared_count < 8:
            break
        left = valid[left_index]
        right = valid[right_index]
        left_shingles = left["shingles"]
        right_shingles = right["shingles"]
        exact_shared = len(left_shingles & right_shingles)
        score = exact_shared / max(1, len(left_shingles | right_shingles))
        pair_results.append((score, str(left["path"]), str(right["path"]), exact_shared))
    pair_results.sort(reverse=True)
    top_pairs = pair_results[:15]
    if top_pairs and top_pairs[0][0] > .55:
        findings.add(
            "corpus_near_duplicate_pair",
            f"jaccard={top_pairs[0][0]:.3f}: {top_pairs[0][1]} <> {top_pairs[0][2]}",
        )

    return {
        "documents": len(valid),
        "unique_source_hashes": len(set(source_hashes)),
        "unique_new_manuscripts": len(set(manuscript_hashes)),
        "source_chars": distribution(int(item["source_chars"]) for item in valid),
        "current_chars": distribution(int(item["current_chars"]) for item in valid),
        "new_chars": distribution(int(item["new_chars"]) for item in valid),
        "source_new_jaccard": distribution(float(item["source_new_jaccard"]) for item in valid),
        "source_new_recall": distribution(float(item["source_new_recall"]) for item in valid),
        "current_new_jaccard": distribution(float(item["current_new_jaccard"]) for item in valid),
        "fact_parity_pages": sum(bool(item["fact_parity"]) for item in valid),
        "corrected_source_pages": sum(bool(item["corrected_localities"]) for item in valid),
        "source_unsafe_claims": sum(int(item["source_unsafe_claims"]) for item in valid),
        "new_unsafe_claims": sum(int(item["new_unsafe_claims"]) for item in valid),
        "new_softened_claims": sum(int(item["new_softened_claims"]) for item in valid),
        "paragraph_max_df": paragraph_max,
        "sentence_max_df": sentence_max,
        "heading_sequence_max_df": heading_max,
        "top_repeated_paragraphs": [
            {"document_frequency": count, "text": text[:240], "paths": paths}
            for count, text, paths in paragraph_top
        ],
        "top_repeated_sentences": [
            {"document_frequency": count, "text": text[:240], "paths": paths}
            for count, text, paths in sentence_top
        ],
        "top_heading_sequences": [
            {"document_frequency": count, "sequence": text, "paths": paths}
            for count, text, paths in heading_top
        ],
        "top_near_duplicate_pairs": [
            {"jaccard": round(score, 6), "left": left, "right": right, "shared_shingles": count}
            for score, left, right, count in top_pairs
        ],
    }


def sitemap_pairs(value: bytes) -> list[tuple[str, str]]:
    root = ET.fromstring(value.decode("utf-8"))
    pairs: list[tuple[str, str]] = []
    for node in root:
        location = next((clean(child.text) for child in node if child.tag.endswith("loc")), "")
        modified = next((clean(child.text) for child in node if child.tag.endswith("lastmod")), "")
        if location:
            pairs.append((location, modified))
    return pairs


def audit_sitemap(
    stage: str,
    baseline: bytes,
    raw_worktree_before: bytes,
    candidate: bytes,
    target_urls: set[str],
    findings: Findings,
) -> dict[str, object]:
    label = f"{stage}:sitemap.xml"
    try:
        before_pairs = sitemap_pairs(baseline)
        after_pairs = sitemap_pairs(candidate)
    except (UnicodeDecodeError, ET.ParseError) as exc:
        findings.add("sitemap_parse", f"{label}: {exc}")
        return {"invalid": True}
    before_urls = [url for url, _ in before_pairs]
    after_urls = [url for url, _ in after_pairs]
    if len(after_urls) != len(set(after_urls)):
        findings.add("sitemap_duplicate_url", label)
    if before_urls != after_urls:
        findings.add("sitemap_url_order_or_scope_changed", label)
    present_targets = {url for url in after_urls if url in target_urls}
    if present_targets != target_urls or len(present_targets) != EXPECTED_DETAIL_DOCUMENTS:
        findings.add(
            "sitemap_detail_target_scope",
            f"{label}: present={len(present_targets)}, expected={len(target_urls)}",
        )
    target_modified = {url: modified for url, modified in after_pairs if url in target_urls}
    wrong_dates = {url: value for url, value in target_modified.items() if value != REVISION_DATE}
    if wrong_dates:
        findings.add("sitemap_detail_lastmod", f"{label}: wrong={len(wrong_dates)} sample={list(wrong_dates.items())[:3]}")
    before_non_target = [(url, modified) for url, modified in before_pairs if url not in target_urls]
    after_non_target = [(url, modified) for url, modified in after_pairs if url not in target_urls]
    if before_non_target != after_non_target:
        findings.add("sitemap_non_target_changed", label)

    source = raw_worktree_before.decode("utf-8")
    pattern = re.compile(r"(<url><loc>([^<]+)</loc><lastmod>)([^<]+)(</lastmod></url>)")
    seen: set[str] = set()

    def replace(match: re.Match[str]) -> str:
        url = match.group(2)
        if url not in target_urls:
            return match.group(0)
        if url in seen:
            findings.add("sitemap_duplicate_detail_target", f"{label}: {url}")
        seen.add(url)
        return match.group(1) + REVISION_DATE + match.group(4)

    exact_expected = pattern.sub(replace, source).encode("utf-8")
    if seen != target_urls:
        findings.add("sitemap_regex_target_coverage", f"{label}: {len(seen)}/{len(target_urls)}")
    if candidate != exact_expected:
        findings.add("sitemap_only_lastmod_bytes_allowed", label)
    hub_prefixes = {
        f"{SITE_ORIGIN}/과목별학원/{category}/" for _, _, category in PROFILE_SPECS
    }
    hub_pairs = [(url, modified) for url, modified in after_pairs if unquote(html.unescape(url)) in hub_prefixes]
    return {
        "total_urls": len(after_pairs),
        "detail_targets": len(present_targets),
        "detail_lastmod_2026_08_22": sum(value == REVISION_DATE for value in target_modified.values()),
        "non_target_entries": len(after_non_target),
        "non_target_preserved": before_non_target == after_non_target,
        "url_order_preserved": before_urls == after_urls,
        "exact_scoped_rewrite": candidate == exact_expected,
        "hub_entries_detected": len(hub_pairs),
        "before_sha256": sha256_bytes(baseline),
        "after_sha256": sha256_bytes(candidate),
    }


def public_observation(item: dict[str, object]) -> dict[str, object]:
    hidden = {"paragraphs", "headings", "normalized_text", "shingles"}
    return {key: value for key, value in item.items() if key not in hidden}


def representative_samples(
    observations: list[dict[str, object]],
    actual_hashes: dict[str, str] | None,
    actual_state: str,
) -> list[dict[str, object]]:
    valid = [item for item in observations if not item.get("invalid")]
    by_key = {(str(item["profile"]), str(item["locality"])): item for item in valid}
    reasons: dict[str, set[str]] = defaultdict(set)
    for profile, locality in REPRESENTATIVE_TARGETS:
        item = by_key.get((profile, locality))
        if item:
            reasons[str(item["path"])].add("representative")
    for profile, _, _ in PROFILE_SPECS:
        members = sorted(
            (item for item in valid if item["profile"] == profile),
            key=lambda item: int(item["source_row"]),
        )
        if members:
            reasons[str(members[0]["path"])].add(f"{profile}_first_row")
            reasons[str(members[-1]["path"])].add(f"{profile}_last_row")
    numeric_boundaries = (
        ("source_chars", "source_length"),
        ("current_chars", "current_length"),
        ("new_chars", "new_length"),
        ("source_new_recall", "source_use"),
        ("current_new_jaccard", "current_similarity"),
    )
    for field, label in numeric_boundaries:
        if valid:
            lowest = min(valid, key=lambda item: float(item[field]))
            highest = max(valid, key=lambda item: float(item[field]))
            reasons[str(lowest["path"])].add(f"min_{label}")
            reasons[str(highest["path"])].add(f"max_{label}")
    corrected = [item for item in valid if item["corrected_localities"]]
    for item in corrected[:8]:
        reasons[str(item["path"])].add("source_contamination_boundary")
    by_path = {str(item["path"]): item for item in valid}
    samples: list[dict[str, object]] = []
    for path in sorted(reasons):
        item = public_observation(by_path[path])
        actual_sha = (actual_hashes or {}).get(path, "")
        item["sample_reasons"] = sorted(reasons[path])
        item["projected"] = {
            "sha256": item["new_sha256"],
            "fact_parity": item["fact_parity"],
        }
        item["actual"] = {
            "state": actual_state,
            "sha256": actual_sha or None,
            "matches_projected": bool(actual_sha) and actual_sha == item["new_sha256"],
        }
        samples.append(item)
    return samples


def provenance_manifest(observations: list[dict[str, object]]) -> str:
    rows = []
    for item in sorted(observations, key=lambda value: str(value.get("path", ""))):
        rows.append("\t".join((
            str(item.get("path", "")), str(item.get("profile", "")), str(item.get("locality", "")),
            str(item.get("source_workbook", "")), str(item.get("source_row", "")),
            str(item.get("source_sha256", "")), str(item.get("current_sha256", "")),
            str(item.get("current_raw_worktree_sha256", "")), str(item.get("new_sha256", "")),
        )))
    return digest_rows(rows)


def expected_detail_paths() -> list[Path]:
    paths: list[Path] = []
    for _, _, category in PROFILE_SPECS:
        members = sorted((SUBJECT_ROOT / category).glob("*/index.html"))
        paths.extend(members)
    return paths


def run_self_test() -> dict[str, object]:
    marker = (
        '<section class="section manuscript-wrap" data-revision="composite-2026-08-22" '
        'data-source-row="2" data-source-sha256="abc"><article class="site-shell manuscript-article">'
        '<div class="manuscript-intro"><span>안내</span><p>현재 자료를 기준으로 학습 순서를 확인합니다.</p></div>'
        '<section class="manuscript-section"><h2>첫 점검 기준</h2>'
        '<p>최근 교재와 오답 기록을 함께 살펴 다음 확인 날짜를 정합니다.</p></section>'
        '<section class="manuscript-section"><h2>복습 기록 방법</h2>'
        '<p>풀이 과정을 다시 설명하고 혼자 재현할 수 있는지 확인합니다.</p></section>'
        '<section class="manuscript-section"><h2>상담 준비 자료</h2>'
        '<p>학생이 실제 사용하는 자료를 준비해 현재 단원부터 대조합니다.</p></section>'
        '</article></section>'
    )
    parsed = manuscript_from_html(marker)
    fake_workbooks: dict[str, list[str]] = {}
    fake_documents: list[SimpleNamespace] = []
    fake_localities: dict[str, list[str]] = {}
    for profile, workbook, category in PROFILE_SPECS:
        localities = [f"{profile}지역{index:03d}" for index in range(EXPECTED_DETAILS_PER_PROFILE)]
        raws = [f"<h1>{locality}</h1><p>{profile} source row {index + 2}</p>" for index, locality in enumerate(localities)]
        fake_workbooks[workbook] = raws
        fake_localities[profile] = localities
        for locality, raw in reversed(list(zip(localities, raws))):
            fake_documents.append(SimpleNamespace(
                profile=profile, locality=locality,
                path=SUBJECT_ROOT / category / locality / "index.html",
                source_sha256=sha256_bytes(raw.encode("utf-8")),
            ))

    class FakeGenerator:
        @staticmethod
        def read_archive(_archive):
            return fake_workbooks, {}

        @staticmethod
        def source_fragment(raw):
            return "제목", visible_text(raw), [("기준", [visible_text(raw)])]

        @staticmethod
        def load_base(profile, _common):
            return SimpleNamespace(load_centers=lambda: {value: {} for value in fake_localities[profile]})

    join_findings = Findings()
    joined, _ = source_records(
        FakeGenerator, SimpleNamespace(documents=fake_documents), Path("ignored"), Path("ignored"), join_findings,
    )
    checks = {
        "marker_attributes": attributes(parsed.opening_tag).get("data-source-row") == "2",
        "manuscript_intro": parsed.intro.startswith("현재 자료"),
        "manuscript_sections": parsed.headings == ("첫 점검 기준", "복습 기록 방법", "상담 준비 자료"),
        "promotional_unsafe": bool(promotional_hits("성적을 끌어올립니다.")[0]),
        "promotional_softened": not promotional_hits("성적 상승을 단정하는 표현보다 기록을 확인합니다.")[0],
        "guarantee_disclaimer": not promotional_hits("성적을 보장한다는 말보다 기록을 확인해야 합니다.")[0],
        "implicit_outcome_claim": (
            bool(promotional_hits("실력이 쌓이도록 관리합니다.")[0])
            and not promotional_hits("실력을 쌓는 과정과 연습 계획이 중요합니다.")[0]
            and bool(promotional_hits("성적 하락을 방지합니다.")[0])
        ),
        "outcome_high_recall": all(promotional_hits(value)[0] for value in (
            "학습 결과가 점수 향상으로 이어지도록 관리합니다.",
            "실력이 성적으로 연결되도록 지도합니다.",
            "성적을 안정적으로 유지하도록 관리합니다.",
            "점수 변화를 만듭니다.",
            "학업 역량을 강화하는 데 활용합니다.",
            "성취도가 향상되도록 코칭합니다.",
            "성적을 성장시킵니다.",
            "실력으로 연결합니다.",
            "실력을 만들어요.",
            "성과가 나오도록 관리합니다.",
            "효과를 극대화합니다.",
            "학습 변화를 유지합니다.",
            "이해력을 높이고 향상시킵니다.",
        )),
        "outcome_nominal_safe": all(not promotional_hits(value)[0] for value in (
            "실력 향상을 위한 계획을 설명합니다.",
            "점수와 오답을 연결하여 확인합니다.",
            "학업 역량 강화 방법을 질문합니다.",
            "점수 변화의 근거를 확인합니다.",
            "실력과 학습 과정이 연결되는 구조를 확인합니다.",
            "목표 성적 달성을 위한 전략을 정리합니다.",
            "독해력 훈련 강화 방법을 질문합니다.",
            "성장 목표와 방향을 상담에서 확인합니다.",
            "실력을 쌓아가고 싶다면 현재 기록부터 봅니다.",
            "꾸준한 성장을 목표로 운영합니다.",
            "지속 성장 로드맵을 제공합니다.",
            "꾸준히 성장하는 방향을 잡습니다.",
            "내신과 실력 향상에 필요한 과정이 꾸준히 쌓이도록 기록합니다.",
            "성취도만 보지 않고 필요한 학습 순서를 먼저 만듭니다.",
            "문장 이해력을 통해 국어 학습의 바탕을 함께 만듭니다.",
            "학업 역량을 기준으로 다음 계획을 만듭니다.",
            "성취 목표(상위권/성장)를 상담에서 정리합니다.",
        )),
        "outcome_mixed_sentence": bool(promotional_hits(
            "성취도 향상을 위한 계획을 세우고 성취도를 높입니다."
        )[0]),
        "outcome_affective_direct": all(promotional_hits(value)[0] for value in (
            "학습 자신감을 키웁니다.",
            "실전 감각을 키웁니다.",
            "이해력과 표현력을 함께 키웁니다.",
            "성취감을 쌓습니다.",
            "성취 경험을 쌓습니다.",
        )),
        "score_conversion_direct": all(promotional_hits(value)[0] for value in (
            "오답을 다음 점수로 바꿉니다.",
            "오답을 ‘다음 점수’로 바꾸는 피드백.",
            "오답 노트 기반으로 약점을 ‘다음 시험 점수’로 바꾸는 관리가 진행됩니다.",
            "오답이 다음 점수로 바뀝니다.",
            "약점이 시험 점수로 바뀌도록 관리합니다.",
        )),
        "score_conversion_nominal_safe": (
            not promotional_hits("오답을 점수로 바꾸는 방법을 질문합니다.")[0]
            and not promotional_hits("오답이 곧 점수로 바뀐다는 표현은 피해야 합니다.")[0]
        ),
        "extended_outcome_high_recall": all(promotional_hits(value)[0] for value in (
            "오답을 성적으로 바꾸는 관리가 진행됩니다.",
            "꾸준함을 실력으로 전환합니다.",
            "학습 변화가 성과로 나오도록 관리합니다.",
            "성적이 따라오도록 지도합니다.",
            "결과가 실전에 반영되도록 관리합니다.",
            "실력이 회복되도록 돕습니다.",
            "성과가 목표에 도달하도록 관리합니다.",
            "점수 변화로 이어지도록 관리합니다.",
            "오답을 다음 성과로 연결합니다.",
            "점수 변동을 최소화합니다.",
            "당일 점수 변화를 최소화합니다.",
            "스스로 공부하는 힘을 길러줍니다.",
            "기본기를 다집니다.",
            "자신감을 회복하도록 관리합니다.",
            "문해력을 보완합니다.",
            "성장 루틴을 만듭니다.",
            "학습 변화를 꾸준히 만듭니다.",
            "실력을 잡습니다.",
            "실력을 다집니다.",
            "성취 경험을 유지합니다.",
            "점수 변동을 줄입니다.",
            "개념 누락을 최소화합니다.",
            "개념 누락을 줄이고 기본 독해 습관을 형성합니다.",
            "개념 누락과 유형별 오답 원인을 분류해 재발을 막습니다.",
            "개념 누락과 풀이 실수를 분석해 재발 방지 루틴으로 정착시킵니다.",
        )),
        "extended_outcome_safe": all(not promotional_hits(value)[0] for value in (
            "학습 결과가 다음 계획에 반영되는지 확인합니다.",
            "결과를 바탕으로 학습 계획을 만듭니다.",
            "실력을 확인하고 보완 순서를 잡습니다.",
            "개념 누락을 줄여 학습 변화를 목표로 합니다.",
            "개념 누락을 줄이도록 유형을 반복 점검합니다.",
            "개념 누락의 재발 방지 전략까지 함께 세웁니다.",
        )),
        "punctuation_balance": (
            punctuation_balance_issues("‘다음 점수’ (확인 기준)") == ()
            and punctuation_balance_issues("학습 변화’으로") == ("single_quote:0:1",)
            and punctuation_balance_issues("영수(영수") == ("parenthesis:1:0",)
        ),
        "token_similarity": token_metrics("교재 오답 기록", "최근 교재와 오답 기록", set())[1] >= .66,
        "administrative_locality": administrative_mentions(
            "인천 남동구 논현동", "논현동"
        ) == {"인천광역시 남동구 논현동"},
        "administrative_subset": administrative_compatible(
            "충청남도 천안시 신방동", "충청남도 천안시 동남구 신방동", "신방동"
        ),
        "locality_boundaries": (
            locality_mentions("자료를 비교하여 대구 유천동을 봅니다", "대구유천동", {"교하", "유천동"}) == []
            and locality_mentions("광명동 자료가 섞였습니다", "개운동", {"광명동"}) == ["광명동"]
            and locality_mentions("불당동 자료를 확인합니다", "신불당", {"불당동"}) == []
            and locality_mentions("운정 자료를 확인합니다", "야당동", {"운정"}) == []
            and locality_mentions("별내동 자료를 확인합니다", "별내중앙", {"별내동"}) == []
        ),
        "embedded_target_corruption": (
            embedded_target_corruptions("삼삼송동 자료", "삼송") == ["삼삼송"]
            and embedded_target_corruptions("정진월동게 확인", "진월동") == ["정진월동게"]
            and embedded_target_corruptions("비전주 장동기보다", "전주 장동") == ["비전주 장동기보다"]
        ),
        "registration_context": (
            not UNSUPPORTED_OPERATIONAL_PATTERNS["registration"].search("주소 제304호")
            and bool(UNSUPPORTED_OPERATIONAL_PATTERNS["registration"].search("교육지원청 제 123호"))
        ),
        "awkward_families": (
            bool(AWKWARD_PATTERNS["duplicate_level_fit"].search("목표에 맞춘 현재 수준에 맞춘 계획"))
            and bool(AWKWARD_PATTERNS["provided_internal_exam"].search("학교 내신 제공된 시험지의 문항 유형"))
            and bool(AWKWARD_PATTERNS["provided_exam"].search("시험 제공된 시험지의 문항 유형"))
            and bool(AWKWARD_PATTERNS["subject_column_duplicate"].search("국어와 국어·영어·수학"))
            and bool(AWKWARD_PATTERNS["subject_particle"].search("국어·영어·수학를 확인"))
            and bool(AWKWARD_PATTERNS["object_rewrite_collision"].search("계획을 설계 기준을 정리"))
            and not AWKWARD_PATTERNS["object_rewrite_collision"].search(
                "과정을 학생 설명과 맞춰 보고 보완 순서를 정합니다"
            )
            and bool(AWKWARD_PATTERNS["elementary_token_corruption"].search("별교과 평가도시"))
            and bool(AWKWARD_PATTERNS["missing_outcome_object"].search(
                "학습 과정에서는 오답 기록을 바탕으로 변화를 확인하는 기준으로 삼습니다"
            ))
            and bool(AWKWARD_PATTERNS["double_topic_consultation"].search(
                "실제 자료를 확인할 때는 학생 상담은 기록부터 봅니다"
            ))
            and bool(AWKWARD_PATTERNS["stacked_lead_in"].search(
                "과목별 기록을 대조하면 최근 교재와 오답을 대조하면, 기준이 보입니다"
            ))
            and bool(AWKWARD_PATTERNS["duplicate_current_level_fit"].search(
                "현재 수준에 맞춘 계획을 현재 수준에 맞춘 자료로 확인합니다"
            ))
            and bool(AWKWARD_PATTERNS["duplicate_dot_token"].search("국어·국어 학습"))
            and bool(AWKWARD_PATTERNS["duplicate_compound_suffix"].search("서술·서술형 풀이"))
            and bool(AWKWARD_PATTERNS["outcome_service_fragment"].search("국어·영수 실력 연결. 다음 문장"))
        ),
        "awkward_block_scope": (
            bool(AWKWARD_PATTERNS["duplicate_current_level_fit"].search(
                "현재 수준에 맞춘 계획 현재 수준에 맞춘 상담"
            ))
            and all(
                not AWKWARD_PATTERNS["duplicate_current_level_fit"].search(block)
                for block in ("현재 수준에 맞춘 계획", "현재 수준에 맞춘 상담")
            )
        ),
        "locality_key_join": (
            join_findings.total == 0
            and len(joined) == EXPECTED_DETAIL_DOCUMENTS
            and joined[("elementary", "elementary지역000")].row == 2
            and joined[("high", "high지역370")].row == 372
        ),
        "sitemap_parser": sitemap_pairs(
            ('<?xml version="1.0" encoding="UTF-8"?>\n'
             '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
             '<url><loc>https://example.test/</loc><lastmod>2026-08-22</lastmod></url>'
             '</urlset>').encode("utf-8")
        ) == [("https://example.test/", "2026-08-22")],
    }
    failed = sorted(key for key, value in checks.items() if not value)
    return {"status": "PASS" if not failed else "FAIL", "read_only": True, "checks": checks, "failed": failed}


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only projected/actual content audit for the 1,113 revised K/E/M detail pages."
    )
    parser.add_argument("--mode", choices=("projected", "actual", "both"), default="both")
    parser.add_argument("--generator", type=Path, default=DEFAULT_GENERATOR)
    parser.add_argument("--archive", type=Path, default=DEFAULT_ARCHIVE)
    parser.add_argument("--common-dir", type=Path, default=DEFAULT_COMMON)
    parser.add_argument(
        "--generator-pin", default=EXPECTED_GENERATOR_SHA256,
        help="PENDING during final generator freeze, or the required 64-character SHA-256 after freeze.",
    )
    parser.add_argument("--skip-repeat", action="store_true", help="Skip the second deterministic projection.")
    parser.add_argument("--full-provenance", action="store_true", help="Include all 1,113 provenance rows in JSON.")
    parser.add_argument("--self-test", action="store_true", help="Run parser/claim/similarity unit checks without loading the generator.")
    args = parser.parse_args()

    if args.self_test:
        result = run_self_test()
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["status"] == "PASS" else 1

    generator_path = args.generator.resolve()
    archive = args.archive.resolve()
    common_dir = args.common_dir.resolve()
    for label, path in (("generator", generator_path), ("archive", archive)):
        if not path.is_file():
            raise AuditFailure(f"missing {label}: {path}")
    if not common_dir.is_dir():
        raise AuditFailure(f"missing common data directory: {common_dir}")

    detail_paths = expected_detail_paths()
    if len(detail_paths) != EXPECTED_DETAIL_DOCUMENTS:
        raise AuditFailure(f"working-tree detail path count={len(detail_paths)} expected={EXPECTED_DETAIL_DOCUMENTS}")
    product_paths = [*detail_paths, ROOT / "sitemap.xml"]
    mutation_before = file_manifest(product_paths)
    generator_sha_before = sha256_file(generator_path)
    generator = load_module(generator_path, "_audit_revised_kem_generator")
    if Path(generator.ROOT).resolve() != ROOT.resolve():
        raise AuditFailure(f"generator root mismatch: {generator.ROOT}")

    contract_findings = Findings()
    if clean(getattr(generator, "BASE_COMMIT", "")) != EXPECTED_BASE_COMMIT:
        contract_findings.add("generator_base_commit", repr(getattr(generator, "BASE_COMMIT", None)))
    if clean(getattr(generator, "REVISION_DATE", "")) != REVISION_DATE:
        contract_findings.add("generator_revision_date", repr(getattr(generator, "REVISION_DATE", None)))
    if int(getattr(generator, "EXPECTED_DETAIL_DOCUMENTS", -1)) != EXPECTED_DETAIL_DOCUMENTS:
        contract_findings.add("generator_detail_contract", repr(getattr(generator, "EXPECTED_DETAIL_DOCUMENTS", None)))
    if int(getattr(generator, "EXPECTED_DOCUMENT_COUNT", -1)) != EXPECTED_PRODUCT_DOCUMENTS:
        contract_findings.add("generator_product_contract", repr(getattr(generator, "EXPECTED_DOCUMENT_COUNT", None)))
    if args.generator_pin == "PENDING":
        contract_findings.warn("generator_pin_pending", generator_sha_before)
    elif not re.fullmatch(r"[0-9a-f]{64}", args.generator_pin):
        contract_findings.add("generator_pin_format", repr(args.generator_pin))
    elif generator_sha_before != args.generator_pin:
        contract_findings.add("generator_pin_mismatch", f"{generator_sha_before} != {args.generator_pin}")

    started = time.perf_counter()
    plan = generator.build_plan(archive, common_dir, run_profile_preflight=False)
    build_seconds = round(time.perf_counter() - started, 3)
    generator_sha_after_build = sha256_file(generator_path)
    if generator_sha_after_build != generator_sha_before:
        contract_findings.add(
            "generator_changed_during_projection", f"{generator_sha_before} -> {generator_sha_after_build}",
        )
    if plan.candidate_sha256() != EXPECTED_CANDIDATE_SHA256:
        contract_findings.add(
            "generator_candidate_pin_mismatch",
            f"{plan.candidate_sha256()} != {EXPECTED_CANDIDATE_SHA256}",
        )
    observed_generator_after_manifest = clean(plan.metrics.get("after_manifest", ""))
    if observed_generator_after_manifest != EXPECTED_GENERATOR_AFTER_MANIFEST_SHA256:
        contract_findings.add(
            "generator_after_manifest_pin_mismatch",
            f"{observed_generator_after_manifest} != {EXPECTED_GENERATOR_AFTER_MANIFEST_SHA256}",
        )
    observed_generator_before_manifest = clean(plan.metrics.get("before_manifest", ""))
    observed_generator_changed = int(plan.metrics.get("changed", -1))

    plan_paths = [Path(doc.path).resolve() for doc in plan.documents]
    expected_product_set = {path.resolve() for path in product_paths}
    if len(plan.documents) != EXPECTED_PRODUCT_DOCUMENTS or len(set(plan_paths)) != EXPECTED_PRODUCT_DOCUMENTS:
        contract_findings.add("product_plan_count", f"documents={len(plan.documents)}, unique={len(set(plan_paths))}")
    if set(plan_paths) != expected_product_set:
        contract_findings.add(
            "product_plan_scope",
            f"missing={len(expected_product_set-set(plan_paths))}, extra={len(set(plan_paths)-expected_product_set)}",
        )
    if any(path != (ROOT / "sitemap.xml").resolve() and not path.is_relative_to(SUBJECT_ROOT.resolve()) for path in plan_paths):
        contract_findings.add("product_plan_escape", "plan contains a path outside the authorized product scope")

    baseline_rels = [relative(path) for path in product_paths]
    baseline_blobs = git_blobs(EXPECTED_BASE_COMMIT, baseline_rels)
    candidate_by_rel = {relative(Path(doc.path)): bytes(doc.after) for doc in plan.documents}
    raw_before_by_rel = {relative(Path(doc.path)): bytes(doc.before) for doc in plan.documents}
    raw_pending_documents = sum(raw_before_by_rel[rel] != candidate_by_rel[rel] for rel in candidate_by_rel)
    known_superseded_state = (
        observed_generator_before_manifest == KNOWN_SUPERSEDED_AFTER_MANIFEST_SHA256
        and observed_generator_changed == EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS
        and raw_pending_documents == EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS
    )
    allowed_pending_counts = {0, EXPECTED_PRODUCT_DOCUMENTS}
    if known_superseded_state:
        allowed_pending_counts.add(EXPECTED_SUPERSEDED_CHANGED_DOCUMENTS)
    if raw_pending_documents not in allowed_pending_counts:
        contract_findings.add(
            "raw_product_partial_materialization",
            f"pending={raw_pending_documents}, expected one of {sorted(allowed_pending_counts)}",
        )
    baseline_changed = sum(candidate_by_rel.get(rel) != value for rel, value in baseline_blobs.items())
    if baseline_changed != EXPECTED_PRODUCT_DOCUMENTS:
        contract_findings.add(
            "projected_baseline_change_count", f"{baseline_changed} != {EXPECTED_PRODUCT_DOCUMENTS}",
        )

    repeat_payload: dict[str, object] = {"checked": not args.skip_repeat}
    if not args.skip_repeat:
        repeat_started = time.perf_counter()
        repeat = generator.build_plan(archive, common_dir, run_profile_preflight=False)
        repeat_documents = {relative(Path(doc.path)): bytes(doc.after) for doc in repeat.documents}
        first_candidate = plan.candidate_sha256()
        second_candidate = repeat.candidate_sha256()
        deterministic = candidate_by_rel == repeat_documents and first_candidate == second_candidate
        if not deterministic:
            contract_findings.add("projection_not_deterministic", f"{first_candidate} != {second_candidate}")
        repeat_payload.update({
            "deterministic": deterministic,
            "candidate_sha256": second_candidate,
            "seconds": round(time.perf_counter() - repeat_started, 3),
        })

    projection_findings = Findings()
    sources, locality_sets = source_records(generator, plan, archive, common_dir, projection_findings)
    all_localities = set().union(*locality_sets.values()) if locality_sets else set()
    observations: list[dict[str, object]] = []
    detail_documents = [doc for doc in plan.documents if getattr(doc, "profile", "") in locality_sets]
    for doc in detail_documents:
        key = (doc.profile, clean(doc.locality))
        record = sources.get(key)
        rel = relative(Path(doc.path))
        if record is None:
            projection_findings.add("source_mapping_missing", f"{doc.profile}/{doc.locality}")
            continue
        observations.append(audit_page(
            "projected", Path(doc.path), doc.profile, record.locality,
            baseline_blobs[rel], bytes(doc.before), bytes(doc.after), record,
            locality_sets[doc.profile], projection_findings,
        ))
    corpus_metrics = audit_corpus(observations, all_localities, projection_findings)
    target_urls = {
        page_url_from_canonical(candidate_by_rel[relative(Path(doc.path))].decode("utf-8"))
        for doc in detail_documents
    }
    if len(target_urls) != EXPECTED_DETAIL_DOCUMENTS:
        projection_findings.add("canonical_target_count", f"{len(target_urls)}")
    projected_sitemap = candidate_by_rel.get("sitemap.xml", b"")
    sitemap_metrics = audit_sitemap(
        "projected", baseline_blobs["sitemap.xml"], raw_before_by_rel["sitemap.xml"],
        projected_sitemap, target_urls, projection_findings,
    )

    release_findings = Findings()
    script_presence = {relative(path): path.is_file() for path in RELEASE_SCRIPTS}
    for path, present in script_presence.items():
        if not present:
            release_findings.add("release_script_missing", path)
    release_scope_paths = set(plan_paths) | {path.resolve() for path in RELEASE_SCRIPTS}
    if len(release_scope_paths) != EXPECTED_RELEASE_SCOPE:
        release_findings.add("release_scope_count", f"{len(release_scope_paths)} != {EXPECTED_RELEASE_SCOPE}")

    actual_findings = Findings()
    working_detail_bytes = {relative(path): path.read_bytes() for path in detail_paths}
    marker_token = f'data-revision="{REVISION_MARKER}"'.encode("utf-8")
    marker_count = sum(marker_token in value for value in working_detail_bytes.values())
    actual_hashes: dict[str, str] = {}
    actual_observations: list[dict[str, object]] = []
    actual_sitemap_metrics: dict[str, object] = {}
    if marker_count == 0:
        actual_state = "PENDING"
        try:
            pairs = sitemap_pairs((ROOT / "sitemap.xml").read_bytes())
            baseline_pair_map = dict(sitemap_pairs(baseline_blobs["sitemap.xml"]))
            actual_pair_map = dict(pairs)
            changed_target_lastmods = {
                url: (baseline_pair_map.get(url, ""), actual_pair_map.get(url, ""))
                for url in target_urls
                if baseline_pair_map.get(url, "") != actual_pair_map.get(url, "")
            }
            actual_sitemap_metrics = {
                "observed_detail_lastmods": dict(Counter(
                    modified for url, modified in pairs if url in target_urls
                )),
                "changed_without_page_markers": len(changed_target_lastmods),
            }
            if changed_target_lastmods:
                actual_state = "PARTIAL"
                actual_findings.add(
                    "actual_sitemap_only_partial_application",
                    f"changed detail lastmods={len(changed_target_lastmods)}",
                )
        except (UnicodeDecodeError, ET.ParseError) as exc:
            actual_findings.add("actual_sitemap_parse", str(exc))
        if args.mode == "actual":
            actual_findings.add("actual_not_applied", "no revised detail markers are present")
    elif marker_count != EXPECTED_DETAIL_DOCUMENTS:
        actual_state = "PARTIAL"
        actual_findings.add(
            "actual_partial_application", f"markers={marker_count}/{EXPECTED_DETAIL_DOCUMENTS}",
        )
    else:
        actual_state = "APPLIED"
        actual_product_bytes = dict(working_detail_bytes)
        actual_product_bytes["sitemap.xml"] = (ROOT / "sitemap.xml").read_bytes()
        mismatches = [
            rel for rel in sorted(candidate_by_rel)
            if actual_product_bytes.get(rel) != candidate_by_rel[rel]
        ]
        for rel in mismatches[:8]:
            actual_findings.add("actual_projected_byte_mismatch", rel)
        if len(mismatches) > 8:
            actual_findings.counts["actual_projected_byte_mismatch"] += len(mismatches) - 8
        for doc in detail_documents:
            rel = relative(Path(doc.path))
            record = sources[(doc.profile, clean(doc.locality))]
            actual_observations.append(audit_page(
                "actual", Path(doc.path), doc.profile, record.locality,
                baseline_blobs[rel], bytes(doc.before), actual_product_bytes[rel], record,
                locality_sets[doc.profile], actual_findings,
            ))
            actual_hashes[rel] = sha256_bytes(actual_product_bytes[rel])
        actual_sitemap_metrics = audit_sitemap(
            "actual", baseline_blobs["sitemap.xml"], raw_before_by_rel["sitemap.xml"],
            actual_product_bytes["sitemap.xml"],
            target_urls, actual_findings,
        )

    mutation_after = file_manifest(product_paths)
    if mutation_before != mutation_after:
        contract_findings.add(
            "read_only_product_mutation",
            f"changed during audit={sorted(path for path in set(mutation_before)|set(mutation_after) if mutation_before.get(path)!=mutation_after.get(path))[:8]}",
        )
    generator_sha_final = sha256_file(generator_path)
    if generator_sha_final != generator_sha_before:
        contract_findings.add(
            "generator_changed_during_audit", f"{generator_sha_before} -> {generator_sha_final}",
        )

    projection_required = args.mode in {"projected", "both"}
    actual_required = args.mode == "actual" or (args.mode == "both" and actual_state != "PENDING")
    error_total = contract_findings.total + release_findings.total
    if projection_required:
        error_total += projection_findings.total
    if actual_required or actual_state == "PARTIAL":
        error_total += actual_findings.total
    status = "PASS" if error_total == 0 else "HOLD"
    public_provenance = [public_observation(item) for item in observations]
    result: dict[str, object] = {
        "schema_version": 1,
        "status": status,
        "mode": args.mode,
        "read_only": True,
        "root": str(ROOT),
        "contract": {
            "detail_pages": EXPECTED_DETAIL_DOCUMENTS,
            "product_documents": EXPECTED_PRODUCT_DOCUMENTS,
            "release_scripts": [relative(path) for path in RELEASE_SCRIPTS],
            "release_scope": EXPECTED_RELEASE_SCOPE,
            "revision_date": REVISION_DATE,
            "base_commit": EXPECTED_BASE_COMMIT,
            "generator_pin": {
                "required": args.generator_pin,
                "state": "PENDING" if args.generator_pin == "PENDING" else "FROZEN",
                "observed_sha256": generator_sha_before,
            },
            "candidate_pin": {
                "required": EXPECTED_CANDIDATE_SHA256,
                "observed_sha256": plan.candidate_sha256(),
            },
            "generator_after_manifest_pin": {
                "required": EXPECTED_GENERATOR_AFTER_MANIFEST_SHA256,
                "observed_sha256": observed_generator_after_manifest,
            },
            "working_product_state": {
                "state": (
                    "CURRENT_CANDIDATE" if raw_pending_documents == 0 else
                    "KNOWN_SUPERSEDED_RELEASE" if known_superseded_state else
                    "VERIFIED_BASELINE_PENDING" if raw_pending_documents == EXPECTED_PRODUCT_DOCUMENTS else
                    "UNRECOGNIZED_PARTIAL"
                ),
                "observed_before_manifest": observed_generator_before_manifest,
                "known_superseded_manifest": KNOWN_SUPERSEDED_AFTER_MANIFEST_SHA256,
                "generator_changed_documents": observed_generator_changed,
                "raw_pending_documents": raw_pending_documents,
            },
            "findings": contract_findings.payload(),
        },
        "release_scope": {
            "product_documents": len(plan.documents),
            "script_documents": len(RELEASE_SCRIPTS),
            "total": len(release_scope_paths),
            "script_presence": script_presence,
            "findings": release_findings.payload(),
        },
        "projected": {
            "status": "PASS" if projection_findings.total == 0 else "HOLD",
            "build_seconds": build_seconds,
            "generator_candidate_sha256": plan.candidate_sha256(),
            "product_documents": len(plan.documents),
            "detail_documents": len(detail_documents),
            "raw_worktree_pending_documents": raw_pending_documents,
            "changed_from_verified_git_baseline": baseline_changed,
            "raw_before_manifest_sha256": digest_rows(
                f"{rel}\t{sha256_bytes(value)}" for rel, value in sorted(raw_before_by_rel.items())
            ),
            "verified_git_baseline_manifest_sha256": digest_rows(
                f"{rel}\t{sha256_bytes(value)}" for rel, value in sorted(baseline_blobs.items())
            ),
            "candidate_after_manifest_sha256": digest_rows(
                f"{rel}\t{sha256_bytes(value)}" for rel, value in sorted(candidate_by_rel.items())
            ),
            "generator_metrics": plan.metrics,
            "source_manifest": plan.source_manifest,
            "repeat": repeat_payload,
            "provenance_manifest_sha256": provenance_manifest(observations),
            "corpus": corpus_metrics,
            "sitemap": sitemap_metrics,
            "findings": projection_findings.payload(),
        },
        "actual": {
            "state": actual_state,
            "marker_pages": marker_count,
            "projected_byte_matches": (
                sum(actual_hashes.get(rel) == item.get("new_sha256") for rel, item in ((str(obs.get("path")), obs) for obs in observations))
                if actual_state == "APPLIED" else 0
            ),
            "sitemap": actual_sitemap_metrics,
            "findings": actual_findings.payload(),
        },
        "samples": representative_samples(observations, actual_hashes, actual_state),
    }
    if args.full_provenance:
        result["provenance"] = public_provenance
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if status == "PASS" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AuditFailure, OSError, subprocess.SubprocessError, RuntimeError, ValueError, AttributeError) as exc:
        print(json.dumps({
            "status": "FAIL",
            "read_only": True,
            "error": f"{type(exc).__name__}: {exc}",
        }, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(2)
