"""Offline contracts for manuscript-based titles and preserved page bytes."""
import unittest
from pathlib import Path
from unittest.mock import patch
import importlib.util

from personalize_title_suffixes import (
    ROOT, DOMAIN, META_RE, attrs, title_of, replace_titles, masked,
    body_blocks, read_page, plan, inventory, resolved_page_url, update_rss,
)
from title_suffix_rules import (
    CORE_GROUPS, HUB_SUFFIXES, checked_focus, select_detail, combined_candidates,
    kem_candidates, corpus_rarity,
)


def page(group="고등수학학원",text="핵심 주제는 ‘학습 기준’입니다. 서술형 풀이와 개념 연결의 출발 상태를 확인합니다.",role="intro"):
    return dict(rel="test",kind="subject",group=group,
                subject="combined" if "영수" in group else "english" if "영어" in group else "math",
                stage="elementary" if group.startswith("초") else "middle" if group.startswith("중") else "high",
                blocks=[dict(text=text,role=role)])


class FocusTests(unittest.TestCase):
    def test_explicit_focus(self):
        title,evidence,mode=select_detail(page())
        self.assertEqual(title,"서술형 풀이와 개념 연결 점검")
        self.assertEqual(evidence[0]["match"],"서술형 풀이와 개념 연결")
        self.assertEqual(mode,"explicit-learning-focus")
    def test_no_topic_hash(self):
        left=page();right=page();right["rel"]="다른동네"
        self.assertEqual(select_detail(left),select_detail(right))
    def test_elementary_start_phrase(self):
        title,_,_=select_detail(page("초등수학학원","핵심 주제는 ‘출발점’입니다. 풀이 과정·서술과 개념 설명·적용의 출발점을 확인합니다."))
        self.assertEqual(title,"풀이 과정·서술과 개념 설명·적용 점검")
    def test_elementary_naturalized_error_phrase(self):
        title,proof,_=select_detail(page("초등영어학원","핵심 주제는 ‘기초’입니다. 듣기·말하기와 오답 원인을 다시 살핀 기록에서 현재선을 확인합니다."))
        self.assertEqual(title,"듣기·말하기와 오답 원인·재확인 점검")
        self.assertIn("기록",proof[0]["match"])
    def test_cross_subject_rejected(self):
        for group,topic in [("고등수학학원","문법과 어휘"),("고등영어학원","분수와 연산")]:
            with self.subTest(group=group),self.assertRaises(ValueError):
                select_detail(page(group,f"핵심 주제는 ‘기준’입니다. {topic}의 현재선을 봅니다."))
    def test_elementary_high_level_rejected(self):
        with self.assertRaises(ValueError):
            select_detail(page("초등영어학원","핵심 주제는 ‘기초’입니다. 수능·모의고사의 현재선을 봅니다."))
    def test_middle_transition_is_explicit(self):
        title,_,_=select_detail(page("중등영어학원","핵심 주제는 ‘다음 단계’입니다. 고등 영어 전환 준비와 문장 구조의 현재선을 봅니다."))
        self.assertIn("고등 영어 전환 준비",title)
    def test_unrecognized_intro_does_not_fabricate(self):
        with self.assertRaises(ValueError):select_detail(page(text="소개만 있습니다."))
    def test_operational_not_learning(self):
        with self.assertRaises(ValueError):select_detail(page(text="핵심 주제는 ‘안내’입니다. 교습비와 주소의 현재선을 봅니다."))
    def test_paired_subjects(self):
        p=page("영수학원","영어 오답 유형과 개념 연결을 확인합니다.","summary")
        title,proof,_=select_detail(p)
        self.assertEqual([x["subject"] for x in proof],["english","math"])
        self.assertEqual(title,"영어 오답 유형·수학 개념 연결 점검")
    def test_unambiguous_secondary_activities(self):
        p=page("영수학원","영어 오답 유형과 개념 연결을 봅니다.","summary")
        p["blocks"].extend([dict(role="prose",text="받아쓰기와 다시 들은 뒤 수정한 부분을 봅니다."),dict(role="prose",text="짧은 구간을 듣고 핵심 표현을 따라 말하기를 연습합니다.")])
        title,proof,_=select_detail(p,{"영어 듣기 확인":4})
        self.assertIn("영어 듣기 확인",title)
        self.assertEqual(len(proof[0]["signals"]),2)
    def test_shared_signal_cannot_identify_subject(self):
        p=page("영수학원","같은 유형 한 문제를 며칠 뒤 다시 풀기","prose")
        self.assertEqual(combined_candidates(p),[])
        with self.assertRaises(ValueError):select_detail(p)
    def test_kem_stage_filter(self):
        p=page("초등학생국영수학원","내신과 모의고사, 수능과 기출, 복습 시점과 과제 이행 기록을 확인합니다.","prose")
        candidates=kem_candidates(p)
        self.assertFalse(any("수능" in c["label"] or "모의고사" in c["label"] for c in candidates))
    def test_korean_words_are_not_english(self):
        p=page("중학생국영수학원","국어 단어 뜻을 다시 말합니다.","prose")
        self.assertFalse(any(c["label"]=="영어 단어 누적 복습" for c in kem_candidates(p)))
    def test_kem_two_distinct_families(self):
        p=page("고등학생국영수학원","오답 원인을 분류하고 다시 풀기 결과를 봅니다. 복습 시점을 정합니다.","prose")
        _,proof,_=select_detail(p)
        self.assertEqual(len({c["family"] for c in proof}),2)
    def test_rarity_deterministic(self):
        p=page("고등학생국영수학원","복습 시점과 과제 이행 기록을 봅니다.","prose")
        q=page("고등학생국영수학원","복습 시점을 정하고 국어 지문 해석을 확인합니다.","prose")
        self.assertEqual(corpus_rarity([p,q]),corpus_rarity([q,p]))


class HtmlTests(unittest.TestCase):
    def test_only_learning_areas(self):
        source='''<main><nav><p>navigation</p></nav><div class="local-summary"><p>learning summary</p></div><div class="local-info-card"><p>center address</p></div><article class="manuscript-article"><div class="manuscript-intro"><p>intro<img src="x"> rest</p></div><section class="manuscript-section"><h2>heading</h2><p>proof <strong>word</strong></p><p style="display:none">secret</p><div hidden><p>hidden</p></div></section></article><section class="faq"><p>faq</p></section></main>'''
        self.assertEqual(body_blocks(source,"subject"),[dict(text="learning summary",role="summary"),dict(text="intro rest",role="intro"),dict(text="heading",role="heading"),dict(text="proof word",role="prose")])
    def test_metadata_only_and_escaping(self):
        source='''<title>A | B</title><meta property="og:title" content="A | B"><meta name='twitter:title' content='A | B'><meta name="description" content="keep"><link rel="canonical" href="/keep/"><script type="application/ld+json">{"name":"keep"}</script><main><h1>A</h1><img alt="keep" style="display:none;"></main>'''
        title="A | X & Y ' Z"
        updated=replace_titles(source,title)
        self.assertEqual(masked(source),masked(updated))
        self.assertEqual(title_of(updated),title)
        self.assertEqual(replace_titles(updated,title),updated)
        self.assertEqual([attrs(m[0])["content"] for m in META_RE.finditer(updated) if attrs(m[0]).get("property",attrs(m[0]).get("name")) in {"og:title","twitter:title"}],[title,title])
    def test_missing_duplicate_title_rejected(self):
        for source in ("<h1>A</h1>","<title>A</title><title>B</title>"):
            with self.assertRaises(ValueError):replace_titles(source,"A | B")
    def test_canonical_route_guard(self):
        rel=Path("과목별학원/고등수학학원/명일동/index.html")
        self.assertIn("%",resolved_page_url("/과목별학원/고등수학학원/명일동/",rel))
        with self.assertRaises(ValueError):resolved_page_url("https://example.com/",rel)
    def test_rss_only_matching_item(self):
        source='<rss><channel><item><title>old</title><link>'+DOMAIN+'/a/</link><description>keep</description></item><item><title>protected</title><link>'+DOMAIN+'/b/</link></item></channel></rss>'
        updated,count=update_rss([dict(url=DOMAIN+'/a/',after='new & title')],source)
        self.assertEqual(count,1)
        self.assertIn('<title>new &amp; title</title>',updated)
        self.assertIn('<title>protected</title>',updated)
        self.assertIn('<description>keep</description>',updated)
        self.assertEqual(update_rss([dict(url=DOMAIN+'/a/',after='new & title')],updated),(updated,0))
    def test_all_real_hubs_have_evidence(self):
        for group in HUB_SUFFIXES:
            with self.subTest(group=group):
                p=read_page(ROOT/'과목별학원'/group/'index.html')
                plan([p])
                self.assertEqual(masked(p['source']),masked(p['updated']))
                self.assertTrue(all(e['match'] in e['excerpt'] for e in p['evidence']))
    def test_real_core_examples(self):
        for group in CORE_GROUPS:
            with self.subTest(group=group):
                p=read_page(ROOT/'과목별학원'/group/'명일동/index.html')
                plan([p]);first=p['after'];plan([p])
                self.assertEqual(first,p['after'])
                self.assertEqual(masked(p['source']),masked(p['updated']))
    def test_problematic_paired_example(self):
        p=read_page(ROOT/'과목별학원/영수학원/도남지구/index.html')
        _,proof,_=select_detail(p)
        self.assertEqual(len(proof),2)
    def test_inventory(self):
        targets,protected=inventory()
        self.assertEqual((len(targets),len(protected)),(3720,6))
    def test_legacy_entry_uses_personalizer(self):
        spec=importlib.util.spec_from_file_location('legacy_titles',ROOT/'scripts/update_title_suffixes.py')
        module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
        with patch('sys.argv',['update_title_suffixes.py']),patch('subprocess.call',return_value=0) as call,self.assertRaises(SystemExit):
            module.main()
        self.assertIn('personalize_title_suffixes.py',call.call_args[0][0][1])
        self.assertEqual(call.call_args[0][0][-1],'--write')


if __name__=='__main__':unittest.main(verbosity=2)
