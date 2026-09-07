"""Site14: literal manuscript foci first; evidenced learning topics for KEM."""
from __future__ import annotations
from collections import Counter,defaultdict
import math,re

CORE_GROUPS = {"고등수학학원","고등영어학원","중등수학학원","중등영어학원","초등수학학원","초등영어학원"}
CORE_FOCUS = re.compile(r"핵심 주제는 .+?입니다\. (?P<focus>.+?)(?:의 (?:현재선|출발 상태|출발점)|에서 현재선)")
OPERATIONAL = re.compile(r"교습비|등록번호|주소|주차|차량|센터명|이용 조건|모집 중|합격 보장")
HUB_SUFFIXES = {
 "고등수학학원":("개념 통합·조건 해석·서술형 점검",["개념 통합·조건 해석·서술형"]),
 "고등영어학원":("어휘 누적·구문·독해 근거 점검",["어휘 누적·구문·독해 근거"]),
 "고등학생국영수학원":("내신·모의고사·수행평가 우선순위",["학교 내신 범위","모의고사 오답","수행평가 일정","우선순위"]),
 "영수학원":("과목별 막힘과 주간 학습 균형",["과목별 병목·주간 균형"]),
 "중등수학학원":("개념 연결·조건 해석·오답 점검",["개념 연결·조건 해석·오답"]),
 "중등영어학원":("어휘·문법 적용·독해 근거 점검",["어휘·문법 적용·독해 근거"]),
 "중학생국영수학원":("개념·학교 시험과 과제·오답 복습",["교과 개념","학교 시험 준비","과제 수행","오답 복습"]),
 "초등수학학원":("연산 원리·개념 설명·문장제 점검",["연산 원리·개념 설명·문장제"]),
 "초등영어학원":("어휘·문장 읽기·기초 표현 점검",["어휘·문장 읽기·기초 표현"]),
 "초등학생국영수학원":("읽기·어휘·계산과 공부 시작 습관",["읽기 이해","어휘 활용","계산 과정","공부를 시작하는 습관"]),
}
ENGLISH_SIGNALS = {
 "어휘 회상":"어휘 회상","문장 구조":"영어 문장 구조","문법 적용":"문법 적용",
 "독해 근거":"독해 근거","듣기 확인":"영어 듣기 확인","서술형 쓰기":"영어 서술형 쓰기",
 "영어 오답 유형":"영어 오답 유형","시험 시간 배분":"영어 시험 시간 배분","영어 학습 루틴":"영어 학습 루틴",
}
MATH_SIGNALS = {
 "개념 연결":"수학 개념 연결","계산 정확도":"계산 정확도","문제 조건 해석":"수학 문제 조건 해석",
 "수학 오답 재학습":"수학 오답 재학습","서술형 풀이":"수학 서술형 풀이","그래프와 자료 해석":"그래프와 자료 해석",
 "단원 간 연결":"수학 단원 간 연결","수학 시간 배분":"수학 시간 배분","수학 학습 루틴":"수학 학습 루틴",
}
PRIMARY_PAIRED_CUES=set(ENGLISH_SIGNALS)|set(MATH_SIGNALS)

# The same explicit diagnostic questions as the existing paired-subject body.
ENGLISH_SIGNALS.update({
    "뜻을 외운 단어를 문장 안에서 다시 알아보는지": ENGLISH_SIGNALS["어휘 회상"],
    "주어와 동사, 수식 범위를 구분해 문장을 읽는지": ENGLISH_SIGNALS["문장 구조"],
    "규칙을 아는 문제와 실제 문장에 적용하지 못한 문제를 나눴는지": ENGLISH_SIGNALS["문법 적용"],
    "답을 고른 문장과 선택지의 표현을 연결하는지": ENGLISH_SIGNALS["독해 근거"],
    "놓친 소리와 뜻을 모르는 표현을 구분했는지": ENGLISH_SIGNALS["듣기 확인"],
    "조건에 맞는 문장을 스스로 구성하고 검토하는지": ENGLISH_SIGNALS["서술형 쓰기"],
    "어휘·문법·독해 중 어디에서 같은 실수가 이어지는지": ENGLISH_SIGNALS["영어 오답 유형"],
    "어느 지문과 문항에서 시간이 길어지는지": ENGLISH_SIGNALS["시험 시간 배분"],
    "계획한 분량과 실제 완료한 분량이 구분되는지": ENGLISH_SIGNALS["영어 학습 루틴"],
})
MATH_SIGNALS.update({
    "정의와 공식을 말로 설명한 뒤 문제에 적용하는지": MATH_SIGNALS["개념 연결"],
    "부호·괄호·연산 순서에서 같은 실수가 반복되는지": MATH_SIGNALS["계산 정확도"],
    "문장에서 필요한 조건을 골라 식이나 그림으로 바꾸는지": MATH_SIGNALS["문제 조건 해석"],
    "틀린 이유를 개념·조건·계산으로 나누어 기록하는지": MATH_SIGNALS["수학 오답 재학습"],
    "답뿐 아니라 식을 세운 이유와 풀이 순서를 적는지": MATH_SIGNALS["서술형 풀이"],
    "그림·표·그래프의 정보를 식과 연결하는지": MATH_SIGNALS["그래프와 자료 해석"],
    "앞 단원의 개념이 현재 문제에서 어떻게 쓰이는지 설명하는지": MATH_SIGNALS["단원 간 연결"],
    "어느 유형과 계산 단계에서 시간이 길어지는지": MATH_SIGNALS["수학 시간 배분"],
    "계획한 문제와 실제 완료한 문제를 구분하는지": MATH_SIGNALS["수학 학습 루틴"],
})
# Exact activities from the existing paired-subject manuscripts; these allow
# their substantive secondary learning tasks to outweigh shared introductions.
ENGLISH_SIGNALS.update({
    "단어 시험과 독해 지문에서 같은 어휘를 놓친 기록": ENGLISH_SIGNALS["어휘 회상"],
    "뜻·품사·예문을 한 묶음으로 확인하기": ENGLISH_SIGNALS["어휘 회상"],
    "하루 뒤 예문 속 단어 뜻을 다시 말해 보기": ENGLISH_SIGNALS["어휘 회상"],
    "긴 문장에서 해석이 끊긴 위치와 표시한 문장 성분": ENGLISH_SIGNALS["문장 구조"],
    "문장 뼈대를 먼저 적고 수식어를 붙여 읽기": ENGLISH_SIGNALS["문장 구조"],
    "한 문장을 짧게 끊어 소리 내어 설명하기": ENGLISH_SIGNALS["문장 구조"],
    "고른 답의 근거를 문법 규칙으로 설명한 흔적": ENGLISH_SIGNALS["문법 적용"],
    "규칙 한 줄 뒤에 맞는 예문과 틀린 예문을 함께 적기": ENGLISH_SIGNALS["문법 적용"],
    "틀린 문장을 고치고 이유를 한 문장으로 남기기": ENGLISH_SIGNALS["문법 적용"],
    "정답 근거에 밑줄을 긋고 선택지를 지운 이유를 적은 기록": ENGLISH_SIGNALS["독해 근거"],
    "문단별 핵심 문장과 연결어를 표시하며 읽기": ENGLISH_SIGNALS["독해 근거"],
    "짧은 지문 한 편에서 근거 문장만 다시 찾기": ENGLISH_SIGNALS["독해 근거"],
    "받아쓰기와 다시 들은 뒤 수정한 부분": ENGLISH_SIGNALS["듣기 확인"],
    "짧은 구간을 듣고 핵심 표현을 따라 말하기": ENGLISH_SIGNALS["듣기 확인"],
    "한 문장을 듣고 의미 단위로 끊어 적기": ENGLISH_SIGNALS["듣기 확인"],
    "서술형 답안에서 빠진 조건과 고친 문장": ENGLISH_SIGNALS["서술형 쓰기"],
    "핵심 표현을 넣어 짧은 답안을 완성하기": ENGLISH_SIGNALS["서술형 쓰기"],
    "답안을 다시 써 보고 빠진 조건을 표시하기": ENGLISH_SIGNALS["서술형 쓰기"],
    "첫 풀이와 다시 푼 답 사이에 달라진 근거": ENGLISH_SIGNALS["영어 오답 유형"],
    "오답을 유형별로 나누고 다시 볼 날짜 정하기": ENGLISH_SIGNALS["영어 오답 유형"],
    "같은 유형 한 문제를 며칠 뒤 다시 풀기": ENGLISH_SIGNALS["영어 오답 유형"],
    "문항별 소요 시간과 끝까지 풀지 못한 구간": ENGLISH_SIGNALS["시험 시간 배분"],
    "읽기와 답 확인 시간을 나누어 연습하기": ENGLISH_SIGNALS["시험 시간 배분"],
    "짧은 세트를 정해진 시간 안에 풀고 기록하기": ENGLISH_SIGNALS["시험 시간 배분"],
    "학습 시작 시각과 완료 여부를 적은 주간 기록": ENGLISH_SIGNALS["영어 학습 루틴"],
    "매일 할 최소 분량과 다시 볼 항목을 나누기": ENGLISH_SIGNALS["영어 학습 루틴"],
    "완료한 분량과 남은 질문을 짧게 적기": ENGLISH_SIGNALS["영어 학습 루틴"],
})
MATH_SIGNALS.update({
    "기본 문제와 조건이 달라진 문제의 풀이 근거": MATH_SIGNALS["개념 연결"],
    "개념 한 줄과 대표 문제를 한 묶음으로 정리하기": MATH_SIGNALS["개념 연결"],
    "해설 없이 개념을 설명하고 한 문제를 다시 풀기": MATH_SIGNALS["개념 연결"],
    "중간 계산과 답을 고친 뒤의 재풀이 기록": MATH_SIGNALS["계산 정확도"],
    "계산 단계를 한 줄씩 나누고 틀린 위치 표시하기": MATH_SIGNALS["계산 정확도"],
    "짧은 계산 세트를 정확도 기준으로 다시 풀기": MATH_SIGNALS["계산 정확도"],
    "조건을 빠뜨린 문제와 식을 잘못 세운 문제의 구분": MATH_SIGNALS["문제 조건 해석"],
    "조건에 밑줄을 긋고 구하려는 값을 먼저 적기": MATH_SIGNALS["문제 조건 해석"],
    "비슷한 문제에서 식을 세우는 과정만 다시 연습하기": MATH_SIGNALS["문제 조건 해석"],
    "첫 풀이와 다시 푼 풀이에서 달라진 단계": MATH_SIGNALS["수학 오답 재학습"],
    "오답 원인을 표시하고 다시 볼 날짜 정하기": MATH_SIGNALS["수학 오답 재학습"],
    "같은 유형 한 문제를 며칠 뒤 다시 풀기": MATH_SIGNALS["수학 오답 재학습"],
    "서술형 답안에서 생략된 조건과 고친 표현": MATH_SIGNALS["서술형 풀이"],
    "풀이 단계마다 근거를 짧은 문장으로 남기기": MATH_SIGNALS["서술형 풀이"],
    "완성한 풀이를 소리 내어 설명해 보기": MATH_SIGNALS["서술형 풀이"],
    "자료를 잘못 읽은 부분과 다시 표시한 값": MATH_SIGNALS["그래프와 자료 해석"],
    "주어진 정보를 그림에 옮기고 관계를 설명하기": MATH_SIGNALS["그래프와 자료 해석"],
    "표나 그림을 보고 조건을 한 문장으로 적기": MATH_SIGNALS["그래프와 자료 해석"],
    "최근 문제에서 다시 필요해진 이전 단원 기록": MATH_SIGNALS["단원 간 연결"],
    "현재 단원과 연결된 앞 개념을 짧게 복습하기": MATH_SIGNALS["단원 간 연결"],
    "오늘 문제에 쓰인 이전 개념을 한 줄로 적기": MATH_SIGNALS["단원 간 연결"],
    "문항별 소요 시간과 끝까지 풀지 못한 구간": MATH_SIGNALS["수학 시간 배분"],
    "풀이와 검산 시간을 나누어 연습하기": MATH_SIGNALS["수학 시간 배분"],
    "짧은 세트를 풀고 오래 걸린 문제를 표시하기": MATH_SIGNALS["수학 시간 배분"],
    "학습 시작 시각과 재풀이 완료 여부를 적은 기록": MATH_SIGNALS["수학 학습 루틴"],
    "매일 풀 분량과 다시 볼 문제를 나누기": MATH_SIGNALS["수학 학습 루틴"],
    "완료한 문제와 남은 질문을 짧게 적기": MATH_SIGNALS["수학 학습 루틴"],
})
# label, sentence-level literal/close-paraphrase cue, family, eligible stages.
# No business claims, promised results, review anecdotes, location hashes or
# randomly assigned phrases enter this catalogue.
KEM_TOPICS = [
 ("과목별 학습 균형",r"세 과목.{0,18}(균형|분량)|국어.{0,8}영어.{0,8}수학.{0,20}균형","balance","all"),
 ("과목별 완료 기준",r"과목별 완료 기준|과목별.{0,12}완료.{0,8}기준","planning","all"),
 ("과목별 보완 순서",r"과목별 (보완 순서|약점)|취약 과목.{0,12}(먼저|순서)","priority","all"),
 ("오답 원인 분류",r"오답 원인.{0,25}(구분|분류)|왜 틀렸는지.{0,35}분류","error","all"),
 ("개념 공백 확인",r"개념 공백|개념의 빈틈|단원의 공백|개념 결손","concept","all"),
 ("선행 전 기초 점검",r"선행.{0,25}(기초|현재|오답)|기초.{0,25}선행","concept","all"),
 ("스스로 풀이 설명",r"혼자 다시 풀|스스로.{0,15}설명|학생이.{0,15}설명할 수","independent","all"),
 ("복습 시점 정하기",r"복습 시점|재확인 날짜|복습 주기|복습 간격","review","all"),
 ("과제 이행 기록",r"과제 이행 기록|과제 수행 기록|과제 완료","homework","all"),
 ("풀이 과정 기록",r"풀이 과정을 기록|풀이 과정을 남|풀이 순서를 남|풀이 노트","process","all"),
 ("계획과 실제 공부 비교",r"계획.{0,15}실제|실제 학습 시간.{0,20}(계획|연결)|학습 시간.{0,15}계획","execution","all"),
 ("실행 가능한 공부 분량",r"실행 가능한 분량|감당할 수 있는 분량|공부량.{0,20}(조절|조정)","planning","all"),
 ("작은 목표부터 시작",r"작은 목표|짧은 목표|작은 성공|작은 성취","confidence","all"),
 ("공부 시작 습관",r"공부를 시작|숙제를 시작|시작, 풀이, 채점, 질문|시작 시각","start","all"),
 ("질문 미루는 습관 점검",r"질문을 미룬|질문을 미루|질문을 못|질문을 하지","question","all"),
 ("배운 내용 다시 설명",r"배운 내용.{0,25}설명|개념을 말로 설명|핵심 개념.{0,15}설명","concept","all"),
 ("국어 지문 근거 찾기",r"국어.{0,35}근거 찾|국어 지문 표시","korean-reading","all"),
 ("국어 지문 해석",r"국어.{0,25}지문 해석","korean-reading","all"),
 ("영어 단어 누적 복습",r"영어 단어 누적|영어.{0,20}어휘.{0,15}(누적|복습)|영어[^.!?]{0,70}단어[^.!?]{0,35}뜻을 다시","english-vocab","all"),
 ("수학 문장제 조건 읽기",r"수학 문장제.{0,25}(조건|흔들|이해)|문장제.{0,15}조건","math-condition","all"),
 ("수학 계산 과정 확인",r"계산 과정|계산.{0,12}(실수|오류)|계산량보다 개념","math-calculation","all"),
 ("영어 문장 이해",r"영어.{0,18}문장 이해|영어 문장 구조","english-sentence","all"),
 ("국어·영어·수학 연결",r"국어.{0,35}영어.{0,35}수학.{0,20}(같이|함께|연결|영향)","cross-subject","all"),
 ("학교 일정과 복습 조율",r"학교 일정.{0,25}(복습|분량)|학교 시험.{0,25}복습|시험 일정.{0,25}(복습|계획)","schedule","all"),
 ("개념·유형·오답 연결",r"개념[-·/]유형[-·/]오답|개념.{0,8}유형.{0,8}오답","math-link","all"),
 ("내신·모의고사 점검",r"내신.{0,15}모의고사|학교 시험.{0,15}모의고사","exam-types","high"),
 ("수능·기출 준비",r"수능.{0,15}기출|기출.{0,15}수능","suneung","high"),
 ("수행평가 준비 순서",r"수행평가.{0,30}(일정|준비|순서|조건)","performance","middle high"),
 ("서술형 답안 확인",r"서술형.{0,20}(답안|풀이|표현)|서술형 준비","written","middle high"),
 ("시간 배분 점검",r"시간 배분|문항별 시간|시험 시간.{0,15}(부족|조절)","pace","middle high"),
 ("시험 뒤 계획 조정",r"시험.{0,15}(뒤|후).{0,20}(계획|조정|재설계)","exam-review","middle high"),
 ("교과 개념과 내신 연결",r"교과 개념.{0,25}(내신|시험)|개념.{0,15}내신","school-concept","middle high"),
 ("중학교 학습 전환 준비",r"중학교 (진학|진입|적응)|중등.{0,12}전환|중학교.{0,12}(학습|과정).{0,12}(준비|적응|전환)","transition","elementary middle"),
 ("단원평가와 기초 이해",r"단원평가|단원 성취도|현재 학년 개념","school","elementary"),
 ("읽기 이해와 어휘 활용",r"읽기 이해.{0,18}어휘|읽기.{0,15}어휘 활용","reading-vocab","elementary"),
 ("집중과 휴식 간격",r"집중.{0,25}휴식|짧게.{0,20}집중|집중 시간","attention","all"),
 ("학습 부담 원인 살피기",r"학습 부담|부담.{0,20}(난도|분량|과제)|난이도.{0,20}부담","load","all"),
 ("피드백을 다음 풀이에 반영",r"피드백.{0,20}(다음|반영)|다음 풀이에.{0,10}반영","feedback","all"),
 ("정답보다 풀이 근거",r"정답.{0,20}(근거|과정)|답만.{0,20}(과정|근거)","process","all"),
 ("오답 수정 뒤 재확인",r"다시 풀기 결과|며칠 뒤.{0,20}(확인|풀)|다시 푼.{0,20}(비교|결과)","error","all"),
 ("질문·풀이·채점 순서",r"질문과 풀이 과정|풀이, 채점, 질문|풀이와 질문","question","all"),
 ("학습 기록으로 목표 조정",r"기록.{0,25}(목표|조정)|복습 이력.{0,20}계획","feedback","all"),
]

def checked_focus(page):
    intro=[b["text"] for b in page["blocks"] if b["role"]=="intro"]
    if len(intro)!=1:
        raise ValueError("Expected one manuscript introduction: "+page["rel"])
    m=CORE_FOCUS.search(intro[0])
    if not m:
        raise ValueError("Explicit learning-focus clause missing: "+page["rel"])
    raw=m["focus"]
    label=raw.replace("오답 원인을 다시 살핀 기록","오답 원인·재확인")+" 점검"
    if not 4<=len(label)<=60 or OPERATIONAL.search(label) or "|" in label:
        raise ValueError("Not a concise learning focus: "+page["rel"])
    if page["subject"]=="math" and re.search(r"영어|어휘|영작|파닉스|문법",label):
        raise ValueError("Math focus contains an English topic")
    if page["subject"]=="english" and re.search(r"수학|연산|분수|방정식|도형",label):
        raise ValueError("English focus contains a math topic")
    if page["stage"]=="elementary" and re.search(r"고등|수능|모의고사",label):
        raise ValueError("Elementary focus contains an unsupported upper-level topic")
    return label,[dict(label=label,match=raw,excerpt=intro[0],role="explicit-learning-focus",subject=page["subject"])],"explicit-learning-focus"

def combined_candidates(page):
    candidates=[]
    ambiguous=set(ENGLISH_SIGNALS).intersection(MATH_SIGNALS)
    for subject,signals in (("english",ENGLISH_SIGNALS),("math",MATH_SIGNALS)):
        grouped=defaultdict(list)
        for index,block in enumerate(page["blocks"]):
            if block["role"] not in {"summary","prose"}:
                continue
            for cue,label in signals.items():
                if cue not in ambiguous and cue in block["text"]:
                    grouped[label].append((index,cue,block["text"]))
        for label,hits in grouped.items():
            cues=sorted({h[1] for h in hits})
            if len(cues)<2 and not PRIMARY_PAIRED_CUES.intersection(cues):
                continue
            index,cue,text=min(hits,key=lambda h:(-len(h[1]),h[0]))
            candidates.append(dict(label=label,match=cue,excerpt=text,role="paired-subject-focus",subject=subject,
                                   signals=cues,paragraphs=len({h[0] for h in hits}),firstPosition=min(h[0] for h in hits)))
    return candidates

def combined_focus(page,rarity=None):
    candidates=combined_candidates(page)
    selected=[]
    for subject in ("english","math"):
        choices=[c for c in candidates if c["subject"]==subject]
        for choice in choices:
            choice["score"]=round((rarity or {}).get(choice["label"],1)*(1+0.2*min(len(choice["signals"]),4)),6)
        if not choices:
            raise ValueError("Fewer than two unambiguous learning signals for "+subject+": "+page["rel"])
        selected.append(min(choices,key=lambda c:(-c["score"],c["firstPosition"],c["label"])))
    return "·".join(e["label"] for e in selected)+" 점검",selected,"paired-subject-focus"

def kem_candidates(page):
    choices=[]
    for label,pattern,family,stages in KEM_TOPICS:
        if stages!="all" and page["stage"] not in stages.split():
            continue
        hits=[]
        for index,block in enumerate(page["blocks"]):
            # Local summary includes school/center facts. Candidate evidence is
            # confined to learning prose inside the manuscript article.
            if block["role"] not in {"intro","prose","heading"}:
                continue
            match=re.search(pattern,block["text"])
            if match:
                hits.append((index,block,match))
        if hits:
            index,block,match=min(hits,key=lambda x:(x[1]["role"]!="heading",x[0]))
            choices.append(dict(label=label,match=match[0],excerpt=block["text"],role="manuscript-topic",
                                subject="combined",family=family,paragraphs=len(hits),firstPosition=index))
    return choices

def corpus_rarity(pages):
    groups=defaultdict(list)
    for page in pages:
        if ("국영수" in page["group"] or page["group"]=="영수학원") and page["kind"]!="category":
            groups[page["group"]].append(page)
    result={}
    for group,members in groups.items():
        counts=Counter()
        for page in members:
            candidates=combined_candidates(page) if group=="영수학원" else kem_candidates(page)
            counts.update(x["label"] for x in candidates)
        result[group]={label:1+math.log((len(members)+1)/(count+1)) for label,count in counts.items()}
    return result

def select_detail(page,rarity=None):
    if page["group"] in CORE_GROUPS:
        return checked_focus(page)
    if page["group"]=="영수학원":
        return combined_focus(page,rarity)
    if "국영수" not in page["group"]:
        raise ValueError("Unsupported category: "+page["group"])
    candidates=kem_candidates(page)
    for candidate in candidates:
        # Use actual within-page support and document rarity; never random
        # rotation or a requirement that a topic be unique to one locality.
        candidate["score"]=round((rarity or {}).get(candidate["label"],1)*(1+0.15*min(candidate["paragraphs"],3)),6)
    candidates.sort(key=lambda c:(-c["score"],c["firstPosition"],c["label"]))
    selected=[]
    for candidate in candidates:
        if any(e["family"]==candidate["family"] for e in selected):
            continue
        selected.append(candidate)
        if len(selected)==2:
            break
    if len(selected)!=2:
        raise ValueError("Fewer than two supported learning topics: "+page["rel"])
    return "·".join(e["label"] for e in selected),selected,"manuscript-topics"
