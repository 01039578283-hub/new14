from __future__ import annotations

from typing import Final, TypedDict


class SubjectCategory(TypedDict):
    order: int
    slug: str
    label: str
    english: str
    description: str


# This tuple is the single display order used by the subject root hub and feeds.
SUBJECT_CATALOG: Final[tuple[SubjectCategory, ...]] = (
    {
        "order": 1,
        "slug": "고등학생국영수학원",
        "label": "고등학생 국영수학원",
        "english": "High school",
        "description": "학교 내신 범위와 모의고사 오답, 수행평가 일정을 함께 살펴 세 과목의 우선순위를 정합니다.",
    },
    {
        "order": 2,
        "slug": "중학생국영수학원",
        "label": "중학생 국영수학원",
        "english": "Middle school",
        "description": "교과 개념의 빈틈, 학교 시험 준비와 과제 실행을 구분해 중등 학습의 주간 흐름을 확인합니다.",
    },
    {
        "order": 3,
        "slug": "초등학생국영수학원",
        "label": "초등학생 국영수학원",
        "english": "Elementary school",
        "description": "읽기 이해와 어휘, 계산 과정과 학습 시작 습관을 중심으로 초등 교과 기초를 점검합니다.",
    },
    {
        "order": 4,
        "slug": "영수학원",
        "label": "영수학원",
        "english": "English & Math",
        "description": "영어와 수학의 현재 교재와 오답 기록을 함께 살펴 두 과목의 학습 우선순위와 주간 실행 계획을 정합니다.",
    },
    {
        "order": 5,
        "slug": "초등영어학원",
        "label": "초등 영어학원",
        "english": "Elementary English",
        "description": "읽기·듣기·어휘의 기초와 학습 습관을 살펴 초등 단계에 맞는 영어 수업과 복습 흐름을 확인합니다.",
    },
    {
        "order": 6,
        "slug": "초등수학학원",
        "label": "초등 수학학원",
        "english": "Elementary Math",
        "description": "수 개념과 연산 정확도, 풀이 과정을 살펴 초등 수학의 기초와 교과 진도에 맞는 학습 흐름을 확인합니다.",
    },
    {
        "order": 7,
        "slug": "중등영어학원",
        "label": "중등 영어학원",
        "english": "Middle English",
        "description": "어휘·문법·독해의 연결과 학교 시험 준비를 살펴 중등 영어의 개념 보완과 복습 계획을 확인합니다.",
    },
    {
        "order": 8,
        "slug": "중등수학학원",
        "label": "중등 수학학원",
        "english": "Middle Math",
        "description": "교과 개념과 문제 풀이 과정, 오답 복습을 살펴 중등 수학의 내신 대비와 주간 학습 계획을 확인합니다.",
    },
    {
        "order": 9,
        "slug": "고등영어학원",
        "label": "고등 영어학원",
        "english": "High English",
        "description": "학교별 내신 범위와 모의고사 독해·어법 오답을 살펴 고등 영어의 학습 우선순위와 대비 계획을 정합니다.",
    },
    {
        "order": 10,
        "slug": "고등수학학원",
        "label": "고등 수학학원",
        "english": "High Math",
        "description": "교과 개념과 학교별 내신 범위, 모의고사 오답을 살펴 고등 수학의 단원별 보완 순서와 학습 계획을 정합니다.",
    },
)
