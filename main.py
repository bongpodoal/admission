import re
from pathlib import Path
import pandas as pd
from pip._internal.cli import req_command

EXCEL_FILE = Path(__file__).resolve().with_name(
    "전국대학_정시입결_공식통합_v1.xlsx"
)
MAJOR_SEARCH_GROUPS = {

    "컴퓨터": [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터",
        "정보"
    ],

    "소프트웨어": [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터",
        "정보"
    ],

    "인공지능": [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터"
    ],

    "기계": [
        "기계",
        "자동차",
        "로봇",
        "메카트로닉스"
    ],

    "전자": [
        "전자",
        "전기",
        "반도체",
        "정보통신"
    ],

    "경영": [
        "경영",
        "회계",
        "금융",
        "무역"
    ],

    "경제": [
        "경제",
        "금융"
    ],

    "화학": [
        "화학",
        "화학공학",
        "신소재"
    ],

    "생명": [
        "생명",
        "생물",
        "바이오"
    ],

    "교육": [
        "교육"
    ]
}

SCIENCE_INQUIRY_KEYWORDS = [
    "물리",
    "화학",
    "생명",
    "지구"
]

SOCIAL_INQUIRY_KEYWORDS = [
    "생활과윤리",
    "윤리와사상",
    "한국지리",
    "세계지리",
    "동아시아사",
    "세계사",
    "경제",
    "정치와법",
    "사회문화",
]
MAJOR_WEIGHTS = {

    "공학": {
        "국어": 0.20,
        "수학": 0.35,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "자연과학": {
        "국어": 0.20,
        "수학": 0.35,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "의약보건": {
        "국어": 0.20,
        "수학": 0.35,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "인문사회": {
        "국어": 0.35,
        "수학": 0.20,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "상경": {
        "국어": 0.30,
        "수학": 0.25,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "교육": {
        "국어": 0.30,
        "수학": 0.25,
        "탐구1": 0.225,
        "탐구2": 0.225
    },

    "기타": {
        "국어": 0.25,
        "수학": 0.25,
        "탐구1": 0.25,
        "탐구2": 0.25
    }
}


def normalize_text(text):
    if text is None:
        return ""

    text = str(text).lower()

    text = re.sub(
        r"[^가-힣a-z0-9]",
        "",
        text
    )

    return text


def classify_inquiry_subject(subject):
    subject = normalize_text(subject)

    for keyword in SCIENCE_INQUIRY_KEYWORDS:
        if normalize_text(keyword) in subject:
            return "과탐"

    for keyword in SOCIAL_INQUIRY_KEYWORDS:
        if normalize_text(keyword) in subject:
            return "사탐"

    return "기타"


def infer_major_group(major_name):
    name = normalize_text(major_name)

    medical_keywords = [
        "의예",
        "의학",
        "치의",
        "치과",
        "한의",
        "약학",
        "수의",
        "간호",
        "물리치료",
        "작업치료"
    ]

    for keyword in medical_keywords:
        if normalize_text(keyword) in name:
            return "의학보건"

    engineering_keywords = [
        "컴퓨터",
        "소프트웨어",
        "인공지능",
        "ai",
        "데이터",
        "전자",
        "전기",
        "기계",
        "로봇",
        "자동차",
        "메카트로닉스",
        "반도체",
        "정보통신",
        "통신",
        "토목",
        "건축공학",
        "산업공학",
        "신소재",
        "애너지",
        "환경공학",
        "화학공학",
        "조선",
        "항공"
    ]

    for keyword in engineering_keywords:
        if normalize_text(keyword) in name:
            return "공학"

    natural_keywords = [
        "수학",
        "통계",
        "물리",
        "화학",
        "생명",
        "생물",
        "지구",
        "천문",
        "환경과학",
        "식품",
        "농학"
    ]

    for keyword in natural_keywords:
        if normalize_text(keyword) in name:
            return "자연과학"

    business_keyword = [
        "경영",
        "경제",
        "회계",
        "세무",
        "금융",
        "무역",
        "국제통상"
    ]

    for keyword in business_keyword:
        if normalize_text(keyword) in name:
            return "상경"

    education_keyword = [
        "교육",
        "초등교육",
        "유아교육"
    ]

    for keyword in education_keyword:
        if normalize_text(keyword) in name:
            return "교육"

    humanities_keyword = [
        "국어",
        "영어",
        "문학",
        "사학",
        "역사",
        "철학",
        "사회",
        "심리",
        "행정",
        "정치",
        "외교",
        "법",
        "미디어",
        "언론",
        "광고",
        "문헌정보",
        "문화",
        "관광"
    ]

    for keyword in humanities_keyword:
        if normalize_text(keyword) in name:
            return "인문사회"

    return "기타"


def get_major_keywords(
        desired_major
):
    if not desired_major:
        return []

    desired = normalize_text(
        desired_major
    )

    for key, keywords in (
            MAJOR_SEARCH_GROUPS.items()
    ):

        normalized_key = (
            normalize_text(key)
        )

        if (
                normalized_key
                in desired
                or desired
                in normalized_key
        ):
            return [
                normalize_text(keyword)
                for keyword
                in keywords
            ]

    # 등록되지 않은 학과
    return [desired]


def input_percentile(subject):
    while True:
        try:
            score = float(
                input(f"{subject} 백분위 입력 (0 ~ 100): ")
            )

            if 0 <= score <= 100:
                return score

            print("0 ~ 100사이의 숫자를 입력하세요.")

        except ValueError:
            print("숫자를 입력하세요.")


def input_grade(subject):
    while True:
        try:
            grade = int(
                input(f"{subject} 등급 입력 (1~9): ")
            )

            if 1 <= grade <= 9:
                return grade

            print("1~9 사이의 등급을 입력하세요")
        except ValueError:
            print("정수를 입력하세요")


def input_student_scores():
    print()
    print("===== 모의고사 성적 입력 =====")

    inquiry1_name = input(
        "탐구1 과목병: "
    ).strip()

    inquiry2_name = input(
        "탐구2 과목병: "
    ).strip()

    korean = input_percentile("국어")
    math = input_percentile("수학")

    inquiry1 = input_percentile(f"{inquiry1_name}")
    inquiry2 = input_percentile(f"{inquiry2_name}")

    english = input_grade("영어")
    history = input_grade("한국사")

    print()
    print("[희망 학과]")
    print("희망 학과가 없다면 Enter를 누르세요")

    desired_major = input(
        "희망 학과: "
    ).strip()

    if desired_major == "":
        desired_major = None

    return {
        "국어": korean,
        "수학": math,
        "탐구1": inquiry1,
        "탐구2": inquiry2,
        "영어": english,
        "한국사": history,
        "희망학과": desired_major
    }


def calculate_weighted_score(
        values,
        weight
):
    total_score = 0
    total_weight = 0

    for subject, weight in weight.items():
        value = values.get(subject)

        if value is None:
            continue

        if pd.isna(value):
            continue

        total_score += float(value) * weight

        total_weight += weight

        if total_weight == 0:
            return None

        return total_score / total_weight


def calculate_university_cutoff(
        row,
        weights
):
    university_score = {
        "국어": row.get(
            "국어70"
        ),
        "수학": row.get(
            "수학70"
        ),
        "탐구1": row.get(
            "탐구1_70"
        ),
        "탐구2": row.get(
            "탐구2_70"
        )
    }

    return calculate_weighted_score(university_score, weights)


def check_inquiry_fit(scores, major_group):
    inquiry1_type = classify_inquiry_subject(scores["탐구1과목"])

    inquiry2_type = classify_inquiry_subject(scores["탐구1과목"])

    types = [inquiry1_type, inquiry2_type]

    science_count = (
        types.count("과탐")
    )

    social_count = (
        types.count("사탐")
    )

    if major_group in ["공학", "자연과학", "의약보건"]:
        if science_count == 2:
            return "높음"
        elif science_count == 1:
            return "보통"
        else:
            return "대학별 허용조건 확인"

    elif major_group in ["인문사회", "상경", "교육"]:
        if social_count == 2:
            return "높음"
        elif social_count == 1:
            return "보통"
        else:
            return "대학별 허용조건 확인"

    return "별도 확인"


def filter_by_desired_major(
        df,
        desired_major
):
    if not desired_major:
        return df.copy()

    keywords = get_major_keywords(
        desired_major
    )

    def is_related(university_major):

        major = normalize_text(university_major)

        for key in keywords:
            if keywords in keywords:
                return True

        return False

    filtered = df[
        df["모집단위"].apply(is_related)].copy

    return filtered


def percentile_to_grade(percentile):
    if percentile >= 95:
        return 1
    elif percentile >= 86:
        return 2
    elif percentile >= 77:
        return 3
    elif percentile >= 62:
        return 4
    elif percentile >= 45:
        return 5
    elif percentile >= 30:
        return 6
    elif percentile >= 19:
        return 7
    elif percentile >= 10:
        return 8
    else:
        return 9


def evaluate_university_row(row, scores):
    major_group = infer_major_group(row["모집단위"])

    weights = MAJOR_WEIGHTS[major_group]

    student_values = {
        "국어": scores["국어"],
        "수학": scores["수학"],
        "탐구1": scores["탐구1"],
        "탐구2": scores["탐구2"]
    }

    student_score = calculate_weighted_score(student_values, weights)

    university_score = calculate_university_cutoff(row, weights)

    if university_score is None:
        return pd.Series({
            "학과계열": major_group,

            "학생가중점수": student_score,

            "대학가중입결": None,

            "학생환산등급": None,

            "대학환산등급": None,

            "추천유형": "자료부족",

            "가중점수치": None,

            "탐구적합도": check_inquiry_fit(scores, major_group)
        })

    student_grade = (
        percentile_to_grade(
            student_score
        )
    )

    university_grade = (
        percentile_to_grade(
            university_score
        )
    )

    grade_diffrence = student_grade - university_grade

    if 0.8 < grade_diffrence < 1.3:
        recommendation = "상향"

    elif 0 < grade_diffrence <= 0.8:
        recommendation = "적정"
    elif grade_diffrence < 0:
        recommandation = "하향"

    else:
        recommandation = "범위밖"

    return pd.Series({
        "학과계열":
            major_group,

        "학생가중점수":
            round(student_score, 2),

        "대학가중입결":
            round(university_score, 2),

        "학생환산등급":
            student_grade,

        "대학환산등급":
            university_grade,

        "추천유형":
            recommendation,

        "가중점수차":
            round(student_score - university_score, 2),

        "탐구적합도":
            check_inquiry_fit(scores, major_group)
    })

def safe_number(value):

    if value is None:
        return None
    if pd.isna(value):
        return None

    try:
        return float(value)
    except (ValueError, TypeError):
        return None

def calculate_student_simple_average(scores):
    return (
        scores["국어"] + scores["수학"] + scores["탐구1"] + scores["탐구2"]
    ) / 4

def calculate_university_reference_score(row, weights):
    korean = safe_number(row.get("국어70"))
    math = safe_number(row.get("수학70"))
    inquiry1 = safe_number(row.get("탐구1_70"))
    inquiry2 = safe_number(row.get("탐구2_70"))

    if all(value is not None for value in [korean, math, inquiry1, inquiry2]):
        university_scores = {"국어":korean, "수학":math, "탐구1":inquiry1, "탐구2":inquiry2}

        weighted_score = (calculate_weighted_score(university_scores, weights))

        return (weighted_score, "과목별 70% 가중평균")

    recommended_score = safe_number(row.get("추천기준_백분위"))

    if recommended_score is not None:
        indicator_type = row.get("추천지표유형", "추천기준 백분위")

        if pd.isna(indicator_type):
            indicator_type("추천기준 백분위")

        return ( recommended_score, str(indicator_type) )


    return ( None, "자료부족" )

def classify_recommendation (student_score, university_score):
    differnce = ( student_score - university_score )

    if differnce < -8:
        return "매우 상향"

    elif differnce < -4:
        return "소신 상향"

    elif differnce <= 2:
        return "매우 적정"

    elif differnce <= 5:
        return "적정"

    elif differnce <= 8:
        return "안정"

    else:
        return "하향"

def calculate_major_match (university_major, desired_major ):
    if not desired_major:
        return 0

    university_text = (normalize_text(university_major))

    desired_text = (normalize_text(desired_major))

    if (desired_major in university_text or university_text in desired_major):
        return 3

    keywords = get_major_keywords(desired_major)

    for keyword in keywords:
        if keyword in university_text:
            return 2

    desired_group = ( infer_major_group(desired_major) )

    university_group = ( infer_major_group(university_major) )

    if ( desired_group == university_group and desired_group != "기타"):
        return 1

    return 0

def major_match_text (match_score):

    if match_score == 3:
        return "직접 관련"
    elif match_score == 2:
        return "유사 학과"
    elif match_score == 1:
        return "같은 계열"
    else:
        return "다른 계열"

def evaluate_university_row(row, scores):

    major_group = (infer_major_group(row["모집단위"]))

    weights = (MAJOR_WEIGHTS[major_group])

    student_values = {"국어": scores["국어"], "수학": scores["수학"],
                      "탐구1":scores["탐구1"], "탐구2": scores["탐구2"]}

    student_weighted_score = calculate_weighted_score(student_values, weights)

    university_score, score_type = calculate_university_reference_score(row, weights)

    if university_score is None:
        return pd.Series({
            "학과계열":
                major_group,
            "학생비교점수": None,
            "대학비교입결": None,
            "점수차": None,
            "학생환산등급": None,
            "대학환산등급": None,
            "추천유형": "자료부족",
            "세부판정": "자료부족",
            "입결기준유형": score_type,
            "탐구적합도": check_inquiry_fit(scores, major_group),
            "학과일치점수": 0,
            "학과일치도": "자료부족"
        })

    if ( score_type == "과목별 70% 가중평균"):
        student_compare_score = student_weighted_score

    else:
        student_compare_score = calculate_student_simple_average(scores)


    score_difference = student_compare_score - university_score

    recommendation = classify_recommendation(student_compare_score, university_score)

    student_grade = percentile_to_grade(student_compare_score)

    university_grade = percentile_to_grade(university_score)

    major_match = calculate_major_match(row["모집단위"], scores["희망학과"])

    return pd.Series({
        "학과계열": major_group,
        "학생비교점수": round(student_compare_score, 2),
        "대학비교입결": round(university_score, 2),
        "점수차": round(score_difference, 2),
        "학생환산등급": student_grade,
        "대학환산등급": university_grade,
        "추천유형": recommendation,
        "입결기준유형": score_type,
        "탐구적합도": check_inquiry_fit(scores, major_group),
        "학과일치점수": major_match,
        "학과일치도": major_match_text(major_match)
    })
def check_grade_subject(student_grade, university_grade):
    if university_grade is None:
        return "자료없음"

    if pd.isna(university_grade):
        return "자료없음"

    try:
        university_grade = float(university_grade)
        student_grade = float(student_grade)

    except (ValueError, TypeError):
        return "자료없음"

    difference = student_grade - university_grade

    if difference <= 0:
        return "양호"

    elif difference <= 1:
        return "주의"

    else:
        return "불리"


def recommend_universities(df, scores):
    df = df.copy

    evaluation = df.apply(lambda row: evaluate_university_row(row, scores), axis=1)

    df = pd.concat([df, evaluation], axis=1)

    df = df[df["추천유형"] != "자료부족"].copy()

    df["영어판정"] = df["영어70_등급"].apply(lambda grade: check_grade_subject(scores["영어"],grade))

    df["한국사판정"] = df["한국사70_등급"].apply(lambda grade: check_grade_subject(scores["한국사"],grade))

    df["절대점수차"] = df["점수차"].abs()

    inquiry_scores = {"높음": 3, "보통" : 2, "별도 확인": 1, "대학별 허용조건 확인": 0},

    df["탐구적합도점수"] = (df["탐구적합도"].map(inquiry_scores).fillna(0))

    df["추천우선순위"] = (df["학과일치점수"] * 8 + df["탐구적합점수"] + df["대학비교입결"] *0.10 - df["절대점수차"] * 1.5)

    df = df.sort_values(by = ["추천우선순위", "대학비교입결"], ascending = [False, False])

    return df

def split_recommendations(df):
    upward = df[df["추천유형"] == "상향"].copy()

    appropriate = df[df["추천유형"] == "적정"].copy()

    downward = df[df["추천유형"] == "하향"].copy()

    return(upward, appropriate, downward)

def diversify_universities(df, limit=30, max_per_unipersity=3):
    if df.empty:
        return df.copy()

    selected_rows = []
    university_counts = {}

    for index, row in df.iterrows():
        university = row["대학명"]

        current_count = (university_counts.get(university, 0))

        if current_count >= max_per_unipersity:
            continue

        selected_rows.append(index)

        university_counts[university] = current_count + 1

        if (len(selected_rows) >= limit):
            break

    return df.loc[selected_rows].copy()

def get_top_commendations(df, limit=30):
    if df.empty:
        return df.copy()

    top = df.sort_values(by=["대학비교입결", "학과일치점수"], ascending = [False, False])

    top = diversify_universities(top, limit=limit, max_per_unipersity=3)

    return top

def print_recommendations(title, data, limit=30):
    print()
    print("="*150)
    print(title)
    print("="*150)
    if data.empty:
        print("해당 추천 결과가 없습니다.")

        return

    print(f"추천 모집단위: " f"{len(data)}개")
    print()

    columns = [
        "지역", "대학명", "모집군", "모집단위", "학과계열", "학과일치도", "학생비교점수",
        "대학비교입결", "점수차", "추천유형", "세부판정", "입결기준유형"," 탐구적합도",
        "영어판정", "취약과목", "경쟁률"
    ]

    existing_colums = [column for column in columns if column in data.columns]

    print(data[existing_colums].head(limit).to_string(index=False))

def print_recommendation_summary(upward, appropriate, downward):
    print()
    print("="*65)
    print("추천 결과 요약")
    print("="*65)

    print(f"상향 : " f"{len(upward)}개")

    print(f"적정 : " f"{len(appropriate)}개")

    print(f"하향 : " f"{len(downward)}개")

    print("-"*65)

    print("전체 : " f"{len(upward) + len(appropriate) + len(downward)}개")
