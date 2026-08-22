import re
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

EXCEL_FILE = Path(__file__).resolve().with_name(
  "전국대학_정시입결_공식통합_v4_지방국립확장.xlsx"
)
SHEET_NAME = "정시입결_2025"

UPWARD_LIMIT = -3
APPROPRIATE_LIMIT = 7

DISPLAY_LIMIT = 30

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




  "데이터": [
      "데이터",
      "통계",
      "인공지능",
      "ai",
      "컴퓨터"
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
      "정보통신",
      "통신"
  ],




  "전기": [
      "전기",
      "전자",
      "반도체",
      "정보통신"
  ],




  "반도체": [
      "반도체",
      "전자",
      "전기",
      "신소재"
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
  ],




  "심리": [
      "심리"
  ],




  "간호": [
      "간호"
  ],




  "의학": [
      "의예",
      "의학"
  ],




  "수의": [
      "수의"
  ]
}




SCIENCE_INQUIRY_KEYWORDS = [
  "물리",
  "화학",
  "생명",
  "생물",
  "지구"
]




SOCIAL_INQUIRY_KEYWORDS = [
  "생활과윤리",
  "생윤",
  "윤리와사상",
  "윤사",
  "한국지리",
  "한지",
  "세계지리",
  "세지",
  "동아시아사",
  "동사",
  "세계사",
  "경제",
  "정치와법",
  "정법",
  "사회문화",
  "사문"
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




REGION_ALIASES = {
  "서울": "서울특별시",
  "서울특별시": "서울특별시",


  "경기": "경기도",
  "경기도": "경기도",


  "인천": "인천광역시",
  "인천광역시": "인천광역시",


  "대전": "대전광역시",
  "대전광역시": "대전광역시",


  "광주": "광주광역시",
  "광주광역시": "광주광역시",


  "대구": "대구광역시",
  "대구광역시": "대구광역시",


  "부산": "부산광역시",
  "부산광역시": "부산광역시",


  "울산": "울산광역시",
  "울산광역시": "울산광역시",


  "세종": "세종특별자치시",
  "세종특별자치시": "세종특별자치시",


  "강원": "강원특별자치도",
  "강원특별자치도": "강원특별자치도",


  "충북": "충청북도",
  "충청북도": "충청북도",


  "충남": "충청남도",
  "충청남도": "충청남도",


  "전북": "전북특별자치도",
  "전북특별자치도": "전북특별자치도",


  "전남": "전라남도",
  "전라남도": "전라남도",


  "경북": "경상북도",
  "경상북도": "경상북도",


  "경남": "경상남도",
  "경상남도": "경상남도",


  "제주": "제주특별자치도",
  "제주특별자치도": "제주특별자치도"
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




  if subject == "":
      return "기타"




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
          return "의약보건"




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
      "에너지",
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




  business_keywords = [
      "경영",
      "경제",
      "회계",
      "세무",
      "금융",
      "무역",
      "국제통상"
  ]




  for keyword in business_keywords:
      if normalize_text(keyword) in name:
          return "상경"




  education_keywords = [
      "교육",
      "초등교육",
      "유아교육"
  ]




  for keyword in education_keywords:
      if normalize_text(keyword) in name:
          return "교육"




  humanities_keywords = [
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




  for keyword in humanities_keywords:
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








def input_student_profile():
  print()
  print("===== 모의고사 성적 입력 =====")




  inquiry1_name = input(
      "탐구1 과목명: "
  ).strip()




  inquiry2_name = input(
      "탐구2 과목명: "
  ).strip()




  korean = input_percentile("국어")
  math = input_percentile("수학")




  inquiry1 = input_percentile(
      inquiry1_name if inquiry1_name else "탐구1"
  )


  inquiry2 = input_percentile(
      inquiry2_name if inquiry2_name else "탐구2"
  )




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
      "탐구1과목": inquiry1_name,
      "탐구2과목": inquiry2_name,
      "영어": english,
      "한국사": history,
      "희망학과": desired_major
  }








def input_region():
  print()
  print("[희망 지역]")
  print("서울, 경기, 인천처럼 입력하세요.")
  print("전체 지역이면 Enter 또는 '전체'를 입력하세요.")




  region = input(
      "희망 지역: "
  ).strip()




  if region == "" or region == "전체":
      return None




  return REGION_ALIASES.get(
      region,
      region
  )








def percentile_to_grade(percentile):
  if percentile >= 96:
      return 1


  elif percentile >= 89:
      return 2


  elif percentile >= 77:
      return 3


  elif percentile >= 60:
      return 4


  elif percentile >= 40:
      return 5


  elif percentile >= 23:
      return 6


  elif percentile >= 11:
      return 7


  elif percentile >= 4:
      return 8


  else:
      return 9








def load_university_data():
  if not EXCEL_FILE.exists():
      raise FileNotFoundError(
          "\n입결 Excel 파일을 찾을 수 없습니다.\n"
          f"{EXCEL_FILE.name}\n\n"
          "파이썬 파일과 Excel 파일을 같은 폴더에 넣으세요."
      )




  df = pd.read_excel(
      EXCEL_FILE,
      sheet_name=SHEET_NAME
  )




  required_columns = [
      "지역",
      "대학명",
      "모집군",
      "모집단위",
      "국어70",
      "수학70",
      "탐구1_70",
      "탐구2_70",
      "영어70_등급",
      "한국사70_등급",
      "경쟁률",
      "추천기준_백분위"
  ]




  missing_columns = [
      column for column in required_columns
      if column not in df.columns
  ]




  if missing_columns:
      raise ValueError(
          "Excel 파일에 필요한 컬럼이 없습니다.\n"
          f"누락 컬럼: {missing_columns}"
      )




  numeric_columns = [
      "국어70",
      "수학70",
      "탐구1_70",
      "탐구2_70",
      "영어70_등급",
      "한국사70_등급",
      "경쟁률",
      "추천기준_백분위"
  ]




  for column in numeric_columns:
      df[column] = pd.to_numeric(
          df[column],
          errors="coerce"
      )




  df = df.dropna(
      subset=[
          "대학명",
          "모집단위"
      ]
  ).copy()




  return df








def filter_by_region(
      df,
      region
):
  if region is None:
      return df.copy()




  return df[
      df["지역"] == region
  ].copy()








def calculate_weighted_score(
      values,
      weights
):
  total_score = 0
  total_weight = 0




  for subject, weight in weights.items():
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








def check_inquiry_fit(scores, major_group):
  inquiry1_type = classify_inquiry_subject(
      scores["탐구1과목"]
  )




  inquiry2_type = classify_inquiry_subject(
      scores["탐구2과목"]
  )




  types = [
      inquiry1_type,
      inquiry2_type
  ]




  science_count = (
      types.count("과탐")
  )




  social_count = (
      types.count("사탐")
  )




  if major_group in [
      "공학",
      "자연과학",
      "의약보건"
  ]:
      if science_count == 2:
          return "높음"


      elif science_count == 1:
          return "보통"


      else:
          return "대학별 허용조건 확인"




  elif major_group in [
      "인문사회",
      "상경",
      "교육"
  ]:
      if social_count == 2:
          return "높음"


      elif social_count == 1:
          return "보통"


      else:
          return "대학별 허용조건 확인"




  return "별도 확인"








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
      scores["국어"]
      + scores["수학"]
      + scores["탐구1"]
      + scores["탐구2"]
  ) / 4








def calculate_university_reference_score(row, weights):
  korean = safe_number(row.get("국어70"))
  math = safe_number(row.get("수학70"))
  inquiry1 = safe_number(row.get("탐구1_70"))
  inquiry2 = safe_number(row.get("탐구2_70"))




  subject_values = [
      korean,
      math,
      inquiry1,
      inquiry2
  ]




  if sum(
      value is not None
      for value in subject_values
  ) >= 3:




      university_scores = {
          "국어": korean,
          "수학": math,
          "탐구1": inquiry1,
          "탐구2": inquiry2
      }




      weighted_score = (
          calculate_weighted_score(
              university_scores,
              weights
          )
      )




      return (
          weighted_score,
          "과목별 70% 가중평균"
      )




  recommended_score = safe_number(
      row.get("추천기준_백분위")
  )




  if recommended_score is not None:
      indicator_type = row.get(
          "추천지표유형",
          "추천기준 백분위"
      )




      if pd.isna(indicator_type):
          indicator_type = (
              "추천기준 백분위"
          )




      return (
          recommended_score,
          str(indicator_type)
      )




  return (
      None,
      "자료부족"
  )








def classify_recommendation(
      student_score,
      university_score
):
  difference = (
      student_score
      - university_score
  )




  if difference < UPWARD_LIMIT:
      return "상향"




  elif difference <= APPROPRIATE_LIMIT:
      return "적정"




  else:
      return "하향"








def classify_detail(
      student_score,
      university_score
):
  difference = (
      student_score
      - university_score
  )




  if difference < -10:
      return "매우 상향"




  elif difference < -3:
      return "소신 상향"




  elif difference <= 2:
      return "매우 적정"




  elif difference <= 7:
      return "적정"




  elif difference <= 12:
      return "안정"




  elif difference <= 20:
      return "매우 안정"




  else:
      return "충분한 안정"








def calculate_major_match(
      university_major,
      desired_major
):
  if not desired_major:
      return 0




  university_text = (
      normalize_text(university_major)
  )




  desired_text = (
      normalize_text(desired_major)
  )




  if (
      desired_text in university_text
      or university_text in desired_text
  ):
      return 3




  keywords = get_major_keywords(
      desired_major
  )




  for keyword in keywords:
      if keyword in university_text:
          return 2




  desired_group = (
      infer_major_group(desired_major)
  )




  university_group = (
      infer_major_group(university_major)
  )




  if (
      desired_group == university_group
      and desired_group != "기타"
  ):
      return 1




  return 0








def major_match_text(match_score):
  if match_score == 3:
      return "직접 관련"


  elif match_score == 2:
      return "유사 학과"


  elif match_score == 1:
      return "같은 계열"


  else:
      return "다른 계열"








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








def check_weak_subjects(row, scores):
  comparisons = [
      ("국어", scores["국어"], row.get("국어70")),
      ("수학", scores["수학"], row.get("수학70")),
      (
          scores["탐구1과목"] or "탐구1",
          scores["탐구1"],
          row.get("탐구1_70")
      ),
      (
          scores["탐구2과목"] or "탐구2",
          scores["탐구2"],
          row.get("탐구2_70")
      )
  ]




  warnings = []




  for (
      subject,
      student_score,
      cutoff
  ) in comparisons:




      cutoff = safe_number(cutoff)




      if cutoff is None:
          continue




      difference = float(student_score) - cutoff




      if difference <= -10:
          warnings.append(subject)




  if warnings:
      return ", ".join(warnings)




  return "없음"








def evaluate_university_row(row, scores):
  major_group = (
      infer_major_group(row["모집단위"])
  )




  weights = (
      MAJOR_WEIGHTS[major_group]
  )




  student_values = {
      "국어": scores["국어"],
      "수학": scores["수학"],
      "탐구1": scores["탐구1"],
      "탐구2": scores["탐구2"]
  }




  student_weighted_score = (
      calculate_weighted_score(
          student_values,
          weights
      )
  )




  (
      university_score,
      score_type
  ) = calculate_university_reference_score(
      row,
      weights
  )




  if university_score is None:
      return pd.Series({
          "학과계열": major_group,
          "학생비교점수": None,
          "대학비교입결": None,
          "점수차": None,
          "학생환산등급": None,
          "대학환산등급": None,
          "추천유형": "자료부족",
          "세부판정": "자료부족",
          "입결기준유형": score_type,
          "탐구적합도": check_inquiry_fit(
              scores,
              major_group
          ),
          "학과일치점수": 0,
          "학과일치도": "자료부족"
      })




  if score_type == "과목별 70% 가중평균":
      student_compare_score = (
          student_weighted_score
      )




  else:
      student_compare_score = (
          calculate_student_simple_average(scores)
      )




  score_difference = (
      student_compare_score
      - university_score
  )




  recommendation = classify_recommendation(
      student_compare_score,
      university_score
  )




  detail = classify_detail(
      student_compare_score,
      university_score
  )




  student_grade = (
      percentile_to_grade(student_compare_score)
  )




  university_grade = (
      percentile_to_grade(university_score)
  )




  major_match = calculate_major_match(
      row["모집단위"],
      scores["희망학과"]
  )




  return pd.Series({
      "학과계열": major_group,
      "학생비교점수": round(student_compare_score, 2),
      "대학비교입결": round(university_score, 2),
      "점수차": round(score_difference, 2),
      "학생환산등급": student_grade,
      "대학환산등급": university_grade,
      "추천유형": recommendation,
      "세부판정": detail,
      "입결기준유형": score_type,
      "탐구적합도": check_inquiry_fit(
          scores,
          major_group
      ),
      "학과일치점수": major_match,
      "학과일치도": major_match_text(major_match)
  })








def recommend_universities(df, scores):
  conflict_columns = [
      "학생평균백분위",
      "점수차",
      "추천구간",




      "학과계열",
      "학생비교점수",
      "대학비교입결",




      "학생환산등급",
      "대학환산등급",




      "추천유형",
      "세부판정",




      "입결기준유형",
      "탐구적합도",




      "학과일치점수",
      "학과일치도",




      "영어판정",
      "한국사판정",




      "취약과목",
      "절대점수차",
      "탐구적합점수",
      "추천우선순위"
  ]




  columns_to_drop = [
      column for column in conflict_columns
      if column in df.columns
  ]




  if columns_to_drop:
      df = df.drop(
          columns=columns_to_drop
      ).copy()




  else:
      df = df.copy()




  evaluation = df.apply(
      lambda row: evaluate_university_row(
          row,
          scores
      ),
      axis=1
  )




  result = pd.concat(
      [
          df.reset_index(drop=True),
          evaluation.reset_index(drop=True)
      ],
      axis=1
  )




  result = result[
      result["추천유형"] != "자료부족"
  ].copy()




  result["영어판정"] = (
      result["영어70_등급"].apply(
          lambda grade: check_grade_subject(
              scores["영어"],
              grade
          )
      )
  )




  result["한국사판정"] = (
      result["한국사70_등급"].apply(
          lambda grade: check_grade_subject(
              scores["한국사"],
              grade
          )
      )
  )




  result["취약과목"] = (
      result.apply(
          lambda row: check_weak_subjects(
              row,
              scores
          ),
          axis=1
      )
  )




  result["절대점수차"] = (
      result["점수차"].abs()
  )




  inquiry_scores = {
      "높음": 3,
      "보통": 2,
      "별도 확인": 1,
      "대학별 허용조건 확인": 0,
      "허용조건 확인": 0
  }




  result["탐구적합점수"] = (
      result["탐구적합도"]
      .map(inquiry_scores)
      .fillna(0)
  )




  result["추천우선순위"] = (
      result["학과일치점수"] * 10
      + result["탐구적합점수"] * 1
      + result["대학비교입결"] * 0.15
      - result["절대점수차"] * 1.2
  )




  result = result.sort_values(
      by=[
          "추천우선순위",
          "대학비교입결"
      ],
      ascending=[
          False,
          False
      ]
  )




  return result








def split_recommendations(df):
  upward = df[
      df["추천유형"] == "상향"
  ].copy()




  appropriate = df[
      df["추천유형"] == "적정"
  ].copy()




  downward = df[
      df["추천유형"] == "하향"
  ].copy()




  return (
      upward,
      appropriate,
      downward
  )








def diversify_universities(
      df,
      limit=30,
      max_per_university=3
):
  if df.empty:
      return df.copy()




  selected_rows = []
  university_counts = {}




  for index, row in df.iterrows():
      university = row["대학명"]




      current_count = (
          university_counts.get(
              university,
              0
          )
      )




      if current_count >= max_per_university:
          continue




      selected_rows.append(index)




      university_counts[university] = (
          current_count + 1
      )




      if len(selected_rows) >= limit:
          break

  return df.loc[
      selected_rows
  ].copy()

def get_top_recommendations(df, limit=30):
  if df.empty:
      return df.copy()

  top = df.sort_values(
      by=[
          "대학비교입결",
          "학과일치점수"
      ],
      ascending=[
          False,
          False
      ]
  )

  top = diversify_universities(
      top,
      limit=limit,
      max_per_university=3
  )

  return top


def print_recommendation_summary(
      upward,
      appropriate,
      downward
):
  print()
  print("=" * 65)
  print("추천 결과 요약")
  print("=" * 65)
  print(
      f"상향 : {len(upward)}개"
  )
  print(
      f"적정 : {len(appropriate)}개"
  )
  print(
      f"하향 : {len(downward)}개"
  )
  print("-" * 65)
  total = (
      len(upward)
      + len(appropriate)
      + len(downward)
  )
  print(
      f"전체 : {total}개"
  )

INQUIRY_SUBJECT = [
    "물리학1",
    "물리학2",
    "화학1",
    "화학2",
    "생명과학1",
    "생명과학2",
    "지구과학1",
    "지구과학2",
    "생활과윤리",
    "윤리와사상",
    "한국지리",
    "세계지리",
    "동아시아사",
    "세계사",
    "경제",
    "정치와법",
    "사회문화"
]

REGION_CHOICES = [
    "전체", "서울", "경기", "인천", "대전", "광주", "대구", "부산", "울산", "세종", "강원",
    "충북", "충남", "전북", "전남", "경북", "경남", "제주"
]



class UniversityRecommenderGUI:
    def __init__(self, root):
        self.root = root

        self.root.title("정시 대학 추천 시스템")

        self.root.geomerty("1500x850")

        self.root.minsize(1150, 650)

        self.df = None

        self.all_results = None
        self.top_results = None
        self.upward_result = None
        self.appropriate_result = None
        self.downward_result = None

        self.last_scores = None
        self.last_region = "전체"

        self.tress = {}
        self.create_widgets()
        self.load_data_at_start()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=15)

        main_frame.pack(fill="both", expand=True)

        title_label = ttk.Label(main_frame, text="정시 대학 추천 시스템", font=("Arial",22,"bold"))

        title_label.pack(anchor="w")

        description_label = ttk.Label(
            main_frame, text=("모의고사 백분위와 탐구 과목, 희망학과, 지역을 입력하면 "
                              "전년도 정시 입결 기준으로 대학교를 추천합니다"))

        description_label.pack(anchor="w" , pady=(5, 15))

        top_frame = ttk.Frame(main_frame)

        top_frame.pack(fill="x")

        input_frame = ttk.LabelFrame(top_frame, text="학생 성적 입력", padding=12)

        input_frame.pack(side="left", fill="x", expand=True)

        ttk.Label(input_frame, text="국어 백분위").grid(row=0, column=0, padx=5, pady=5)

        self.korean_var = tk.StringVar()

        ttk.Entry(
            input_frame,
            textvariable=self.korean_var,
            width=10
        ).grid(row=0, column=1, padx=5, pady=5)


        ttk.Label(input_frame, text="수학 백분위").grid(row=0, column=2, padx=5, pady=5)

        self.math_var = tk.StringVar()

        ttk.Entry(
            input_frame,
            textvariable=self.math_var,
            width=10
        ).grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(input_frame, text="영어 등급").grid(row=0, column=4, padx=5, pady=5)

        self.english_var = tk.StringVar()

        ttk.Combobox(input_frame, textvariable=self.english_var, values=list(
            range(1, 10)),width=8, state="readonly",).grid(row=0, column=5, padx=5, pady=5)




        ttk.Label(input_frame, text="한국사 등급").grid(row=0, column=6, padx=5, pady=5)

        self.history_var = tk.StringVar()

        ttk.Combobox(input_frame, textvariable=self.history_var, values=list(
            range(1, 10)),width=8, state="readonly",).grid(row=0, column=7, padx=5, pady=5)



        ttk.Label(input_frame, text="탐구1 과목").grid(row=1, column=0, padx=5, pady=5)

        self.inquiry1_name_var = tk.StringVar()

        ttk.Combobox(input_frame,
                    textvariable=self.inquiry1_name_var,
                    values=INQUIRY_SUBJECT,
                    width=14
                     ).grid(
                    row=1,
                    column=1,
                    padx=5, pady=5)

        ttk.Label(input_frame, text="탐구1 백분위").grid(row=1, column=2, padx=5, pady=5)

        self.inquiry1_var = tk.StringVar()

        ttk.Entry(
            input_frame,
            textvariable=self.inquiry1_var,
            width=10
        ).grid(row=1, column=3, padx=5, pady=5)




        ttk.Label(input_frame, text="탐구2 과목").grid(row=1, column=4, padx=5, pady=5)

        self.inquiry2_name_var = tk.StringVar()

        ttk.Combobox(input_frame,
                    textvariable=self.inquiry2_name_var,
                    values=INQUIRY_SUBJECT,
                    width=14
                     ).grid(
                    row=1,
                    column=5,
                    padx=5, pady=5)

        ttk.Label(input_frame, text="탐구2 백분위").grid(row=1, column=6, padx=5, pady=5)

        self.inquiry2_var = tk.StringVar()

        ttk.Entry(
            input_frame,
            textvariable=self.inquiry2_var,
            width=10
        ).grid(row=1, column=7, padx=5, pady=5)


        ttk.Label(input_frame, text="희망 학과").grid(row=2, column=0, padx=5, pady=5)
        self.major_var = tk.StringVar()

        ttk.Entry(input_frame, textvariable=self.major_var, width=25).grid(row=2, column=1, padx=5, pady=5, sticky="ew")



        ttk.Label(input_frame, text="희망 지역").grid(row=2, column=3, padx=5, pady=5)

        self.region_var = tk.StringVar()

        ttk.Combobox(input_frame,
                     textvariable=self.region_var,
                     values=REGION_CHOICES,
                     width=12,
                     state="readonly",
                     ).grid(
                     row=2,
                     column=4,
                     padx=5,
                     pady=5)

        ttk.Button(input_frame, text="대학 추천하기", command=self.run_recommendation).grid(
            row=2,
            column=5,
            padx=5,
            pady=5,
            sticky="ew",
        )

        ttk.Button(input_frame, text="입력 초기화", command=self.clear_inputs).grid(
            row=2,
            column=6,
            padx=5,
            pady=5,
            sticky="ew"
        )

        ttk.Button(input_frame, text="결과 Excel 저장", command=self.save_results).grid(
            row=2,
            column=7,
            padx=5,
            pady=5,
            sticky="ew"
        )

        summary_frame = ttk.LabelFrame(top_frame, text="추천 요약", padding=12)
        summary_frame.pack(side="left", fill="y", padx=(12, 0))

        self.summary_var = tk.StringVar(value="입결 데이터를 불러오는 중...")

        ttk.Label(summary_frame, textvariable=self.summary_var, justify="left",
                  font=("Arial", 11, "bold")).pack(anchor="w")

        self.notebook = ttk.Notebook(main_frame)

        self.notebook.pack(fill="both", expand=True, pady=(15, 0))

        self.result_columns = [
            "지역",
            "대학명",
            "모집군",
            "모집단위",
            "학과계열",
            "학과일치도",
            "학생비교점수",
            "대학비교입결",
            "점수차",
            "추천유형",
            "세부판정",
            "입결기준유형",
            "탐구적합도",
            "영어판정",
            "한국사판정",
            "취약과목",
            "경쟁률"
        ]

        self.create_result_tab("top", "상위권 추천")
        self.create_result_tab("upward", "상향")
        self.create_result_tab("appropriate", "적정")
        self.create_result_tab("downward", "하향")
        self.create_result_tab("all", "전체")


    def create_result_tab(self, key, title):
        pass








