# 필요한 라이브러리 설치/불러오기
# pip install pandas scipy matplotlib openpyxl

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats
import platform
import os

# -------------------
# 경로 설정
# -------------------
input_path = r"C:\Users\User\Desktop\data.xlsx"   # 엑셀 파일 절대 경로
out_dir = "output_images"
os.makedirs(out_dir, exist_ok=True)

# 현재 작업 폴더 출력
print(f"\n[현재 작업 폴더]: {os.getcwd()}")

# -------------------
# 한글 폰트 설정
# -------------------
import matplotlib as mpl
mpl.rcParams['axes.unicode_minus'] = False
_system = platform.system()
if _system == "Windows":
    mpl.rcParams['font.family'] = 'Malgun Gothic'
elif _system == "Darwin":
    mpl.rcParams['font.family'] = 'AppleGothic'
else:
    mpl.rcParams['font.family'] = 'sans-serif'

# -------------------
# 1) 데이터 불러오기 및 전처리
# -------------------
df = pd.read_excel(input_path, sheet_name='Sheet1')

df = df.rename(columns={
    'Unnamed: 0': '구군',
    '면적(제곱미터)': '면적_m2',
    '녹지면적': '녹지_m2',
    '평균기온': '평균기온'
})

df['면적_km2'] = df['면적_m2'] / 1e6
df['녹지_km2'] = df['녹지_m2'] / 1e6
df['녹지비율_pct'] = df['녹지_m2'] / df['면적_m2'] * 100

# -------------------
# 2) 기초 통계 출력
# -------------------
print("\n=== 데이터 샘플 ===")
print(df.head().to_string(index=False))
print("\n=== 기초 통계 ===")
print(df[['면적_km2','녹지_km2','녹지비율_pct','평균기온']].describe().T)

# -------------------
# 3) 상관분석 및 단순 선형회귀
# -------------------
x = df['녹지비율_pct'].values
y = df['평균기온'].values

if len(x) < 3:
    print("\n샘플 수가 매우 적습니다(n < 3) — 통계검정 결과 신뢰도 낮음")
else:
    pearson_r, pearson_p = stats.pearsonr(x, y)
    slope, intercept, rvalue, pvalue, stderr = stats.linregress(x, y)
    print("\nPearson 상관계수: r = {:.4f}, p-value = {:.4g}".format(pearson_r, pearson_p))
    print("단순 선형회귀: slope = {:.4f}, intercept = {:.4f}, R^2 = {:.4f}, p-value = {:.4g}"
          .format(slope, intercept, rvalue**2, pvalue))

# -------------------
# 4) 시각화
# -------------------
plt.figure(figsize=(6,5))
plt.scatter(x, y)
x_vals = np.linspace(np.nanmin(x), np.nanmax(x), 100)
y_vals = intercept + slope * x_vals
plt.plot(x_vals, y_vals)
plt.xlabel('녹지비율 (%)')
plt.ylabel('평균기온 (°C)')
plt.title('녹지비율 vs 평균기온')
plt.text(0.95, 0.05, f"r={pearson_r:.3f}\nR²={rvalue**2:.3f}\np={pearson_p:.3g}",
         transform=plt.gca().transAxes, ha='right', va='bottom')
plt.tight_layout()
scatter_path = os.path.join(out_dir, "scatter_regression.png")
plt.savefig(scatter_path, dpi=300)
plt.close()

df['녹지_group'] = pd.qcut(df['녹지비율_pct'], q=3, labels=['낮음','중간','높음'])
group_means = df.groupby('녹지_group')['평균기온'].mean().reindex(['낮음','중간','높음'])

plt.figure(figsize=(5,4))
group_means.plot(kind='bar')
plt.xlabel('녹지비율 그룹')
plt.ylabel('평균기온 (°C)')
plt.title('녹지 그룹별 평균기온')
plt.tight_layout()
bar_path = os.path.join(out_dir, "group_bar.png")
plt.savefig(bar_path, dpi=300)
plt.close()

# -------------------
# 5) 결과 저장 및 경로 안내
# -------------------
out_csv = "incheon_processed.csv"
df.to_csv(out_csv, index=False, encoding='utf-8-sig')

print("\n=== 생성된 파일 경로 ===")
print(f"- 전처리된 CSV: {os.path.abspath(out_csv)}")
print(f"- 산점도 + 회귀선 이미지: {os.path.abspath(scatter_path)}")
print(f"- 녹지그룹 막대그래프 이미지: {os.path.abspath(bar_path)}")
