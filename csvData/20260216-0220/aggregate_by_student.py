#!/usr/bin/env python3
"""
csvData/20260216-0220 フォルダ内のCSVファイルを学生氏名ごとに統合し、サマリーを作成するスクリプト
"""
import pandas as pd
import os
import glob
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- 1. 全CSVファイルを読み込み ---
csv_files = []

# ルート直下のCSVファイル（コピー系を除外）
for f in glob.glob(os.path.join(BASE_DIR, "*.csv")):
    basename = os.path.basename(f)
    # コピーファイルを除外
    if "コピー" in basename:
        continue
    csv_files.append(f)

# student_data フォルダ内のCSVファイル
student_data_dir = os.path.join(BASE_DIR, "student_data")
if os.path.isdir(student_data_dir):
    for f in glob.glob(os.path.join(student_data_dir, "*.csv")):
        csv_files.append(f)

print(f"読み込み対象CSVファイル数: {len(csv_files)}")

# --- 2. 全データを統合 ---
all_dfs = []
for f in csv_files:
    try:
        df = pd.read_csv(f, encoding="utf-8")
        # ヘッダーが期待通りか確認
        if "氏名" in df.columns and "分野" in df.columns:
            all_dfs.append(df)
        else:
            print(f"  スキップ（ヘッダー不一致）: {os.path.basename(f)}")
    except Exception as e:
        print(f"  読み込みエラー: {os.path.basename(f)} - {e}")

combined = pd.concat(all_dfs, ignore_index=True)
print(f"全レコード数: {len(combined)}")

# --- 3. 重複除去 ---
# 同じ学生・日付・時刻・分野・問題数・正答数 の完全重複を除去
before_dedup = len(combined)
combined = combined.drop_duplicates()
print(f"重複除去: {before_dedup} → {len(combined)} レコード")

# --- 4. 氏名の正規化（全角/半角スペースの統一） ---
combined["氏名"] = combined["氏名"].str.strip().str.replace(r"\s+", " ", regex=True)

# --- 4.5 除外学生 ---
EXCLUDE_STUDENTS = ["藤野滉大"]
combined = combined[~combined["氏名"].isin(EXCLUDE_STUDENTS)]
print(f"除外後レコード数: {len(combined)} ({', '.join(EXCLUDE_STUDENTS)} を除外)")

# --- 5. 学生ごとのサマリー作成 ---
# 氏名ごとにグループ化
student_groups = combined.groupby("氏名")

summary_rows = []
for name, group in student_groups:
    student_ids = group["学籍番号"].unique()
    # 代表学籍番号（最も頻度の高いものを使用）
    student_id = group["学籍番号"].value_counts().index[0]
    
    total_questions = group["問題数"].sum()
    total_correct = group["正答数"].sum()
    overall_accuracy = round(total_correct / total_questions * 100, 1) if total_questions > 0 else 0
    
    # 分野別集計
    field_stats = group.groupby("分野").agg(
        問題数合計=("問題数", "sum"),
        正答数合計=("正答数", "sum"),
        取り組み回数=("分野", "count")
    ).reset_index()
    field_stats["正答率(%)"] = round(field_stats["正答数合計"] / field_stats["問題数合計"] * 100, 1)
    
    # 最も弱い分野（正答率が最も低い分野、問題数10以上）
    weak_fields = field_stats[field_stats["問題数合計"] >= 10].sort_values("正答率(%)")
    weakest = weak_fields.iloc[0]["分野"] if len(weak_fields) > 0 else "N/A"
    weakest_rate = weak_fields.iloc[0]["正答率(%)"] if len(weak_fields) > 0 else 0
    
    # 最も強い分野
    strong_fields = field_stats[field_stats["問題数合計"] >= 10].sort_values("正答率(%)", ascending=False)
    strongest = strong_fields.iloc[0]["分野"] if len(strong_fields) > 0 else "N/A"
    strongest_rate = strong_fields.iloc[0]["正答率(%)"] if len(strong_fields) > 0 else 0
    
    # 日付の範囲
    dates = pd.to_datetime(group["日付"], format="%Y/%m/%d", errors="coerce")
    date_min = dates.min().strftime("%m/%d") if pd.notna(dates.min()) else "N/A"
    date_max = dates.max().strftime("%m/%d") if pd.notna(dates.max()) else "N/A"
    
    # 取り組み分野数
    num_fields = group["分野"].nunique()
    
    summary_rows.append({
        "氏名": name,
        "学籍番号": student_id,
        "総問題数": total_questions,
        "総正答数": total_correct,
        "総合正答率(%)": overall_accuracy,
        "取り組み分野数": num_fields,
        "最弱分野": weakest,
        "最弱分野正答率(%)": weakest_rate,
        "最強分野": strongest,
        "最強分野正答率(%)": strongest_rate,
        "学習期間": f"{date_min}〜{date_max}",
        "総取り組み回数": len(group),
    })

summary_df = pd.DataFrame(summary_rows)
summary_df = summary_df.sort_values("氏名")

# --- 6. 学生別の全データシート ---
# 各学生の分野別集計
student_field_rows = []
for name, group in student_groups:
    student_ids = group["学籍番号"].unique()
    student_id = group["学籍番号"].value_counts().index[0]
    
    field_stats = group.groupby("分野").agg(
        問題数合計=("問題数", "sum"),
        正答数合計=("正答数", "sum"),
        取り組み回数=("分野", "count")
    ).reset_index()
    field_stats["正答率(%)"] = round(field_stats["正答数合計"] / field_stats["問題数合計"] * 100, 1)
    
    for _, row in field_stats.iterrows():
        student_field_rows.append({
            "氏名": name,
            "学籍番号": student_id,
            "分野": row["分野"],
            "問題数合計": row["問題数合計"],
            "正答数合計": row["正答数合計"],
            "正答率(%)": row["正答率(%)"],
            "取り組み回数": row["取り組み回数"],
        })

student_field_df = pd.DataFrame(student_field_rows)

# --- 7. Excelに出力 ---
output_path = os.path.join(BASE_DIR, "学生別サマリー_20260216-0220.xlsx")
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    summary_df.to_excel(writer, sheet_name="サマリー", index=False)
    student_field_df.to_excel(writer, sheet_name="学生別分野別詳細", index=False)
    combined.sort_values(["氏名", "日付", "時刻"]).to_excel(writer, sheet_name="全データ", index=False)

print(f"\n=== 出力完了 ===")
print(f"ファイル: {output_path}")
print(f"学生数: {len(summary_df)}")
print(f"\n--- サマリー概要 ---")
print(f"総合正答率の平均: {summary_df['総合正答率(%)'].mean():.1f}%")
print(f"総合正答率の中央値: {summary_df['総合正答率(%)'].median():.1f}%")
print(f"最高正答率: {summary_df['総合正答率(%)'].max():.1f}% ({summary_df.loc[summary_df['総合正答率(%)'].idxmax(), '氏名']})")
print(f"最低正答率: {summary_df['総合正答率(%)'].min():.1f}% ({summary_df.loc[summary_df['総合正答率(%)'].idxmin(), '氏名']})")

# --- 8. Markdownサマリーも作成 ---
md_path = os.path.join(BASE_DIR, "学生別サマリー_20260216-0220.md")
with open(md_path, "w", encoding="utf-8") as f:
    f.write("# 学生別学習サマリー（2026/02/16〜02/20）\n\n")
    f.write(f"- **対象期間**: 2026年2月16日〜2月20日\n")
    f.write(f"- **学生数**: {len(summary_df)}名\n")
    f.write(f"- **総レコード数**: {len(combined)}件\n")
    f.write(f"- **生成日時**: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")
    
    f.write("## 全体統計\n\n")
    f.write(f"| 指標 | 値 |\n|---|---|\n")
    f.write(f"| 総合正答率（平均） | {summary_df['総合正答率(%)'].mean():.1f}% |\n")
    f.write(f"| 総合正答率（中央値） | {summary_df['総合正答率(%)'].median():.1f}% |\n")
    f.write(f"| 最高正答率 | {summary_df['総合正答率(%)'].max():.1f}% ({summary_df.loc[summary_df['総合正答率(%)'].idxmax(), '氏名']}) |\n")
    f.write(f"| 最低正答率 | {summary_df['総合正答率(%)'].min():.1f}% ({summary_df.loc[summary_df['総合正答率(%)'].idxmin(), '氏名']}) |\n\n")
    
    f.write("## 学生別サマリー\n\n")
    f.write("| 氏名 | 学籍番号 | 総問題数 | 総正答数 | 正答率(%) | 分野数 | 回数 | 最弱分野 | 最弱率 | 最強分野 | 最強率 |\n")
    f.write("|---|---|---|---|---|---|---|---|---|---|---|\n")
    for _, row in summary_df.iterrows():
        f.write(f"| {row['氏名']} | {row['学籍番号']} | {row['総問題数']} | {row['総正答数']} | {row['総合正答率(%)']} | {row['取り組み分野数']} | {row['総取り組み回数']} | {row['最弱分野']} | {row['最弱分野正答率(%)']} | {row['最強分野']} | {row['最強分野正答率(%)']} |\n")
    
    f.write("\n## 正答率ランキング（上位10名）\n\n")
    top10 = summary_df.nlargest(10, "総合正答率(%)")
    f.write("| 順位 | 氏名 | 正答率(%) | 総問題数 |\n|---|---|---|---|\n")
    for i, (_, row) in enumerate(top10.iterrows(), 1):
        f.write(f"| {i} | {row['氏名']} | {row['総合正答率(%)']} | {row['総問題数']} |\n")
    
    f.write("\n## 正答率ランキング（下位10名）\n\n")
    bottom10 = summary_df.nsmallest(10, "総合正答率(%)")
    f.write("| 順位 | 氏名 | 正答率(%) | 総問題数 |\n|---|---|---|---|\n")
    for i, (_, row) in enumerate(bottom10.iterrows(), 1):
        f.write(f"| {i} | {row['氏名']} | {row['総合正答率(%)']} | {row['総問題数']} |\n")
    
    f.write("\n## 分野別の全体正答率\n\n")
    field_overall = combined.groupby("分野").agg(
        問題数合計=("問題数", "sum"),
        正答数合計=("正答数", "sum"),
        取り組み人数=("氏名", "nunique"),
    ).reset_index()
    field_overall["正答率(%)"] = round(field_overall["正答数合計"] / field_overall["問題数合計"] * 100, 1)
    field_overall = field_overall.sort_values("正答率(%)")
    
    f.write("| 分野 | 問題数合計 | 正答数合計 | 正答率(%) | 取り組み人数 |\n|---|---|---|---|---|\n")
    for _, row in field_overall.iterrows():
        f.write(f"| {row['分野']} | {row['問題数合計']} | {row['正答数合計']} | {row['正答率(%)']} | {row['取り組み人数']} |\n")

print(f"Markdownサマリー: {md_path}")
print("完了!")
