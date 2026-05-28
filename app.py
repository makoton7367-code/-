import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill
from io import BytesIO
import re

st.title("発注明細リスト 自動加工ツール")

uploaded_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx"])

if uploaded_file:
    st.info("処理を実行中… 少しお待ちください")

    # =========================
    # ① 元データ読み込み
    # =========================
    df = pd.read_excel(uploaded_file)

    # =========================
    # ② 不要な支店を削除
    # =========================
    df = df[~df["売上支店"].isin(["(支店)リフォーム", "(支店)外構工事", "リフォーム部"])]

    # =========================
    # ③ 業者名の統一
    # =========================
    replace_rules = {
        r"㈱丸増ﾍﾞﾆﾔ商会.*": "㈱丸増ﾍﾞﾆﾔ商会",
        r"中部ﾎｰﾑｻｰﾋﾞｽ㈱.*": "中部ﾎｰﾑｻｰﾋﾞｽ㈱",
        r"㈱ｺﾆｼ.*": "㈱ｺﾆｼ",
        r"㈱ｼﾞｭｰﾃｯｸ.*": "㈱ｼﾞｭｰﾃｯｸ",
        r"ｿﾆﾃｯｸ㈱.*": "ｿﾆﾃｯｸ㈱",
        r"ﾏﾃﾘｱﾙｴｰﾄﾞ㈱.*": "ﾏﾃﾘｱﾙｴｰﾄﾞ㈱",
        r"㈱ﾋﾟｺｲ.*": "㈱ﾋﾟｺｲ",
        r"ﾊﾟﾅｿﾆｯｸﾘﾋﾞﾝｸﾞ.*": "ﾊﾟﾅｿﾆｯｸﾘﾋﾞﾝｸﾞ",
        r"㈱丸八.*": "㈱丸八",
        r"岡田電気産業㈱.*": "岡田電気産業㈱",
        r"㈱山善.*": "㈱山善",
        r"ﾌｫｰﾑ断熱㈱.*": "ﾌｫｰﾑ断熱㈱",
        r"三協立山㈱.*": "三協立山㈱",
        r"㈱ﾔﾏｹﾝ.*": "㈱ﾔﾏｹﾝ",
        r"鳥居金属興業㈱.*": "鳥居金属興業㈱",
        r"杉田ｴｰｽ㈱.*": "杉田ｴｰｽ㈱"
    }

    for pattern, name in replace_rules.items():
        df["業者名"] = df["業者名"].str.replace(pattern, name, regex=True)

    # =========================
    # ④ 日付整形
    # =========================
    date_cols = [
        "契約日", "着工予定日", "着工実績日",
        "上棟予定日", "上棟実績日",
        "木完予定日", "木完実績日",
        "引渡予定日", "引渡実績日"
    ]

    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y/%m/%d")
        df[col] = df[col].astype(str).str.replace(r"/0", "/", regex=True)

    # =========================
    # ⑤ 修正済みシート（メモリ上）
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="修正済み", index=False)

    # =========================
    # ⑥ 集計処理
    # =========================

    # --- vendors_front（あなたの長い辞書） ---
        vendors_back = [
        "日本制震ｼｽﾃﾑ㈱","Solvvy㈱","㈱篠原商店","㈱ｹｰ･ｴｲﾁ･ｹｰ",
        "㈱ｼｰ･ｴｽ･ﾗﾝﾊﾞｰ","ﾛｰﾔﾙ電機㈱","㈱富士通ｾﾞﾈﾗﾙ","㈱ｻﾑｼﾝｸﾞ",
        "ﾕｱｻｸｵﾋﾞｽ㈱","田村駒ｴﾝｼﾞﾆｱﾘﾝｸﾞ㈱","ﾌｫﾜｰﾄﾞ.H.S㈱","東京ｸﾞﾗｽﾛﾝ㈱",
        "ﾌｫｰﾑ断熱㈱","三協立山㈱","㈱ｵﾝﾀﾞ製作所","㈱ｱﾄﾞｳﾞｧﾝｸﾞﾙｰﾌﾟ",
        "㈱ﾔﾏｹﾝ","YKｱｸﾛｽ㈱","菊地合板木工㈱","㈲YSKｻﾎﾟｰﾄ",
        "鳥居金属興業㈱","㈱竹屋化学研究所","㈱関家具","協立ｴｱﾃｯｸ㈱",
        "杉田ｴｰｽ㈱","㈱ﾖﾈｷﾝ"
    ]

    def summarize_vendor(df, vendor, special_items=None, grouped_items=None):
        df_v = df[df["業者名"] == vendor]
        summary = {}

        if special_items:
            for name, items in special_items.items():
                summary[name] = df_v.loc[df_v["工種名"].isin(items), "金額"].sum()

        if grouped_items:
            for item in grouped_items:
                summary[item] = df_v.loc[df_v["工種名"] == item, "金額"].sum()

        exclude = []
        if special_items:
            for items in special_items.values():
                exclude += items
        if grouped_items:
            exclude += grouped_items

        summary["その他"] = df_v.loc[~df_v["工種名"].isin(exclude), "金額"].sum()
        summary[f"{vendor} 合計"] = df_v["金額"].sum()

        return summary

    all_summaries = []

    for vendor, rules in vendors_front.items():
        summary = summarize_vendor(
            df,
            vendor,
            special_items=rules.get("special"),
            grouped_items=rules.get("grouped")
        )
        all_summaries.extend(summary.items())

    for vendor in vendors_back:
        total = df.loc[df["業者名"] == vendor, "金額"].sum()
        all_summaries.append((f"{vendor} 合計", total))

    summary_df = pd.DataFrame(all_summaries, columns=["工種名", "合計金額"])

    # =========================
    # ⑦ 集計シートを追加
    # =========================
    with pd.ExcelWriter(output, engine="openpyxl", mode="a") as writer:
        summary_df.to_excel(writer, sheet_name="集計", index=False)

    # =========================
    # ⑧ ダウンロード
    # =========================
    st.success("加工が完了しました！")

    st.download_button(
        label="加工済みExcelをダウンロード",
        data=output.getvalue(),
        file_name="発注明細_加工済み.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
