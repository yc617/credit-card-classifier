import streamlit as st
import pandas as pd
from openpyxl import load_workbook

st.set_page_config(page_title="信用卡新聞分類工具", layout="wide")

st.title("信用卡新聞分類工具")
st.write("MVP：多卡分列、即時修改、防止重複計算")


# -----------------------------
# 讀取 Mastercard raw data 超連結
# -----------------------------
@st.cache_data
def read_news_with_hyperlinks(uploaded_file):
    workbook = load_workbook(uploaded_file, data_only=True)
    sheet = workbook.active

    rows = []

    for row in sheet.iter_rows(min_row=2):

        date_cell = row[0]
        title_cell = row[1]

        date_value = date_cell.value
        title_value = title_cell.value

        url_value = None

        if title_cell.hyperlink:
            url_value = title_cell.hyperlink.target

        if date_value is not None and title_value is not None:
            rows.append({
                "監測日期": date_value,
                "訊息標題": title_value,
                "網址": url_value
            })

    news_df = pd.DataFrame(rows)

    return news_df


# -----------------------------
# 讀 Excel
# -----------------------------
@st.cache_data
def read_excel_file(uploaded_file):
    return pd.read_excel(uploaded_file)


# -----------------------------
# 清理信用卡資料
# -----------------------------
def clean_card_data(card_df):

    # 刪除 Unnamed 欄位
    card_df = card_df.loc[:, ~card_df.columns.astype(str).str.contains("Unnamed")]

    # 刪除全空欄位
    card_df = card_df.dropna(axis=1, how="all")

    # 清理欄位名稱
    card_df.columns = card_df.columns.astype(str).str.strip()

    # 欄位名稱統一
    rename_map = {
        "銀行": "銀行別",
        "卡片名稱": "提及信用卡",
        "發卡組織": "卡組織"
    }

    card_df = card_df.rename(columns=rename_map)

    required_columns = ["銀行別", "提及信用卡", "卡組織"]

    for col in required_columns:
        if col not in card_df.columns:
            st.error(f"信用卡清單缺少欄位：{col}")
            st.stop()

    card_df = card_df[required_columns]

    # 清理內容
    card_df["銀行別"] = card_df["銀行別"].astype(str).str.strip()

    card_df["提及信用卡"] = card_df["提及信用卡"].astype(str).str.strip()

    card_df["卡組織"] = card_df["卡組織"].fillna("ALL").astype(str).str.strip()

    # 空白視為 ALL
    card_df["卡組織"] = card_df["卡組織"].replace(
        ["", "None", "none", "nan", "NaN"],
        "ALL"
    )

    # 刪除空白卡名
    card_df = card_df[
        (card_df["提及信用卡"] != "") &
        (card_df["提及信用卡"] != "nan")
    ]

    return card_df


# -----------------------------
# 初始化 session
# -----------------------------
if "result_df" not in st.session_state:
    st.session_state.result_df = pd.DataFrame(
        columns=[
            "監測日期",
            "訊息標題",
            "網址",
            "銀行別",
            "提及信用卡",
            "卡組織",
            "全文出現次數",
            "實際計入"
        ]
    )

if "processed_urls" not in st.session_state:
    st.session_state.processed_urls = set()


# -----------------------------
# 上傳區
# -----------------------------
news_file = st.file_uploader(
    "請上傳 Mastercard raw data",
    type=["xlsx"]
)

card_file = st.file_uploader(
    "請上傳信用卡清單",
    type=["xlsx"]
)


# -----------------------------
# 讀取新聞
# -----------------------------
if news_file is not None:

    news_df = read_news_with_hyperlinks(news_file)

    st.subheader("新聞 raw data 預覽")

    st.write(f"新聞資料筆數：{len(news_df)}")

    st.dataframe(news_df.head(10), use_container_width=True)


# -----------------------------
# 讀取信用卡清單
# -----------------------------
if card_file is not None:

    raw_card_df = read_excel_file(card_file)

    card_df = clean_card_data(raw_card_df)

    st.subheader("信用卡清單預覽")

    st.write(f"信用卡資料筆數：{len(card_df)}")

    st.dataframe(card_df.head(10), use_container_width=True)


# -----------------------------
# 主流程
# -----------------------------
if news_file is not None and card_file is not None:

    st.divider()

    st.subheader("Step 1：選擇新聞並貼上全文")

    news_options = []

    for index, row in news_df.iterrows():

        url = row["網址"]

        status = "已處理" if url in st.session_state.processed_urls else "未處理"

        option_text = f"{index + 1}. [{status}] {row['訊息標題']}"

        news_options.append(option_text)

    selected_news = st.selectbox(
        "請選擇新聞",
        news_options
    )

    selected_index = news_options.index(selected_news)

    selected_row = news_df.iloc[selected_index]

    selected_date = selected_row["監測日期"]

    selected_title = selected_row["訊息標題"]

    selected_url = selected_row["網址"]

    st.write(f"日期：{selected_date}")

    st.write(f"標題：{selected_title}")

    st.write(f"網址：{selected_url}")

    article_text = st.text_area(
        "請貼上這篇新聞全文",
        height=250
    )

    detected_cards = []

    # -----------------------------
    # 卡片辨識
    # -----------------------------
    if article_text.strip() != "":

        with st.spinner("正在辨識卡片..."):

            for _, card_row in card_df.iterrows():

                bank_name = card_row["銀行別"]

                card_name = card_row["提及信用卡"]

                card_org = card_row["卡組織"]

                mention_count = article_text.count(card_name)

                # 有出現卡片
                if mention_count > 0:

                    detected_cards.append({
                        "銀行別": bank_name,
                        "提及信用卡": card_name,
                        "卡組織": card_org,
                        "全文出現次數": mention_count,
                        "實際計入": 1,
                        "保留": True
                    })

        detected_cards_df = pd.DataFrame(detected_cards)

        # -----------------------------
        # 有抓到卡片
        # -----------------------------
        if len(detected_cards_df) > 0:

            detected_cards_df = detected_cards_df.drop_duplicates(
                subset=["銀行別", "提及信用卡"]
            )

            st.subheader("Step 2：確認抓到的卡片")

            st.write("一篇新聞出現多張卡時，系統會自動拆成多列。")

            edited_detected_df = st.data_editor(
                detected_cards_df,
                use_container_width=True,
                num_rows="dynamic",
                key="detected_editor"
            )

            # -----------------------------
            # 加入分類表
            # -----------------------------
            if st.button("加入或更新這篇新聞"):

                kept_df = edited_detected_df[
                    edited_detected_df["保留"] == True
                ].copy()

                if len(kept_df) == 0:

                    st.warning("沒有保留任何卡片")

                else:

                    # 先刪除同網址舊資料
                    st.session_state.result_df = st.session_state.result_df[
                        st.session_state.result_df["網址"] != selected_url
                    ]

                    new_rows = []

                    for _, row in kept_df.iterrows():

                        new_rows.append({
                            "監測日期": selected_date,
                            "訊息標題": selected_title,
                            "網址": selected_url,
                            "銀行別": row["銀行別"],
                            "提及信用卡": row["提及信用卡"],
                            "卡組織": row["卡組織"],
                            "全文出現次數": row["全文出現次數"],
                            "實際計入": 1
                        })

                    new_df = pd.DataFrame(new_rows)

                    st.session_state.result_df = pd.concat(
                        [st.session_state.result_df, new_df],
                        ignore_index=True
                    )

                    # 同篇同卡防重複
                    st.session_state.result_df = st.session_state.result_df.drop_duplicates(
                        subset=["網址", "銀行別", "提及信用卡"]
                    )

                    st.session_state.processed_urls.add(selected_url)

                    st.success(f"已更新這篇新聞，共加入 {len(new_df)} 張卡片")

        # -----------------------------
        # 沒抓到卡片
        # -----------------------------
        else:

            st.warning("沒有偵測到信用卡，可能是新卡或別名")

st.subheader("手動補充卡片")

manual_bank = st.selectbox(
    "選擇銀行",
    sorted(card_df["銀行別"].dropna().unique()),
    key="manual_bank"
)

manual_card_options = card_df[
    card_df["銀行別"] == manual_bank
]["提及信用卡"].dropna().unique()

manual_card = st.selectbox(
    "選擇要補充的卡片",
    sorted(manual_card_options),
    key="manual_card"
)

manual_card_org = card_df[
    (card_df["銀行別"] == manual_bank) &
    (card_df["提及信用卡"] == manual_card)
]["卡組織"].iloc[0]

if st.button("手動加入這張卡到分類表"):
    new_manual_row = {
        "監測日期": selected_date,
        "訊息標題": selected_title,
        "網址": selected_url,
        "銀行別": manual_bank,
        "提及信用卡": manual_card,
        "卡組織": manual_card_org,
        "全文出現次數": 0,
        "實際計入": 1
    }

    manual_key_exists = (
        (st.session_state.result_df["網址"] == selected_url) &
        (st.session_state.result_df["銀行別"] == manual_bank) &
        (st.session_state.result_df["提及信用卡"] == manual_card)
    ).any()

    if manual_key_exists:
        st.warning("這篇新聞已經有這張卡，系統不會重複加入。")
    else:
        st.session_state.result_df = pd.concat(
            [
                st.session_state.result_df,
                pd.DataFrame([new_manual_row])
            ],
            ignore_index=True
        )

        st.session_state.processed_urls.add(selected_url)

        st.success("已手動加入這張卡。")
    # -----------------------------
    # 分類總表
    # -----------------------------
    st.divider()

    st.subheader("Step 3：即時修改分類總表")

    st.write("可直接修改、刪除、新增列")

    edited_result_df = st.data_editor(
        st.session_state.result_df,
        use_container_width=True,
        num_rows="dynamic",
        key="result_editor"
    )

    # -----------------------------
    # 儲存修改
    # -----------------------------
    if st.button("儲存分類總表修改"):

        edited_result_df = edited_result_df.drop_duplicates(
            subset=["網址", "銀行別", "提及信用卡"]
        )

        st.session_state.result_df = edited_result_df

        st.success("分類總表修改已儲存")


    # -----------------------------
    # 最終輸出
    # -----------------------------
    if len(st.session_state.result_df) > 0:

        final_df = st.session_state.result_df.copy()

        # 提及次數
        final_df["提及次數"] = final_df.groupby(
            ["銀行別", "提及信用卡"]
        )["提及信用卡"].transform("count")

        st.subheader("Step 4：輸出預覽")

        st.dataframe(final_df, use_container_width=True)

    else:

        st.info("目前尚未加入分類結果")