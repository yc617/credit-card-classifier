import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re
import time
from urllib.parse import urljoin, urlparse

import requests
import urllib3
from bs4 import BeautifulSoup, Tag

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    import trafilatura
except Exception:
    trafilatura = None


st.set_page_config(page_title="信用卡新聞分類工具 v11", layout="wide", initial_sidebar_state="collapsed")

def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r"\s+", "", text)
    return text


def split_terms(value):
    if pd.isna(value):
        return []
    value = str(value).strip()
    if value == "" or value.lower() in ["nan", "none"]:
        return []
    separators = ["，", "、", "；", ";", "/", "|", "\n"]
    for sep in separators:
        value = value.replace(sep, ",")
    return [term.strip() for term in value.split(",") if term.strip() != ""]


def contains_term(text_norm, term):
    term_norm = normalize_text(term)
    if term_norm == "":
        return False
    return term_norm in text_norm


def count_term(text_norm, term):
    term_norm = normalize_text(term)
    if term_norm == "":
        return 0
    return text_norm.count(term_norm)


def is_yes(value):
    if pd.isna(value):
        return False
    value = str(value).strip().upper()
    return value in ["Y", "YES", "TRUE", "1", "是", "需"]


def safe_int(value, default=1):
    try:
        return int(value)
    except Exception:
        return default


def get_contexts(text, keyword, window=100):
    if pd.isna(text) or pd.isna(keyword):
        return [""]
    text = str(text)
    keyword = str(keyword)
    text_lower = text.lower()
    keyword_lower = keyword.lower()
    contexts = []
    start = 0
    while True:
        pos = text_lower.find(keyword_lower, start)
        if pos == -1:
            break
        left = max(0, pos - window)
        right = min(len(text), pos + len(keyword) + window)
        contexts.append(text[left:right])
        start = pos + len(keyword)
    if len(contexts) == 0:
        contexts.append(text[:300])
    return contexts


def is_external_http_url(url):
    if not isinstance(url, str):
        return False
    url = url.strip()
    return url.startswith("http")


def get_domain(url):
    try:
        return urlparse(url).netloc.lower()
    except Exception:
        return ""


# 這些詞屬於泛稱或卡等級，不能單獨判定為某一張信用卡。
# 例如新聞只寫「現金回饋卡」或「世界卡」，不能把所有銀行的現金回饋卡、世界卡都列入。
HIGH_RISK_GENERIC_TERMS = {
    "信用卡", "金融卡", "簽帳金融卡", "debit卡", "debit",
    "雙幣卡", "商務卡", "商旅卡", "白金卡", "御璽卡", "御璽",
    "鈦金卡", "鈦金", "晶緻卡", "晶緻", "世界卡", "世界",
    "無限卡", "無限", "極緻卡", "極緻", "聯名卡", "一卡通聯名卡",
    "現金回饋卡", "現金回饋", "分期卡", "分期", "哩程卡",
    "旅遊卡", "悠遊卡", "一卡通", "icash", "icash卡",
    # 卡優長文常見「卡等 / 泛稱」，只能做輔助資訊，不能獨立產生信用卡列。
    "世界商務卡", "鈦金商務卡", "鈦金商旅卡", "jcb晶緻卡", "現金回饋jcb卡",
    "銀行卡", "企業卡", "企業聯名卡",
}


BANK_ALIAS_MAP = {
    "台北富邦銀行": ["台北富邦", "富邦", "北富銀", "富邦銀"],
    "富邦銀行": ["台北富邦", "富邦", "北富銀", "富邦銀"],
    "國泰世華銀行": ["國泰世華", "國泰", "國泰世華銀"],
    "國泰世華": ["國泰世華", "國泰", "國泰世華銀"],
    "中國信託銀行": ["中國信託", "中信", "中信銀", "中國信託銀"],
    "中國信託": ["中國信託", "中信", "中信銀", "中國信託銀"],
    "玉山銀行": ["玉山", "玉山銀"],
    "台新銀行": ["台新", "台新銀", "Richart"],
    "永豐銀行": ["永豐", "永豐銀", "大戶", "DAWHO"],
    "遠東商銀": ["遠東商銀", "遠銀", "遠東銀行", "快樂銀行"],
    "遠東銀行": ["遠東商銀", "遠銀", "遠東銀行"],
    "聯邦銀行": ["聯邦", "聯邦銀"],
    "第一銀行": ["第一銀行", "一銀", "第一銀"],
    "華南銀行": ["華南", "華南銀"],
    "兆豐銀行": ["兆豐", "兆豐銀"],
    "上海商銀": ["上海商銀", "上海銀行"],
    "合作金庫": ["合作金庫", "合庫"],
    "凱基銀行": ["凱基", "凱基銀"],
    "新光銀行": ["新光", "新光銀"],
    "元大銀行": ["元大", "元大銀"],
    "星展銀行": ["星展", "DBS"],
    "滙豐銀行": ["滙豐", "匯豐", "HSBC"],
    "渣打銀行": ["渣打", "渣打銀"],
    "美國運通": ["美國運通", "Amex", "American Express"],
}


def is_high_risk_generic_term(term):
    term_norm = normalize_text(term)
    if term_norm == "":
        return False
    return any(term_norm == normalize_text(g) for g in HIGH_RISK_GENERIC_TERMS)


def normalize_card_for_dedupe(card_name):
    value = normalize_text(card_name)
    # 處理 CUBE / CUBE卡、熊本熊 / 熊本熊卡 這類重複命中。
    for suffix in ["信用卡", "聯名卡", "卡"]:
        suffix_norm = normalize_text(suffix)
        if value.endswith(suffix_norm) and len(value) > len(suffix_norm) + 1:
            value = value[: -len(suffix_norm)]
            break
    return value


def get_bank_alias_terms(bank_name):
    bank_name = str(bank_name).strip()
    aliases = [bank_name]
    aliases.extend(BANK_ALIAS_MAP.get(bank_name, []))
    # 銀行名稱本身也常會有「銀行」二字，可以補一個去掉銀行的版本。
    if bank_name.endswith("銀行"):
        aliases.append(bank_name.replace("銀行", ""))
    return sorted(set([a for a in aliases if a]))


def has_bank_context(text_norm, bank_name):
    return any(contains_term(text_norm, alias) for alias in get_bank_alias_terms(bank_name))


@st.cache_data(show_spinner=False)
def read_excel_file(uploaded_file, sheet_name=0):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


@st.cache_data(show_spinner=False)
def read_news_with_hyperlinks(uploaded_file):
    """讀 Mastercard raw data：A 欄日期、B 欄標題與 hyperlink 報告頁。"""
    uploaded_file.seek(0)
    workbook = load_workbook(uploaded_file, data_only=True)
    sheet = workbook.active
    rows = []
    for row_index, row in enumerate(sheet.iter_rows(min_row=2), start=1):
        date_cell = row[0]
        title_cell = row[1]
        date_value = date_cell.value
        title_value = title_cell.value
        report_url = None
        if title_cell.hyperlink:
            report_url = title_cell.hyperlink.target
        if date_value is not None and title_value is not None:
            rows.append({
                "news_order": row_index,
                "監測日期": date_value,
                "訊息標題": title_value,
                "網址": report_url,
            })
    return pd.DataFrame(rows)


def clean_card_data(card_df):
    card_df = card_df.loc[:, ~card_df.columns.astype(str).str.contains("Unnamed")]
    card_df = card_df.dropna(axis=1, how="all")
    card_df.columns = card_df.columns.astype(str).str.strip()
    rename_map = {
        "銀行": "銀行別",
        "卡片名稱": "提及信用卡",
        "正式卡名": "提及信用卡",
        "發卡組織": "卡組織",
    }
    card_df = card_df.rename(columns=rename_map)
    required_columns = ["銀行別", "提及信用卡", "卡組織"]
    for col in required_columns:
        if col not in card_df.columns:
            st.error(f"信用卡清單缺少欄位：{col}")
            st.stop()
    card_df = card_df[required_columns].copy()
    card_df["銀行別"] = card_df["銀行別"].astype(str).str.strip()
    card_df["提及信用卡"] = card_df["提及信用卡"].astype(str).str.strip()
    card_df["卡組織"] = card_df["卡組織"].fillna("ALL").astype(str).str.strip()
    card_df["卡組織"] = card_df["卡組織"].replace(["", "None", "none", "nan", "NaN"], "ALL")
    card_df = card_df[(card_df["提及信用卡"] != "") & (card_df["提及信用卡"].str.lower() != "nan")].copy()
    return card_df.drop_duplicates(subset=["銀行別", "提及信用卡", "卡組織"])


def clean_keyword_data(keyword_df):
    keyword_df = keyword_df.loc[:, ~keyword_df.columns.astype(str).str.contains("Unnamed")]
    keyword_df = keyword_df.dropna(axis=1, how="all")
    keyword_df.columns = keyword_df.columns.astype(str).str.strip()
    rename_map = {
        "關鍵字": "主關鍵字",
        "正式卡名": "提及信用卡",
        "卡片名稱": "提及信用卡",
        "發卡組織": "卡組織",
        "類型": "判定類型",
    }
    keyword_df = keyword_df.rename(columns=rename_map)
    required_columns = ["主關鍵字", "銀行別", "提及信用卡"]
    for col in required_columns:
        if col not in keyword_df.columns:
            st.error(f"關鍵字判定表缺少欄位：{col}")
            st.stop()
    optional_columns = {
        "啟用": "Y",
        "輔助關鍵字": "",
        "排除關鍵字": "",
        "卡組織": "ALL",
        "判定類型": "精準",
        "優先級": 1,
        "需人工確認": "N",
        "判定依據": "",
        "是否通用詞": "N",
        "必須命中輔助關鍵字": "N",
    }
    for col, default_value in optional_columns.items():
        if col not in keyword_df.columns:
            keyword_df[col] = default_value
    keyword_df["啟用"] = keyword_df["啟用"].fillna("Y").astype(str).str.strip()
    keyword_df = keyword_df[keyword_df["啟用"].str.upper() == "Y"].copy()
    keyword_df["主關鍵字"] = keyword_df["主關鍵字"].astype(str).str.strip()
    keyword_df["銀行別"] = keyword_df["銀行別"].astype(str).str.strip()
    keyword_df["提及信用卡"] = keyword_df["提及信用卡"].astype(str).str.strip()
    keyword_df["卡組織"] = keyword_df["卡組織"].fillna("ALL").astype(str).str.strip()
    keyword_df["卡組織"] = keyword_df["卡組織"].replace(["", "None", "none", "nan", "NaN"], "ALL")
    keyword_df["判定類型"] = keyword_df["判定類型"].fillna("精準").astype(str).str.strip()
    keyword_df["是否通用詞"] = keyword_df["是否通用詞"].fillna("N").astype(str).str.strip()
    keyword_df["必須命中輔助關鍵字"] = keyword_df["必須命中輔助關鍵字"].fillna("N").astype(str).str.strip()
    keyword_df = keyword_df[
        (keyword_df["主關鍵字"] != "")
        & (keyword_df["主關鍵字"].str.lower() != "nan")
        & (keyword_df["提及信用卡"] != "")
        & (keyword_df["提及信用卡"].str.lower() != "nan")
    ].copy()
    return keyword_df


def get_default_org_rules():
    data = [
        {"判定關鍵字": "Mastercard", "卡組織": "MC"},
        {"判定關鍵字": "MasterCard", "卡組織": "MC"},
        {"判定關鍵字": "萬事達", "卡組織": "MC"},
        {"判定關鍵字": "鈦金", "卡組織": "MC"},
        {"判定關鍵字": "鈦金卡", "卡組織": "MC"},
        {"判定關鍵字": "世界卡", "卡組織": "MC"},
        {"判定關鍵字": "Visa", "卡組織": "VISA"},
        {"判定關鍵字": "VISA", "卡組織": "VISA"},
        {"判定關鍵字": "御璽", "卡組織": "VISA"},
        {"判定關鍵字": "御璽卡", "卡組織": "VISA"},
        {"判定關鍵字": "無限卡", "卡組織": "VISA"},
        {"判定關鍵字": "JCB", "卡組織": "JCB"},
        {"判定關鍵字": "晶緻", "卡組織": "JCB"},
        {"判定關鍵字": "晶緻卡", "卡組織": "JCB"},
        {"判定關鍵字": "極緻卡", "卡組織": "JCB"},
    ]
    return pd.DataFrame(data)


def clean_org_rules(org_df):
    org_df = org_df.loc[:, ~org_df.columns.astype(str).str.contains("Unnamed")]
    org_df = org_df.dropna(axis=1, how="all")
    org_df.columns = org_df.columns.astype(str).str.strip()
    if "啟用" in org_df.columns:
        org_df["啟用"] = org_df["啟用"].fillna("Y").astype(str).str.strip()
        org_df = org_df[org_df["啟用"].str.upper() == "Y"].copy()
    if "判定關鍵字" not in org_df.columns or "卡組織" not in org_df.columns:
        return get_default_org_rules()
    org_df["判定關鍵字"] = org_df["判定關鍵字"].astype(str).str.strip()
    org_df["卡組織"] = org_df["卡組織"].astype(str).str.strip()
    org_df = org_df[(org_df["判定關鍵字"] != "") & (org_df["判定關鍵字"].str.lower() != "nan")].copy()
    return org_df[["判定關鍵字", "卡組織"]]


REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/123.0 Safari/537.36"
}


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_html(url, timeout=12):
    if not is_external_http_url(url):
        return None, "網址格式不正確"
    try:
        response = requests.get(
            url,
            headers=REQUEST_HEADERS,
            timeout=timeout,
            verify=False
        )
        response.raise_for_status()
        if response.encoding is None or response.encoding.lower() == "iso-8859-1":
            response.encoding = response.apparent_encoding
        return response.text, None
    except Exception as e:
        return None, str(e)


def extract_sourceweb_url(report_url):
    html, error = fetch_html(report_url)
    if error:
        return None, error
    soup = BeautifulSoup(html, "lxml")
    for tr in soup.find_all("tr"):
        row_text = tr.get_text(" ", strip=True)
        if "SourceWeb" in row_text or "原始位置" in row_text:
            link = tr.find("a", href=True)
            if link:
                return urljoin(report_url, link["href"]), None
    report_domain = get_domain(report_url)
    candidates = []
    for a in soup.find_all("a", href=True):
        href = urljoin(report_url, a["href"])
        href_domain = get_domain(href)
        if not href.startswith("http") or href_domain == "":
            continue
        if report_domain and href_domain == report_domain:
            continue
        bad_domains = ["facebook.com", "line.me", "google.", "doubleclick", "instagram.com"]
        if any(bad in href_domain for bad in bad_domains):
            continue
        candidates.append(href)
    if len(candidates) > 0:
        return candidates[0], None
    return None, "找不到 SourceWeb 原始新聞網址"


def clean_article_text_lines(text, mode="strict"):
    """移除常見廣告、推薦卡片、導購與頁尾文字，降低誤抓廣告卡名。

    strict：一般新聞使用，會排除較多導購與側欄文字。
    loose：卡優信用卡推薦長文使用，避免把真正卡片介紹段落誤刪。
    """
    if not text:
        return ""

    base_ad_patterns = [
        "謹慎理財", "信用至上", "循環利率", "預借現金", "年百分率",
        "相關新聞", "延伸閱讀", "熱門新聞", "最新新聞", "看更多",
        "廣告", "ADVERTISEMENT", "Sponsored", "贊助", "分享此文",
    ]

    strict_only_patterns = [
        "立即辦卡", "線上申辦", "前往申辦", "信用卡推薦", "推薦信用卡",
        "熱門信用卡", "信用卡比較", "卡優推薦", "優惠活動", "更多信用卡",
    ]

    ad_patterns = base_ad_patterns if mode == "loose" else base_ad_patterns + strict_only_patterns

    cleaned = []
    seen = set()
    for raw_line in str(text).splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if len(line) < 8:
            continue
        line_norm = normalize_text(line)
        if any(normalize_text(p) in line_norm for p in ad_patterns):
            continue
        if line_norm in seen:
            continue
        seen.add(line_norm)
        cleaned.append(line)
    return "\n".join(cleaned).strip()


def is_cardu_recommendation_article(title, article_url="", html_text=""):
    """判斷卡優文章是否為信用卡推薦長文。這類文章不可用太嚴格的廣告排除規則。"""
    domain = get_domain(article_url)
    combined = f"{title}\n{html_text}"[:3000]
    combined_norm = normalize_text(combined)
    if "cardu.com.tw" not in domain:
        return False
    indicators = [
        "推薦信用卡", "必辦", "夯卡", "信用卡總整理", "懶人包",
        "文章目錄", "張必辦", "卡優限定", "專屬連結線上申辦",
    ]
    return any(normalize_text(term) in combined_norm for term in indicators)


def extract_text_from_candidate_container(soup, clean_mode="strict"):
    """優先從新聞正文容器抓文字，避免抓到側欄與廣告。"""
    selectors = [
        "article",
        "main",
        "#article",
        "#content",
        ".article",
        ".article-content",
        ".article_content",
        ".news-content",
        ".news_content",
        ".newsDetail",
        ".news_detail",
        ".detail",
        ".detailContent",
        ".detail_content",
        ".content",
        ".main-content",
        ".main_content",
        ".story",
        ".caas-body",
    ]
    candidates = []
    for selector in selectors:
        for node in soup.select(selector):
            text = node.get_text("\n", strip=True)
            text = clean_article_text_lines(text, mode=clean_mode)
            if len(text) >= 120:
                candidates.append(text)
    if candidates:
        return max(candidates, key=len)
    return ""


def extract_cardu_article_text(soup, article_url="", title="", html_text=""):
    """卡優新聞網專用正文萃取。

    一般新聞：嚴格排除右欄、廣告、熱門推薦。
    信用卡推薦長文：使用寬鬆模式，避免把真正的卡片介紹段落刪掉。
    """
    is_recommendation = is_cardu_recommendation_article(title, article_url, html_text)
    clean_mode = "loose" if is_recommendation else "strict"

    for tag in list(soup.find_all(True)):
        if not isinstance(tag, Tag):
            continue

        try:
            tag_id = tag.get("id", "") or ""
            tag_class = tag.get("class", []) or []

            if isinstance(tag_class, (list, tuple)):
                class_text = " ".join([str(item) for item in tag_class])
            else:
                class_text = str(tag_class)

            attrs = f"{tag_id} {class_text}".lower()
            text_preview = tag.get_text(" ", strip=True)[:120]
        except Exception:
            continue

        structural_noise = ["ad", "banner", "menu", "nav", "footer", "header", "right", "side", "recommend", "hot"]
        if any(key in attrs for key in structural_noise):
            tag.decompose()
            continue

        if not is_recommendation:
            if any(key in text_preview for key in ["熱門信用卡", "推薦信用卡", "立即辦卡", "謹慎理財", "信用至上"]):
                tag.decompose()

    text = extract_text_from_candidate_container(soup, clean_mode=clean_mode)
    if text:
        return text

    paragraphs = []
    for p in soup.find_all(["p", "td", "div", "li", "h2", "h3"]):
        if not isinstance(p, Tag):
            continue
        value = p.get_text(" ", strip=True)
        if len(value) >= 20:
            paragraphs.append(value)

    fallback_text = "\n".join(paragraphs)
    if is_recommendation and "文章目錄" in fallback_text:
        after_toc = fallback_text.split("文章目錄", 1)[1]
        if len(after_toc) > 500:
            fallback_text = "文章目錄\n" + after_toc

    return clean_article_text_lines(fallback_text, mode=clean_mode)

def extract_article_text_from_html(html, article_url="", title=""):
    if not html:
        return ""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript", "header", "footer", "nav", "iframe", "form", "aside"]):
        tag.decompose()

    domain = get_domain(article_url)
    if "cardu.com.tw" in domain:
        text = extract_cardu_article_text(soup, article_url=article_url, title=title, html_text=html)
        if len(text.strip()) >= 80:
            return text

    # 先嘗試正文容器。
    text = extract_text_from_candidate_container(soup)
    if len(text.strip()) >= 120:
        return text

    # 再嘗試 trafilatura。
    if trafilatura is not None:
        try:
            extracted = trafilatura.extract(
                str(soup),
                include_comments=False,
                include_tables=True,
                no_fallback=False,
            )
            extracted = clean_article_text_lines(extracted)
            if extracted and len(extracted.strip()) >= 80:
                return extracted.strip()
        except Exception:
            pass

    paragraphs = []
    for p in soup.find_all(["p", "article", "div"]):
        line = p.get_text(" ", strip=True)
        if len(line) >= 30:
            paragraphs.append(line)
    return clean_article_text_lines("\n".join(paragraphs))


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_article_text(article_url, article_title=""):
    html, error = fetch_html(article_url)
    if error:
        return "", error
    text = extract_article_text_from_html(html, article_url=article_url, title=article_title)
    if len(text.strip()) < 80:
        return text, "抓到的正文過短，可能需要人工貼全文"
    return text, None


def detect_org_from_context(context_text, org_rules_df):
    context_norm = normalize_text(context_text)
    for _, row in org_rules_df.iterrows():
        keyword = row["判定關鍵字"]
        org = row["卡組織"]
        if contains_term(context_norm, keyword):
            return org, keyword
    return None, None


def get_org_from_card_list(card_df, bank_name, card_name):
    matched = card_df[(card_df["銀行別"] == bank_name) & (card_df["提及信用卡"] == card_name)]
    orgs = matched["卡組織"].dropna().unique().tolist()
    orgs = [org for org in orgs if str(org).strip() not in ["", "ALL", "nan", "None"]]
    if len(orgs) == 1:
        return orgs[0]
    return "ALL"



def is_generic_like_card_name(card_name):
    """判斷卡片名稱是否太像泛稱。

    這類名稱即使出現在文章中，也必須在同一個附近段落看到銀行名稱，
    否則容易把所有銀行的同名 / 同等級卡片都抓出來。
    """
    name_norm = normalize_text(card_name)
    if not name_norm:
        return False
    if is_high_risk_generic_term(card_name):
        return True
    risky_phrases = [
        "世界商務卡", "鈦金商務卡", "鈦金商旅卡", "JCB晶緻卡",
        "現金回饋JCB卡", "一卡通聯名卡", "商旅御璽卡",
    ]
    if any(name_norm == normalize_text(item) for item in risky_phrases):
        return True
    # 很短、且主要由泛稱組成的卡名，也要求銀行脈絡。
    if len(name_norm) <= 10:
        generic_hits = [g for g in HIGH_RISK_GENERIC_TERMS if normalize_text(g) and normalize_text(g) in name_norm]
        if generic_hits:
            return True
    return False


def is_cardu_longform_detection_text(title, article_text):
    """判斷是否要使用卡優推薦長文的區塊化偵測。"""
    combined_norm = normalize_text(f"{title}\n{str(article_text)[:5000]}")
    indicators = [
        "推薦信用卡", "海外信用卡", "必辦", "夯卡", "懶人包", "總整理",
        "文章目錄", "有限定國外實體商店消費", "不限定國外實體商店",
    ]
    return any(contains_term(combined_norm, term) for term in indicators)


def is_probable_cardu_card_heading(line):
    """卡優長文中常見的一張卡標題列。"""
    line = re.sub(r"\s+", " ", str(line)).strip()
    if not line:
        return False
    line_norm = normalize_text(line)
    if len(line) > 90:
        return False
    deny_exact = {
        "信用卡", "海外回饋", "條件門檻", "有限定國外實體商店消費(信用卡)",
        "不限定國外實體商店、網路消費", "限定國家", "有上限", "無上限",
        "商務卡卡等", "文章目錄",
    }
    if line in deny_exact or line_norm in {normalize_text(x) for x in deny_exact}:
        return False
    deny_contains = [
        "https://", "http://", "imgcloud", "活動期間", "立即線上申辦", "專屬連結",
        "看更多", "新戶加碼", "卡優獨家", "謹慎理財", "信用無價",
    ]
    if any(contains_term(line_norm, item) for item in deny_contains):
        return False
    card_markers = [
        "信用卡", "聯名卡", "卡", "CUBE", "Unicard", "DAWHO", "LINE Pay", "Richart",
        "Gogoro", "JCB", "Bankee", "iLEO", "U Bear", "uniopen", "Panda", "利HIGH",
    ]
    bank_terms = []
    for bank, aliases in BANK_ALIAS_MAP.items():
        bank_terms.append(bank)
        bank_terms.extend(aliases)
    has_card_marker = any(contains_term(line_norm, marker) for marker in card_markers)
    has_bank = any(contains_term(line_norm, term) for term in bank_terms)
    # 卡優表格常以「銀行 + 卡名」當一個短行，例如「永豐幣倍卡」。
    return has_card_marker and (has_bank or line.endswith("卡") or "card" in line.lower())


def extract_cardu_card_blocks(article_text):
    """把卡優推薦長文切成一張卡一個區塊，避免整篇交叉誤判。"""
    raw_lines = [re.sub(r"\s+", " ", x).strip() for x in str(article_text).splitlines()]
    lines = [x for x in raw_lines if x]
    if not lines:
        return []

    # 先從文章目錄後開始，排除頁首重複推薦模組。
    joined = "\n".join(lines)
    if "文章目錄" in joined:
        joined = joined.split("文章目錄", 1)[1]
        lines = [re.sub(r"\s+", " ", x).strip() for x in joined.splitlines() if re.sub(r"\s+", " ", x).strip()]

    heading_indexes = []
    for idx, line in enumerate(lines):
        if is_probable_cardu_card_heading(line):
            heading_indexes.append(idx)

    # 去掉太密集或重複的 heading。
    cleaned_indexes = []
    seen_heading_norms = set()
    for idx in heading_indexes:
        norm = normalize_text(lines[idx])
        if norm in seen_heading_norms:
            continue
        seen_heading_norms.add(norm)
        cleaned_indexes.append(idx)

    if len(cleaned_indexes) < 3:
        return []

    blocks = []
    for pos, start_idx in enumerate(cleaned_indexes):
        end_idx = cleaned_indexes[pos + 1] if pos + 1 < len(cleaned_indexes) else min(len(lines), start_idx + 80)
        block_lines = lines[start_idx:end_idx]
        block = "\n".join(block_lines).strip()
        if len(block) >= 20:
            blocks.append(block)
    return blocks


def dedupe_detected_cards(detected_df):
    if detected_df is None or len(detected_df) == 0:
        return pd.DataFrame()
    detected_df = detected_df.copy()
    detected_df["確認排序"] = detected_df["需人工確認"].apply(lambda x: 1 if x == "N" else 2)
    detected_df["方式排序"] = detected_df["判定方式"].apply(lambda x: 1 if x == "正式卡名比對" else 2)
    detected_df["卡片去重鍵"] = detected_df.apply(
        lambda r: str(r["銀行別"]) + "|" + normalize_card_for_dedupe(r["提及信用卡"]),
        axis=1,
    )
    detected_df = detected_df.sort_values(by=["確認排序", "方式排序"])
    detected_df = detected_df.drop_duplicates(subset=["卡片去重鍵"], keep="first")
    return detected_df.drop(columns=["確認排序", "方式排序", "卡片去重鍵"])


def detect_cards_core(article_text, selected_title, card_df, keyword_df, org_rules_df, basis_prefix=""):
    full_text = f"{selected_title}\n{article_text}"
    full_text_norm = normalize_text(full_text)
    candidates = []

    duplicate_card_names = set(
        card_df.groupby("提及信用卡").size().loc[lambda s: s > 1].index.astype(str).tolist()
    )

    # A. 正式卡名比對
    for _, card_row in card_df.iterrows():
        bank_name = card_row["銀行別"]
        card_name = card_row["提及信用卡"]
        default_org = card_row["卡組織"]

        if is_high_risk_generic_term(card_name):
            continue

        mention_count = count_term(full_text_norm, card_name)
        if mention_count <= 0:
            continue

        contexts = get_contexts(full_text, card_name, window=120)
        requires_bank_context = (str(card_name) in duplicate_card_names) or is_generic_like_card_name(card_name)
        valid_contexts = []
        for context in contexts:
            context_norm = normalize_text(context)
            if requires_bank_context and not has_bank_context(context_norm, bank_name):
                continue
            valid_contexts.append(context)

        if len(valid_contexts) == 0:
            continue

        detected_org = None
        org_basis = None
        for context in valid_contexts:
            detected_org, org_basis = detect_org_from_context(context, org_rules_df)
            if detected_org is not None:
                break

        final_org = detected_org if detected_org is not None else default_org
        if final_org in ["", "nan", "None"]:
            final_org = "ALL"

        basis = f"正式卡名：{card_name}"
        if requires_bank_context:
            basis += "；已確認同段落銀行脈絡"
        if org_basis:
            basis += f"；卡組織：{org_basis}"
        if basis_prefix:
            basis = f"{basis_prefix}；" + basis

        candidates.append({
            "銀行別": bank_name,
            "提及信用卡": card_name,
            "卡組織": final_org,
            "全文出現次數": mention_count,
            "實際計入": 1,
            "判定方式": "正式卡名比對",
            "判定依據": basis,
            "需人工確認": "N",
            "保留": True,
        })

    # B. 關鍵字判定表比對
    if keyword_df is not None and len(keyword_df) > 0:
        for _, rule in keyword_df.iterrows():
            main_keyword = rule["主關鍵字"]
            helper_keywords = split_terms(rule.get("輔助關鍵字", ""))
            exclude_keywords = split_terms(rule.get("排除關鍵字", ""))
            bank_name = rule["銀行別"]
            card_name = rule["提及信用卡"]
            rule_org = rule.get("卡組織", "ALL")
            rule_type = str(rule.get("判定類型", "精準")).strip()
            priority = safe_int(rule.get("優先級", 1), default=1)
            need_manual = is_yes(rule.get("需人工確認", "N"))

            is_generic_term = (
                is_yes(rule.get("是否通用詞", "N"))
                or is_high_risk_generic_term(main_keyword)
                or rule_type == "通用詞"
            )

            # 重要：通用詞不再獨立產生卡片列。它只能當卡組織 / 卡等級 / 人工提醒線索。
            if is_generic_term:
                continue

            must_hit_helper = (
                is_yes(rule.get("必須命中輔助關鍵字", "N"))
                or rule_type in ["需銀行"]
            )

            if not contains_term(full_text_norm, main_keyword):
                continue

            helper_keywords = sorted(set(helper_keywords + (get_bank_alias_terms(bank_name) if must_hit_helper else [])))
            if must_hit_helper and len(helper_keywords) == 0:
                continue

            contexts = get_contexts(full_text, main_keyword, window=160)
            matched_context = None
            matched_helpers = []
            for context in contexts:
                context_norm = normalize_text(context)
                current_excludes = [term for term in exclude_keywords if contains_term(context_norm, term)]
                if len(current_excludes) > 0:
                    continue

                # 重要：輔助關鍵字只能看同一個 context，不看整篇文章，避免長文交叉誤判。
                current_helpers = [term for term in helper_keywords if contains_term(context_norm, term)]
                if must_hit_helper and len(current_helpers) == 0:
                    continue

                matched_context = context
                matched_helpers = current_helpers
                break

            if matched_context is None:
                continue

            detected_org, org_basis = detect_org_from_context(matched_context, org_rules_df)
            if detected_org is not None:
                final_org = detected_org
            elif pd.notna(rule_org) and str(rule_org).strip() not in ["", "ALL", "nan", "None"]:
                final_org = str(rule_org).strip()
            else:
                final_org = get_org_from_card_list(card_df, bank_name, card_name)

            mention_count = count_term(full_text_norm, main_keyword)
            basis_parts = [f"主關鍵字：{main_keyword}"]
            if len(matched_helpers) > 0:
                basis_parts.append(f"輔助關鍵字：{', '.join(matched_helpers)}")
            if rule_type:
                basis_parts.append(f"判定類型：{rule_type}")
            if org_basis:
                basis_parts.append(f"卡組織：{org_basis}")
            existing_basis = str(rule.get("判定依據", "")).strip()
            if existing_basis not in ["", "nan", "None"]:
                basis_parts.append(existing_basis)
            if basis_prefix:
                basis_parts.insert(0, basis_prefix)

            manual_value = "Y" if (need_manual or priority >= 2) else "N"
            keep_default = False if manual_value == "Y" else True
            candidates.append({
                "銀行別": bank_name,
                "提及信用卡": card_name,
                "卡組織": final_org,
                "全文出現次數": mention_count,
                "實際計入": 1,
                "判定方式": "關鍵字判讀",
                "判定依據": "；".join(basis_parts),
                "需人工確認": manual_value,
                "保留": keep_default,
            })

    if len(candidates) == 0:
        return pd.DataFrame()
    return dedupe_detected_cards(pd.DataFrame(candidates))


def detect_cards_from_article(article_text, selected_title, card_df, keyword_df, org_rules_df):
    """信用卡偵測主入口。

    一般新聞：直接在全文偵測。
    卡優信用卡推薦長文：先切成單卡區塊，再逐區塊偵測，避免不同段落的銀行名與通用詞互相誤配。
    """
    if is_cardu_longform_detection_text(selected_title, article_text):
        blocks = extract_cardu_card_blocks(article_text)
        if len(blocks) >= 3:
            frames = []
            for idx, block in enumerate(blocks, start=1):
                block_title = block.splitlines()[0][:60] if block.splitlines() else f"區塊{idx}"
                df = detect_cards_core(
                    article_text=block,
                    selected_title=block_title,
                    card_df=card_df,
                    keyword_df=keyword_df,
                    org_rules_df=org_rules_df,
                    basis_prefix=f"CardU長文區塊{idx}：{block_title}",
                )
                if len(df) > 0:
                    frames.append(df)
            if frames:
                return dedupe_detected_cards(pd.concat(frames, ignore_index=True))

    return detect_cards_core(article_text, selected_title, card_df, keyword_df, org_rules_df)


RESULT_COLUMNS = [
    "news_order", "card_order", "監測日期", "訊息標題", "網址", "SourceWeb",
    "銀行別", "提及信用卡", "卡組織", "全文出現次數", "實際計入", "判定方式", "判定依據",
]


def ensure_result_columns(df):
    for col in RESULT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[RESULT_COLUMNS]


def sort_result_df(result_df):
    if len(result_df) == 0:
        return result_df
    result_df = result_df.copy()
    result_df["news_order"] = pd.to_numeric(result_df["news_order"], errors="coerce").fillna(999999)
    result_df["card_order"] = pd.to_numeric(result_df["card_order"], errors="coerce").fillna(999999)
    return result_df.sort_values(by=["news_order", "card_order"]).reset_index(drop=True)


def get_next_card_order(result_df, selected_news_order):
    if len(result_df) == 0:
        return 1
    same_news = result_df[pd.to_numeric(result_df["news_order"], errors="coerce") == int(selected_news_order)]
    if len(same_news) == 0:
        return 1
    try:
        return int(same_news["card_order"].max()) + 1
    except Exception:
        return len(same_news) + 1


def remove_auto_rows_for_order(result_df, selected_news_order):
    if len(result_df) == 0:
        return result_df
    return result_df[~((pd.to_numeric(result_df["news_order"], errors="coerce") == int(selected_news_order)) & (result_df["判定方式"] != "人工補卡"))].copy()


def append_rows_to_result(result_df, rows):
    if len(rows) == 0:
        return result_df
    new_df = pd.DataFrame(rows)
    new_df = ensure_result_columns(new_df)
    result_df = ensure_result_columns(result_df)
    result_df = pd.concat([result_df, new_df], ignore_index=True)
    # 只在同一個原始新聞列內去重同一張卡；不同 raw row 即使網址或標題相同，也必須保留。
    result_df["__dedupe_key"] = result_df.apply(
        lambda r: str(r["news_order"]) + "|" + str(r["銀行別"]) + "|" + normalize_card_for_dedupe(r["提及信用卡"]) + "|" + str(r.get("卡組織", "")),
        axis=1,
    )
    result_df = result_df.drop_duplicates(subset=["__dedupe_key"], keep="first").drop(columns=["__dedupe_key"])
    return sort_result_df(result_df)


def rerun_page():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def _remove_order_from_lists(selected_news_order):
    order = int(selected_news_order)
    st.session_state.deleted_rows = [row for row in st.session_state.deleted_rows if int(row.get("news_order", -1)) != order]
    st.session_state.failed_rows = [row for row in st.session_state.failed_rows if int(row.get("news_order", -1)) != order]
    st.session_state.pending_rows = [row for row in st.session_state.pending_rows if int(row.get("news_order", -1)) != order]
    st.session_state.deleted_orders.discard(order)
    st.session_state.failed_orders.discard(order)
    st.session_state.pending_orders.discard(order)


def mark_news_as_no_card(selected_url, selected_date, selected_title, selected_news_order, reason="無信用卡，已排除"):
    order = int(selected_news_order)
    if len(st.session_state.result_df) > 0:
        st.session_state.result_df = st.session_state.result_df[pd.to_numeric(st.session_state.result_df["news_order"], errors="coerce") != order].copy()
    _remove_order_from_lists(order)
    st.session_state.deleted_orders.add(order)
    st.session_state.processed_orders.add(order)
    st.session_state.deleted_rows.append({
        "news_order": order,
        "監測日期": selected_date,
        "訊息標題": selected_title,
        "網址": selected_url,
        "SourceWeb": st.session_state.last_fetch_info.get("SourceWeb", "") if st.session_state.last_fetch_info else "",
        "處理狀態": reason,
    })
    st.session_state.detected_df = pd.DataFrame()
    st.session_state.detected_url = None
    st.session_state.detected_order = None


def reset_news_to_unprocessed(selected_url, selected_news_order=None):
    if selected_news_order is None:
        # 向下相容：若舊呼叫只傳 URL，盡量用目前選取的原始列號。
        selected_news_order = st.session_state.get("selected_news_order")
    if selected_news_order is None:
        return
    order = int(selected_news_order)
    if len(st.session_state.result_df) > 0:
        st.session_state.result_df = st.session_state.result_df[pd.to_numeric(st.session_state.result_df["news_order"], errors="coerce") != order].copy()
    st.session_state.processed_orders.discard(order)
    _remove_order_from_lists(order)
    st.session_state.detected_df = pd.DataFrame()
    st.session_state.detected_url = None
    st.session_state.detected_order = None


def mark_fetch_failed(selected_url, selected_date, selected_title, selected_news_order, reason):
    order = int(selected_news_order)
    st.session_state.failed_orders.add(order)
    st.session_state.failed_rows = [row for row in st.session_state.failed_rows if int(row.get("news_order", -1)) != order]
    st.session_state.failed_rows.append({
        "news_order": order,
        "監測日期": selected_date,
        "訊息標題": selected_title,
        "網址": selected_url,
        "SourceWeb": st.session_state.last_fetch_info.get("SourceWeb", "") if st.session_state.last_fetch_info else "",
        "失敗原因": reason,
        "全文字數": st.session_state.last_fetch_info.get("抓取字數", "") if st.session_state.last_fetch_info else "",
    })


def add_pending_rows(detected_df, selected_url, selected_date, selected_title, selected_news_order, source_url):
    if detected_df is None or len(detected_df) == 0:
        return
    order = int(selected_news_order)
    review_df = detected_df[detected_df["需人工確認"] == "Y"].copy()
    if len(review_df) == 0:
        return
    st.session_state.pending_rows = [row for row in st.session_state.pending_rows if int(row.get("news_order", -1)) != order]
    for _, row in review_df.iterrows():
        st.session_state.pending_rows.append({
            "news_order": order,
            "監測日期": selected_date,
            "訊息標題": selected_title,
            "網址": selected_url,
            "SourceWeb": source_url,
            "銀行別": row.get("銀行別", ""),
            "提及信用卡": row.get("提及信用卡", ""),
            "卡組織": row.get("卡組織", ""),
            "判定方式": row.get("判定方式", ""),
            "判定依據": row.get("判定依據", ""),
            "待確認原因": row.get("判定依據", ""),
        })
    st.session_state.pending_orders.add(order)


def get_news_status(news_order):
    order = int(news_order)
    if order in st.session_state.deleted_orders:
        return "已排除-無卡"
    if order in st.session_state.failed_orders:
        return "抓取失敗"
    if order in st.session_state.pending_orders:
        return "待確認"
    if order in st.session_state.processed_orders:
        return "已處理"
    return "未處理"


def get_status_rank(status):
    if status == "未處理":
        return 1
    if status == "抓取失敗":
        return 2
    if status == "已處理":
        return 3
    return 4


def _rename_news_order(df):
    if df is None or len(df) == 0:
        return pd.DataFrame()
    out = df.copy()
    if "news_order" in out.columns:
        out = out.rename(columns={"news_order": "原始列號"})
    return out


def build_unprocessed_df(news_df):
    handled = set(st.session_state.processed_orders) | set(st.session_state.deleted_orders) | set(st.session_state.failed_orders) | set(st.session_state.pending_orders)
    df = news_df[~news_df["news_order"].astype(int).isin(handled)].copy()
    if len(df) == 0:
        return pd.DataFrame(columns=["原始列號", "監測日期", "訊息標題", "網址", "未處理原因"])
    df["未處理原因"] = "尚未處理"
    return df.rename(columns={"news_order": "原始列號"})[["原始列號", "監測日期", "訊息標題", "網址", "未處理原因"]]


def build_tracking_df(news_df, result_df, pending_rows, deleted_rows, failed_rows):
    """原始列完整追蹤表：設計 B，一張卡/一個狀態一列。

    不做新聞層級去重；每一列 raw data 都至少會在追蹤表出現一次。
    """
    result_df = ensure_result_columns(result_df.copy()) if result_df is not None else pd.DataFrame(columns=RESULT_COLUMNS)
    pending_df = pd.DataFrame(pending_rows) if pending_rows else pd.DataFrame()
    deleted_df = pd.DataFrame(deleted_rows) if deleted_rows else pd.DataFrame()
    failed_df = pd.DataFrame(failed_rows) if failed_rows else pd.DataFrame()

    rows = []
    for _, raw in news_df.iterrows():
        order = int(raw["news_order"])
        base = {
            "原始列號": order,
            "監測日期": raw.get("監測日期", ""),
            "訊息標題": raw.get("訊息標題", ""),
            "Mastercard URL": raw.get("網址", ""),
        }
        has_any = False

        if len(result_df) > 0:
            matched = result_df[pd.to_numeric(result_df["news_order"], errors="coerce") == order]
            for _, r in matched.iterrows():
                has_any = True
                item = dict(base)
                item.update({
                    "處理狀態": "已分類",
                    "SourceWeb URL": r.get("SourceWeb", ""),
                    "銀行別": r.get("銀行別", ""),
                    "提及信用卡": r.get("提及信用卡", ""),
                    "卡組織": r.get("卡組織", ""),
                    "判定方式": r.get("判定方式", ""),
                    "判定依據": r.get("判定依據", ""),
                    "失敗原因": "",
                    "人工備註": "",
                })
                rows.append(item)

        if len(pending_df) > 0 and "news_order" in pending_df.columns:
            matched = pending_df[pd.to_numeric(pending_df["news_order"], errors="coerce") == order]
            for _, r in matched.iterrows():
                has_any = True
                item = dict(base)
                item.update({
                    "處理狀態": "待確認",
                    "SourceWeb URL": r.get("SourceWeb", ""),
                    "銀行別": r.get("銀行別", ""),
                    "提及信用卡": r.get("提及信用卡", ""),
                    "卡組織": r.get("卡組織", ""),
                    "判定方式": r.get("判定方式", ""),
                    "判定依據": r.get("判定依據", r.get("待確認原因", "")),
                    "失敗原因": "",
                    "人工備註": "",
                })
                rows.append(item)

        if len(deleted_df) > 0 and "news_order" in deleted_df.columns:
            matched = deleted_df[pd.to_numeric(deleted_df["news_order"], errors="coerce") == order]
            for _, r in matched.iterrows():
                has_any = True
                item = dict(base)
                item.update({
                    "處理狀態": "無卡排除",
                    "SourceWeb URL": r.get("SourceWeb", ""),
                    "銀行別": "",
                    "提及信用卡": "",
                    "卡組織": "",
                    "判定方式": "無卡排除",
                    "判定依據": r.get("處理狀態", ""),
                    "失敗原因": "",
                    "人工備註": "",
                })
                rows.append(item)

        if len(failed_df) > 0 and "news_order" in failed_df.columns:
            matched = failed_df[pd.to_numeric(failed_df["news_order"], errors="coerce") == order]
            for _, r in matched.iterrows():
                has_any = True
                item = dict(base)
                item.update({
                    "處理狀態": "抓取失敗",
                    "SourceWeb URL": r.get("SourceWeb", ""),
                    "銀行別": "",
                    "提及信用卡": "",
                    "卡組織": "",
                    "判定方式": "抓取失敗",
                    "判定依據": "",
                    "失敗原因": r.get("失敗原因", ""),
                    "人工備註": "",
                })
                rows.append(item)

        if not has_any:
            item = dict(base)
            item.update({
                "處理狀態": "未處理",
                "SourceWeb URL": "",
                "銀行別": "",
                "提及信用卡": "",
                "卡組織": "",
                "判定方式": "",
                "判定依據": "",
                "失敗原因": "",
                "人工備註": "",
            })
            rows.append(item)

    return pd.DataFrame(rows)


def build_summary_df(news_df, result_df, pending_rows, deleted_rows, failed_rows):
    handled = set(st.session_state.processed_orders) | set(st.session_state.deleted_orders) | set(st.session_state.failed_orders) | set(st.session_state.pending_orders)
    unprocessed = max(0, len(news_df) - len(handled))
    data = [
        ["原始新聞列數", len(news_df)],
        ["已分類原始列數", len(st.session_state.processed_orders)],
        ["待確認原始列數", len(st.session_state.pending_orders)],
        ["無卡排除原始列數", len(st.session_state.deleted_orders)],
        ["抓取失敗原始列數", len(st.session_state.failed_orders)],
        ["未處理原始列數", unprocessed],
        ["分類卡片列數", len(result_df) if result_df is not None else 0],
        ["待確認列數", len(pending_rows)],
    ]
    return pd.DataFrame(data, columns=["項目", "數量"])


def to_complete_workbook_bytes(news_df, result_df, pending_rows, deleted_rows, failed_rows):
    output = BytesIO()
    result_df = ensure_result_columns(result_df.copy()) if result_df is not None else pd.DataFrame(columns=RESULT_COLUMNS)
    pending_df = pd.DataFrame(pending_rows) if pending_rows else pd.DataFrame()
    deleted_df = pd.DataFrame(deleted_rows) if deleted_rows else pd.DataFrame()
    failed_df = pd.DataFrame(failed_rows) if failed_rows else pd.DataFrame()
    unprocessed_df = build_unprocessed_df(news_df)
    tracking_df = build_tracking_df(news_df, result_df, pending_rows, deleted_rows, failed_rows)
    summary_df = build_summary_df(news_df, result_df, pending_rows, deleted_rows, failed_rows)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        tracking_df.to_excel(writer, index=False, sheet_name="01_原始列完整追蹤表")
        _rename_news_order(result_df).to_excel(writer, index=False, sheet_name="02_完整卡別分類表")
        _rename_news_order(pending_df).to_excel(writer, index=False, sheet_name="03_待確認")
        _rename_news_order(deleted_df).to_excel(writer, index=False, sheet_name="04_無卡排除")
        _rename_news_order(failed_df).to_excel(writer, index=False, sheet_name="05_抓取失敗")
        unprocessed_df.to_excel(writer, index=False, sheet_name="06_未處理")
        summary_df.to_excel(writer, index=False, sheet_name="07_統計摘要")

    return output.getvalue()



def is_credit_card_related_title(title):
    title_norm = normalize_text(title)
    terms = [
        "信用卡", "刷卡", "卡友", "聯名卡", "御璽", "鈦金", "世界卡", "現金回饋",
        "推薦信用卡", "必辦", "夯卡", "卡優", "首刷禮", "辦卡", "回饋",
    ]
    return any(contains_term(title_norm, term) for term in terms)


def should_auto_exclude_no_card(title, article_text, source_url):
    """判斷沒有偵測到卡時，能不能安全標記為無卡。"""
    text_length = len(str(article_text).strip())
    if text_length < 500:
        return False, f"抓取全文字數過短：{text_length} 字"

    if is_credit_card_related_title(title):
        return False, "標題疑似信用卡相關，但未偵測到卡片"

    if "cardu.com.tw" in get_domain(source_url):
        if is_cardu_recommendation_article(title, source_url, article_text):
            return False, "卡優信用卡推薦長文未命中，需檢查抓文或關鍵字"

    return True, "全文完整且未偵測到信用卡"


def fetch_report_page_fallback_text(report_url, title=""):
    """當 SourceWeb 失效、404 或正文過短時，回頭抓 Mastercard 監測頁文字作為備援。"""
    if not is_external_http_url(report_url):
        return "", "沒有可用的 Mastercard 報告網址"
    html, error = fetch_html(report_url)
    if error:
        return "", error
    text = extract_article_text_from_html(html, article_url=report_url, title=title)
    return text or "", None


def merge_article_with_report_fallback(input_url, source_url, article_text, article_error, title):
    """正文過短或抓取異常時，合併 Mastercard 報告頁文字再判讀。"""
    original_text = str(article_text or "")
    note = article_error or ""
    if len(original_text.strip()) >= 500 and not article_error:
        return original_text, article_error, note
    # 只有當輸入網址可能是 Mastercard 監測頁時才嘗試。若直接貼原始網址，就不重抓。
    if not input_url or input_url == source_url:
        return original_text, article_error, note
    fallback_text, fallback_error = fetch_report_page_fallback_text(input_url, title)
    if fallback_text and len(fallback_text.strip()) > len(original_text.strip()):
        merged = (original_text + "\n\n" + fallback_text).strip() if original_text else fallback_text.strip()
        return merged, None if len(merged) >= 80 else article_error, f"{note}；已合併 Mastercard 監測頁文字"
    if fallback_error:
        return original_text, article_error, f"{note}；監測頁備援失敗：{fallback_error}"
    return original_text, article_error, note


def resolve_input_url_to_source_url(input_url):
    """支援 Mastercard 報告網址或原始新聞網址。

    先嘗試從頁面中抓 SourceWeb；若抓不到，則把輸入網址視為原始新聞網址。
    """
    if not is_external_http_url(input_url):
        return None, "網址格式不正確", ""

    source_url, source_error = extract_sourceweb_url(input_url)
    if source_url:
        return source_url, None, "Mastercard 報告頁 → SourceWeb"

    return input_url, None, f"直接使用輸入網址；未找到 SourceWeb：{source_error}"


def update_last_fetch_info(input_url, source_url, article_text, status, note=""):
    st.session_state.last_fetch_info = {
        "輸入網址": input_url,
        "SourceWeb": source_url,
        "抓取狀態": status,
        "抓取字數": len(str(article_text or "")),
        "備註": note,
        "全文預覽": str(article_text or "")[:3000],
    }

def auto_detect_and_update(
    article_text,
    source_url,
    selected_url,
    selected_date,
    selected_title,
    selected_news_order,
    card_df,
    keyword_df,
    org_rules_df,
    auto_add_high_confidence=True,
    auto_delete_no_card=True,
):
    detected_df = detect_cards_from_article(article_text, selected_title, card_df, keyword_df, org_rules_df)
    if len(detected_df) == 0:
        can_exclude, exclude_reason = should_auto_exclude_no_card(selected_title, article_text, source_url)
        if auto_delete_no_card and can_exclude:
            mark_news_as_no_card(
                selected_url,
                selected_date,
                selected_title,
                selected_news_order,
                reason="自動抓全文後未偵測到信用卡，已排除",
            )
            return "無卡排除", 0, detected_df

        order = int(selected_news_order)
        st.session_state.pending_rows = [row for row in st.session_state.pending_rows if int(row.get("news_order", -1)) != order]
        st.session_state.pending_rows.append({
            "news_order": order,
            "監測日期": selected_date,
            "訊息標題": selected_title,
            "網址": selected_url,
            "SourceWeb": source_url,
            "銀行別": "",
            "提及信用卡": "",
            "卡組織": "",
            "判定方式": "未命中保護",
            "判定依據": f"未偵測到信用卡，但未自動排除：{exclude_reason}",
            "待確認原因": exclude_reason,
        })
        st.session_state.pending_orders.add(order)
        st.session_state.detected_df = pd.DataFrame()
        st.session_state.detected_url = selected_url
        st.session_state.detected_order = int(selected_news_order)
        return "未偵測到卡，已保留人工確認", 0, detected_df
    st.session_state.detected_df = detected_df
    st.session_state.detected_url = selected_url
    st.session_state.detected_order = int(selected_news_order)
    if auto_add_high_confidence:
        high_df = detected_df[detected_df["需人工確認"] == "N"].copy()
        if len(high_df) > 0:
            st.session_state.result_df = remove_auto_rows_for_order(st.session_state.result_df, selected_news_order)
            rows_to_add = []
            card_order = 1
            for _, row in high_df.iterrows():
                rows_to_add.append({
                    "news_order": selected_news_order,
                    "card_order": card_order,
                    "監測日期": selected_date,
                    "訊息標題": selected_title,
                    "網址": selected_url,
                    "SourceWeb": source_url,
                    "銀行別": row["銀行別"],
                    "提及信用卡": row["提及信用卡"],
                    "卡組織": row["卡組織"],
                    "全文出現次數": row["全文出現次數"],
                    "實際計入": 1,
                    "判定方式": row["判定方式"],
                    "判定依據": row["判定依據"],
                })
                card_order += 1
            st.session_state.result_df = append_rows_to_result(st.session_state.result_df, rows_to_add)
            st.session_state.processed_orders.add(int(selected_news_order))
            pending_count = len(detected_df[detected_df["需人工確認"] == "Y"])
            if pending_count > 0:
                add_pending_rows(detected_df, selected_url, selected_date, selected_title, selected_news_order, source_url)
                return "已加入高可信，仍有待人工確認", len(rows_to_add), detected_df
            st.session_state.pending_orders.discard(int(selected_news_order))
            st.session_state.pending_rows = [row for row in st.session_state.pending_rows if int(row.get("news_order", -1)) != int(selected_news_order)]
            return "已自動加入", len(rows_to_add), detected_df
    add_pending_rows(detected_df, selected_url, selected_date, selected_title, selected_news_order, source_url)
    return "已偵測，待確認", len(detected_df), detected_df



# =====================================================
# v10：一頁式控制台 UI
# =====================================================

APP_VERSION = "v11：不去重新聞列｜原始列完整追蹤表｜v9易讀表支援｜v10一頁式控制台"

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.1rem; padding-bottom: 1.2rem; max-width: 1500px;}
    .v10-hero {
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 14px 18px;
        background: linear-gradient(90deg, #fff7ed 0%, #ffffff 48%, #eff6ff 100%);
        margin-bottom: 10px;
    }
    .v10-title {font-size: 1.45rem; font-weight: 800; margin: 0; color: #111827;}
    .v10-subtitle {font-size: .88rem; color: #4b5563; margin-top: 3px;}
    .status-strip {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 8px 12px;
        background: #ffffff;
        margin-bottom: 10px;
        font-size: .9rem;
        color: #374151;
    }
    .control-box {
        border: 1px solid #dbeafe;
        border-radius: 14px;
        padding: 14px;
        background: #f8fbff;
        margin-top: 8px;
        margin-bottom: 10px;
    }
    .mini-title {font-weight: 800; color: #111827; margin-bottom: 4px;}
    .muted-small {font-size:.83rem; color:#6b7280;}
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        padding: 8px 10px;
        border-radius: 12px;
    }
    div[data-testid="stExpander"] {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        background: #ffffff;
    }
    .result-chip {
        display:inline-block;
        border-radius:999px;
        padding:4px 9px;
        margin:2px 3px 2px 0;
        background:#eff6ff;
        border:1px solid #bfdbfe;
        color:#1d4ed8;
        font-size:12px;
        font-weight:600;
    }
    .result-chip-review {
        background:#fff7ed;
        border-color:#fed7aa;
        color:#c2410c;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def init_state():
    defaults = {
        "result_df": pd.DataFrame(columns=RESULT_COLUMNS),
        "processed_orders": set(),
        "deleted_orders": set(),
        "deleted_rows": [],
        "failed_rows": [],
        "failed_orders": set(),
        "pending_rows": [],
        "pending_orders": set(),
        "selected_news_order": None,
        "detected_df": pd.DataFrame(),
        "detected_url": None,
        "detected_order": None,
        "last_message": "",
        "last_fetch_info": {},
        "last_action_status": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()

if st.session_state.last_message:
    st.success(st.session_state.last_message)
    st.session_state.last_message = ""


def load_card_excel(uploaded_file):
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)
    sheet = "信用卡計算" if "信用卡計算" in xl.sheet_names else xl.sheet_names[0]
    return pd.read_excel(uploaded_file, sheet_name=sheet)


def _find_sheet_name(sheet_names, keyword, fallback_index=0):
    for sheet in sheet_names:
        if keyword in str(sheet):
            return sheet
    return sheet_names[fallback_index]


def _read_excel_with_detected_header(uploaded_file, sheet_name, required_terms):
    """支援 v9 易讀版：第一列可能是人工區/補充區/程式區，真正欄位在第二列。"""
    uploaded_file.seek(0)
    preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, nrows=8)
    header_row = 0
    for idx in range(len(preview)):
        row_values = [str(x).strip() for x in preview.iloc[idx].tolist() if str(x).strip() not in ["", "nan", "None"]]
        row_text = "|".join(row_values)
        if all(term in row_text for term in required_terms):
            header_row = idx
            break
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)


def load_keyword_excel(uploaded_file):
    """讀取關鍵字表，支援 v7 舊表與 v9 易讀版。

    v9 易讀版特點：
    - 工作表名稱可能是「② 關鍵字判定表」
    - 第 1 列可能是「人工區 / 補充區 / 程式區」
    - 真正 header 可能在第 2 列
    """
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)
    keyword_sheet = _find_sheet_name(xl.sheet_names, "關鍵字判定表")
    keyword_raw = _read_excel_with_detected_header(uploaded_file, keyword_sheet, ["主關鍵字", "提及信用卡"])

    org_rules = get_default_org_rules()
    org_sheet = None
    for sheet in xl.sheet_names:
        if "卡組織判定表" in str(sheet):
            org_sheet = sheet
            break
    if org_sheet:
        try:
            uploaded_file.seek(0)
            org_raw = _read_excel_with_detected_header(uploaded_file, org_sheet, ["判定關鍵字", "卡組織"])
            org_rules = clean_org_rules(org_raw)
        except Exception:
            org_rules = get_default_org_rules()
    return keyword_raw, org_rules


def build_news_view(news_df):
    display_df = news_df.copy()
    display_df["處理狀態"] = display_df["news_order"].apply(get_news_status)
    display_df["狀態排序"] = display_df["處理狀態"].apply(get_status_rank)
    display_df = display_df.sort_values(by=["news_order"]).reset_index(drop=True)
    return display_df


def make_news_options(display_df):
    options = []
    mapping = {}
    for _, row in display_df.iterrows():
        option = f"{int(row['news_order'])}. [{row['處理狀態']}] {row['訊息標題']}"
        options.append(option)
        mapping[option] = int(row["news_order"])
    return options, mapping


def get_selected_row(news_df, selected_order):
    row_df = news_df[news_df["news_order"] == selected_order]
    if len(row_df) == 0:
        return news_df.iloc[0]
    return row_df.iloc[0]


def add_detected_results_to_table(detected_df, selected_url, selected_date, selected_title, selected_news_order, source_url, mode="high"):
    if detected_df is None or len(detected_df) == 0:
        return 0
    if mode == "high":
        add_df = detected_df[detected_df["需人工確認"] == "N"].copy()
    elif mode == "all":
        add_df = detected_df.copy()
    else:
        add_df = detected_df.copy()
    if len(add_df) == 0:
        return 0
    st.session_state.result_df = remove_auto_rows_for_order(st.session_state.result_df, selected_news_order)
    rows = []
    card_order = 1
    for _, row in add_df.iterrows():
        rows.append({
            "news_order": selected_news_order,
            "card_order": card_order,
            "監測日期": selected_date,
            "訊息標題": selected_title,
            "網址": selected_url,
            "SourceWeb": source_url,
            "銀行別": row.get("銀行別", ""),
            "提及信用卡": row.get("提及信用卡", ""),
            "卡組織": row.get("卡組織", ""),
            "全文出現次數": row.get("全文出現次數", 1),
            "實際計入": 1,
            "判定方式": row.get("判定方式", ""),
            "判定依據": row.get("判定依據", ""),
        })
        card_order += 1
    st.session_state.result_df = append_rows_to_result(st.session_state.result_df, rows)
    st.session_state.processed_orders.add(int(selected_news_order))
    st.session_state.failed_orders.discard(int(selected_news_order))
    st.session_state.failed_rows = [r for r in st.session_state.failed_rows if int(r.get("news_order", -1)) != int(selected_news_order)]
    if mode == "all":
        st.session_state.pending_orders.discard(int(selected_news_order))
        st.session_state.pending_rows = [r for r in st.session_state.pending_rows if int(r.get("news_order", -1)) != int(selected_news_order)]
    return len(rows)


def render_compact_detected(detected_df):
    if detected_df is None or len(detected_df) == 0:
        st.info("目前沒有本次偵測結果。")
        return
    high_df = detected_df[detected_df["需人工確認"] == "N"].copy()
    review_df = detected_df[detected_df["需人工確認"] == "Y"].copy()
    left, right = st.columns([1, 1])
    with left:
        st.markdown("**高信心**")
        if len(high_df) == 0:
            st.caption("無")
        else:
            chips = "".join(
                f"<span class='result-chip'>{r['銀行別']}｜{r['提及信用卡']}</span>"
                for _, r in high_df.head(18).iterrows()
            )
            st.markdown(chips, unsafe_allow_html=True)
            if len(high_df) > 18:
                st.caption(f"另有 {len(high_df)-18} 筆，請到收納區查看。")
    with right:
        st.markdown("**待人工確認**")
        if len(review_df) == 0:
            st.caption("無")
        else:
            chips = "".join(
                f"<span class='result-chip result-chip-review'>{r['銀行別']}｜{r['提及信用卡']}</span>"
                for _, r in review_df.head(18).iterrows()
            )
            st.markdown(chips, unsafe_allow_html=True)
            if len(review_df) > 18:
                st.caption(f"另有 {len(review_df)-18} 筆，請到收納區查看。")


# Header
st.markdown(
    f"""
    <div class="v10-hero">
      <div class="v10-title">信用卡新聞分類工具</div>
      <div class="v10-subtitle">{APP_VERSION}。主畫面只保留偵測、批次、加入與下載；其他功能全部收納。</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Upload area: first visit expands automatically; after files exist it stays compact.
_pre_uploaded = all([
    st.session_state.get("v10_news_file") is not None,
    st.session_state.get("v10_card_file") is not None,
    st.session_state.get("v10_keyword_file") is not None,
])
with st.expander("管理上傳檔案", expanded=not _pre_uploaded):
    up1, up2, up3 = st.columns(3)
    with up1:
        news_file = st.file_uploader("① Mastercard raw data", type=["xlsx"], key="v10_news_file")
    with up2:
        card_file = st.file_uploader("② 信用卡清單", type=["xlsx"], key="v10_card_file")
    with up3:
        keyword_file = st.file_uploader("③ 關鍵字判定表", type=["xlsx"], key="v10_keyword_file")
    st.caption("上傳後會自動讀取。資料預覽與欄位檢查已收納到下方『資料檢查』。")

files_ready = all([news_file, card_file, keyword_file])
status_text = "｜".join([
    "raw data 已上傳" if news_file else "raw data 未上傳",
    "信用卡清單已上傳" if card_file else "信用卡清單未上傳",
    "關鍵字表已上傳" if keyword_file else "關鍵字表未上傳",
])
st.markdown(f"<div class='status-strip'><b>資料狀態：</b>{status_text}</div>", unsafe_allow_html=True)

if not files_ready:
    st.warning("請先展開『管理上傳檔案』並上傳三份檔案。")
    st.stop()

try:
    news_df = read_news_with_hyperlinks(news_file)
    card_df = clean_card_data(load_card_excel(card_file))
    raw_keyword_df, org_rules_df = load_keyword_excel(keyword_file)
    keyword_df = clean_keyword_data(raw_keyword_df)
except Exception as e:
    st.error("檔案讀取失敗，請確認工作表名稱與欄位格式。")
    st.exception(e)
    st.stop()

if len(news_df) == 0:
    st.error("raw data 沒有讀到任何新聞。")
    st.stop()

# Progress metrics
processed_count = len(st.session_state.processed_orders)
deleted_count = len(st.session_state.deleted_orders)
failed_count = len(st.session_state.failed_orders)
pending_news_count = len(st.session_state.pending_orders)
result_news_count = st.session_state.result_df["news_order"].nunique() if len(st.session_state.result_df) > 0 else 0
review_count = len(st.session_state.pending_rows)
handled_orders = set(st.session_state.processed_orders) | set(st.session_state.deleted_orders) | set(st.session_state.failed_orders) | set(st.session_state.pending_orders)
unprocessed_count = max(0, len(news_df) - len(handled_orders))

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("未處理", unprocessed_count)
m2.metric("已分類", result_news_count)
m3.metric("待確認", review_count)
m4.metric("抓取失敗", failed_count)
m5.metric("無卡排除", deleted_count)

# News selection
view_df = build_news_view(news_df)
options, option_to_order = make_news_options(view_df)
remembered_order = st.session_state.get("selected_news_order")
try:
    query_order = st.query_params.get("selected_news_order", None)
    if isinstance(query_order, list):
        query_order = query_order[0]
    if query_order is not None:
        remembered_order = int(query_order)
except Exception:
    pass

orders = list(option_to_order.values())
default_index = 0
if remembered_order in orders:
    default_index = orders.index(remembered_order)

st.markdown("<div class='control-box'>", unsafe_allow_html=True)
st.markdown("<div class='mini-title'>主操作台</div>", unsafe_allow_html=True)
row_a, row_b = st.columns([2.5, 1])
with row_a:
    selected_news = st.selectbox("選擇新聞", options, index=default_index, key="v10_selected_news")
with row_b:
    batch_size = st.selectbox("批次筆數", [5, 10, 20, 50, 100], index=2, key="v10_batch_size")

selected_news_order = option_to_order[selected_news]
st.session_state.selected_news_order = selected_news_order
try:
    st.query_params["selected_news_order"] = str(selected_news_order)
except Exception:
    pass

selected_row = get_selected_row(news_df, selected_news_order)
selected_date = selected_row["監測日期"]
selected_title = str(selected_row["訊息標題"])
selected_url = str(selected_row["網址"] or "")
selected_status = get_news_status(selected_news_order)

manual_url = st.text_input(
    "或貼上 Mastercard 監測網址 / 原始新聞網址",
    value="",
    placeholder="未貼網址時，會使用目前選取新聞的 Mastercard 報告網址。",
    key="v10_manual_url",
)
st.caption(f"目前選取：第 {selected_news_order} 筆｜{selected_status}｜{selected_title[:80]}")

b1, b2, b3, b4, b5, b6 = st.columns([1, 1, 1, 1, 1, 1])
run_detect = b1.button("一鍵偵測", use_container_width=True)
add_high = b2.button("加入高信心", use_container_width=True)
add_all = b3.button("全部加入", use_container_width=True)
mark_none = b4.button("標記無卡", use_container_width=True)
recover = b5.button("恢復未處理", use_container_width=True)
run_batch = b6.button("批次處理", use_container_width=True)
st.markdown("</div>", unsafe_allow_html=True)

# Actions
if mark_none:
    mark_news_as_no_card(selected_url, selected_date, selected_title, selected_news_order, reason="使用者手動標記無卡")
    st.session_state.last_message = "已標記為無卡新聞。"
    rerun_page()

if recover:
    reset_news_to_unprocessed(selected_url, selected_news_order)
    st.session_state.last_message = "已恢復為未處理。"
    rerun_page()

if run_detect:
    input_url = manual_url.strip() or selected_url
    if not input_url:
        st.error("目前新聞沒有可用網址，請手動貼上網址。")
    else:
        with st.spinner("正在抓取網址、擷取全文並偵測信用卡..."):
            source_url, source_error, source_note = resolve_input_url_to_source_url(input_url)
            if source_error:
                update_last_fetch_info(input_url, "", "", "失敗", source_error)
                st.error(source_error)
            else:
                article_text, article_error = fetch_article_text(source_url, selected_title)
                article_text, article_error, fallback_note = merge_article_with_report_fallback(
                    input_url, source_url, article_text, article_error, selected_title
                )
                fetch_status = "成功" if not article_error else "疑似不完整 / 需檢查"
                update_last_fetch_info(input_url, source_url, article_text, fetch_status, fallback_note or source_note)
                if article_error and len(str(article_text).strip()) < 80:
                    mark_fetch_failed(selected_url, selected_date, selected_title, selected_news_order, article_error)
                    st.session_state.last_message = "抓取全文失敗或過短，已放入抓取失敗區。"
                else:
                    status, added_count, detected_df = auto_detect_and_update(
                        article_text,
                        source_url,
                        selected_url,
                        selected_date,
                        selected_title,
                        selected_news_order,
                        card_df,
                        keyword_df,
                        org_rules_df,
                        auto_add_high_confidence=True,
                        auto_delete_no_card=True,
                    )
                    st.session_state.last_message = f"一鍵偵測完成：{status}；自動加入 {added_count} 筆。"
        rerun_page()

if add_high:
    if st.session_state.detected_order != selected_news_order or len(st.session_state.detected_df) == 0:
        st.warning("目前沒有可加入的本篇偵測結果。")
    else:
        source_url = st.session_state.last_fetch_info.get("SourceWeb", "") if st.session_state.last_fetch_info else ""
        added = add_detected_results_to_table(
            st.session_state.detected_df,
            selected_url,
            selected_date,
            selected_title,
            selected_news_order,
            source_url,
            mode="high",
        )
        st.session_state.last_message = f"已加入高信心結果 {added} 筆。"
        rerun_page()

if add_all:
    if st.session_state.detected_order != selected_news_order or len(st.session_state.detected_df) == 0:
        st.warning("目前沒有可加入的本篇偵測結果。")
    else:
        source_url = st.session_state.last_fetch_info.get("SourceWeb", "") if st.session_state.last_fetch_info else ""
        added = add_detected_results_to_table(
            st.session_state.detected_df,
            selected_url,
            selected_date,
            selected_title,
            selected_news_order,
            source_url,
            mode="all",
        )
        st.session_state.last_message = f"已加入本次全部偵測結果 {added} 筆。"
        rerun_page()

if run_batch:
    candidate_df = news_df[
        (news_df["news_order"] >= selected_news_order)
        & ~news_df["news_order"].astype(int).isin(st.session_state.processed_orders)
        & ~news_df["news_order"].astype(int).isin(st.session_state.deleted_orders)
        & ~news_df["news_order"].astype(int).isin(st.session_state.failed_orders)
        & ~news_df["news_order"].astype(int).isin(st.session_state.pending_orders)
    ].copy()
    success_count = 0
    skipped_count = 0
    scanned_count = 0
    progress_bar = st.progress(0)
    status_box = st.empty()
    with st.spinner("批次處理中..."):
        for _, row in candidate_df.iterrows():
            if success_count >= int(batch_size):
                break
            scanned_count += 1
            row_url = str(row["網址"] or "")
            row_title = str(row["訊息標題"])
            row_date = row["監測日期"]
            row_order = int(row["news_order"])
            status_box.write(f"處理第 {row_order} 筆：{row_title[:60]}")
            if not row_url:
                mark_fetch_failed(row_url, row_date, row_title, row_order, "沒有 Mastercard 報告網址")
                skipped_count += 1
                continue
            source_url, source_error, source_note = resolve_input_url_to_source_url(row_url)
            if source_error or not source_url:
                mark_fetch_failed(row_url, row_date, row_title, row_order, source_error or "找不到 SourceWeb")
                skipped_count += 1
                continue
            article_text, article_error = fetch_article_text(source_url, row_title)
            article_text, article_error, fallback_note = merge_article_with_report_fallback(
                row_url, source_url, article_text, article_error, row_title
            )
            update_last_fetch_info(row_url, source_url, article_text, "批次處理", fallback_note or source_note)
            if article_error and len(str(article_text).strip()) < 80:
                mark_fetch_failed(row_url, row_date, row_title, row_order, article_error)
                skipped_count += 1
                continue
            status, added_count, detected_df = auto_detect_and_update(
                article_text,
                source_url,
                row_url,
                row_date,
                row_title,
                row_order,
                card_df,
                keyword_df,
                org_rules_df,
                auto_add_high_confidence=True,
                auto_delete_no_card=True,
            )
            success_count += 1
            progress_bar.progress(min(success_count / int(batch_size), 1.0))
            time.sleep(0.05)
    status_box.empty()
    st.session_state.last_message = f"批次完成：成功處理 {success_count} 筆；跳過抓取失敗 {skipped_count} 筆；掃描 {scanned_count} 筆。"
    rerun_page()

# Current result, compact above fold
st.markdown("### 本次偵測結果")
if st.session_state.detected_order == selected_news_order:
    render_compact_detected(st.session_state.detected_df)
else:
    st.info("目前選取新聞尚無本次偵測結果。按『一鍵偵測』開始。")

# Export bar stays visible and compact
st.markdown("### 下載")
complete_workbook = to_complete_workbook_bytes(
    news_df,
    st.session_state.result_df,
    st.session_state.pending_rows,
    st.session_state.deleted_rows,
    st.session_state.failed_rows,
)
st.download_button(
    "下載完整分類工作簿",
    data=complete_workbook,
    file_name="信用卡新聞分類完整工作簿.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
st.caption("工作簿包含：原始列完整追蹤表、完整卡別分類表、待確認、無卡排除、抓取失敗、未處理、統計摘要。")

# Collapsed tools
st.markdown("### 收納功能")
tab_review, tab_failed, tab_no_card, tab_manual, tab_preview, tab_check, tab_result = st.tabs([
    "待確認", "抓取失敗", "無卡排除", "手動補卡", "全文預覽", "資料檢查", "完整總表"
])

with tab_review:
    st.caption("所有待確認資料會自動流入這裡。勾選後可加入分類總表；未勾選會保留在待確認清單。")
    if len(st.session_state.pending_rows) > 0:
        pending_df = pd.DataFrame(st.session_state.pending_rows).copy()
        pending_df["保留"] = True
        edited_review = st.data_editor(pending_df, use_container_width=True, hide_index=True, key="v11_review_editor")
        if st.button("加入勾選的待確認結果", use_container_width=True):
            keep_df = edited_review[edited_review["保留"] == True].drop(columns=["保留"], errors="ignore")
            added_total = 0
            for order, group in keep_df.groupby("news_order"):
                raw_row = get_selected_row(news_df, int(order))
                group_as_detected = group.rename(columns={
                    "待確認原因": "判定依據",
                }).copy()
                # 補齊 add_detected_results_to_table 需要的欄位。
                for col in ["全文出現次數", "實際計入", "需人工確認"]:
                    if col not in group_as_detected.columns:
                        group_as_detected[col] = 1 if col != "需人工確認" else "Y"
                added_total += add_detected_results_to_table(
                    group_as_detected,
                    str(raw_row.get("網址", "") or ""),
                    raw_row.get("監測日期", ""),
                    str(raw_row.get("訊息標題", "")),
                    int(order),
                    str(group.iloc[0].get("SourceWeb", "") or ""),
                    mode="all",
                )
            st.session_state.last_message = f"已加入待確認勾選結果 {added_total} 筆。"
            rerun_page()
    else:
        st.info("目前沒有待人工確認資料。")

with tab_failed:
    if len(st.session_state.failed_rows) > 0:
        failed_df = pd.DataFrame(st.session_state.failed_rows)
        st.dataframe(failed_df, use_container_width=True)
        with st.expander("人工貼全文補救", expanded=False):
            article_text = st.text_area("貼上新聞全文", height=220, key="v10_manual_full_text")
            if st.button("用人工全文重新偵測目前選取新聞", use_container_width=True):
                if not article_text.strip():
                    st.warning("請先貼上新聞全文。")
                else:
                    source_url = st.session_state.last_fetch_info.get("SourceWeb", "") if st.session_state.last_fetch_info else ""
                    status, added_count, detected_df = auto_detect_and_update(
                        article_text,
                        source_url,
                        selected_url,
                        selected_date,
                        selected_title,
                        selected_news_order,
                        card_df,
                        keyword_df,
                        org_rules_df,
                        auto_add_high_confidence=False,
                        auto_delete_no_card=False,
                    )
                    update_last_fetch_info("人工貼全文", source_url, article_text, "人工全文", "")
                    st.session_state.last_message = f"人工全文偵測完成：{status}。"
                    rerun_page()
    else:
        st.info("目前沒有抓取失敗紀錄。")

with tab_no_card:
    if len(st.session_state.deleted_rows) > 0:
        st.dataframe(pd.DataFrame(st.session_state.deleted_rows), use_container_width=True)
    else:
        st.info("目前沒有無卡排除紀錄。")

with tab_manual:
    st.caption("系統漏抓時才使用。")
    c1, c2, c3 = st.columns(3)
    with c1:
        manual_bank = st.selectbox("銀行", sorted(card_df["銀行別"].dropna().unique()), key="v10_manual_bank")
    manual_card_options = card_df[card_df["銀行別"] == manual_bank]["提及信用卡"].dropna().unique()
    with c2:
        manual_card = st.selectbox("卡片", sorted(manual_card_options), key="v10_manual_card")
    org_options = card_df[(card_df["銀行別"] == manual_bank) & (card_df["提及信用卡"] == manual_card)]["卡組織"].dropna().unique()
    if len(org_options) == 0:
        org_options = ["ALL"]
    with c3:
        manual_org = st.selectbox("卡組織", sorted(org_options), key="v10_manual_org")
    if st.button("手動加入目前新聞", use_container_width=True):
        already_exists = (
            len(st.session_state.result_df) > 0
            and len(st.session_state.result_df[
                (st.session_state.result_df["網址"] == selected_url)
                & (st.session_state.result_df["銀行別"] == manual_bank)
                & (st.session_state.result_df["提及信用卡"] == manual_card)
            ]) > 0
        )
        if already_exists:
            st.warning("這篇新聞已經有這張卡，系統不會重複加入。")
        else:
            manual_row = {
                "news_order": selected_news_order,
                "card_order": get_next_card_order(st.session_state.result_df, selected_news_order),
                "監測日期": selected_date,
                "訊息標題": selected_title,
                "網址": selected_url,
                "SourceWeb": st.session_state.last_fetch_info.get("SourceWeb", "") if st.session_state.last_fetch_info else "",
                "銀行別": manual_bank,
                "提及信用卡": manual_card,
                "卡組織": manual_org,
                "全文出現次數": 1,
                "實際計入": 1,
                "判定方式": "人工補卡",
                "判定依據": "使用者手動下拉選擇",
            }
            st.session_state.result_df = append_rows_to_result(st.session_state.result_df, [manual_row])
            st.session_state.processed_orders.add(int(selected_news_order))
            st.session_state.last_message = "已手動加入這張卡。"
            rerun_page()

with tab_preview:
    info = st.session_state.last_fetch_info or {}
    if info:
        st.write(f"輸入網址：{info.get('輸入網址', '')}")
        st.write(f"SourceWeb：{info.get('SourceWeb', '')}")
        st.write(f"抓取狀態：{info.get('抓取狀態', '')}")
        st.write(f"抓取字數：{info.get('抓取字數', 0)}")
        if info.get("備註"):
            st.write(f"備註：{info.get('備註')}")
        st.text_area("全文前 3000 字預覽", info.get("全文預覽", ""), height=260)
    else:
        st.info("尚未抓取任何全文。")

with tab_check:
    st.caption("除錯用。平常不用打開。")
    with st.expander("raw data 預覽", expanded=False):
        st.dataframe(news_df[["news_order", "監測日期", "訊息標題", "網址"]].head(20), use_container_width=True)
    with st.expander("信用卡清單預覽", expanded=False):
        st.dataframe(card_df.head(20), use_container_width=True)
    with st.expander("關鍵字判定表預覽", expanded=False):
        st.dataframe(keyword_df.head(20), use_container_width=True)
    with st.expander("卡組織判定表預覽", expanded=False):
        st.dataframe(org_rules_df.head(20), use_container_width=True)

with tab_result:
    st.session_state.result_df = sort_result_df(ensure_result_columns(st.session_state.result_df.copy()))
    if len(st.session_state.result_df) == 0:
        st.info("目前分類總表沒有資料。")
    else:
        edit_mode = st.toggle("開啟編輯模式", value=False, key="v10_edit_mode")
        if edit_mode:
            edited_result = st.data_editor(st.session_state.result_df, use_container_width=True, num_rows="dynamic", hide_index=True)
            if st.button("儲存分類總表修改", use_container_width=True):
                st.session_state.result_df = sort_result_df(ensure_result_columns(edited_result.copy()))
                st.session_state.last_message = "分類總表修改已儲存。"
                rerun_page()
        else:
            st.dataframe(st.session_state.result_df.tail(50), use_container_width=True)
