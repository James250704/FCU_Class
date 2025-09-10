import os, re, json, pickle, time, requests, ddddocr
from requests.adapters import HTTPAdapter
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from bs4 import BeautifulSoup
from configparser import ConfigParser
from io import BytesIO
from lxml import html as lxml_html

BASE = "https://course.fcu.edu.tw"
COOKIE_FILE = Path("cookies.pkl")
SESSION_META = Path("session.json")

ENABLE_FILE_DUMP = False  # 是否啟用網頁內容落檔功能
RE_SPACE = re.compile(r"\s+")

X_COURSE_NAME = "string(//table[@id='ctl00_MainContent_TabContainer1_tabSelected_gvToAdd']//td[contains(@class,'gvAddWithdrawCellThree')][1])"
X_MSG = "string(//span[@id='ctl00_MainContent_TabContainer1_tabSelected_lblMsgBlock'])"
# 假設餘額資訊在表格的特定位置，根據頁面結構調整此 XPath
# 例如，如果餘額在第 4 個 td，可以改為 td[4]；這裡假設為 gvAddWithdrawCellQuota 或類似
X_QUOTA = "string(//table[@id='ctl00_MainContent_TabContainer1_tabSelected_gvToAdd']//tr[2]/td[4])"  # 請根據實際頁面調整 XPath


def text_xpath(page_text: str, xpath: str, default="") -> str:
    """
    用 XPath 直接取文字，並自動標準化空白。
    支援 string(...) 或節點。
    """
    try:
        tree = lxml_html.fromstring(page_text)
        val = tree.xpath(xpath)
        if isinstance(val, list):
            val = val[0] if val else ""

        # 直接使用 re.sub 替換多餘的空白，並移除前後空白
        normalized_val = RE_SPACE.sub(" ", str(val)).strip()

        return normalized_val or default
    except Exception:
        return default


def _parse_tb_ids(raw: str) -> list[str]:
    """支援逗號/空白/換行或 JSON 陣列，回傳去重後的有序清單"""
    raw = (raw or "").strip()
    if not raw:
        return []
    ids: list[str] = []
    if raw.startswith("["):
        try:
            arr = json.loads(raw)
            ids = [str(x).strip() for x in arr if str(x).strip()]
        except Exception:
            pass
    if not ids:
        # 以逗號、空白、換行切分
        tokens = re.split(r"[,\s]+", raw)
        ids = [t.strip() for t in tokens if t.strip()]
    # 去重保序
    seen = set()
    dedup: list[str] = []
    for x in ids:
        if x not in seen:
            seen.add(x)
            dedup.append(x)
    return dedup


def load_config(path: str | Path = "config.ini"):
    """讀取 config.ini 取得 NID、PASS、tbSubIDs(list) 以及重試設定"""
    cfg = ConfigParser()
    default_path = (
        Path(__file__).with_name("config.ini")
        if "__file__" in globals()
        else Path(path)
    )
    target = default_path if default_path.exists() else Path(path)
    if not cfg.read(target, encoding="utf-8"):
        raise FileNotFoundError(f"找不到設定檔：{target.resolve()}")
    try:
        nid = cfg.get("auth", "NID").strip()
        pwd = cfg.get("auth", "PASS").strip()
        tb_raw = ""
        if cfg.has_option("course", "tbSubIDs"):
            tb_raw = cfg.get("course", "tbSubIDs")
        elif cfg.has_option("course", "tbSubID"):
            tb_raw = cfg.get("course", "tbSubID")  # 仍相容單值或多值字串
        tb_ids = _parse_tb_ids(tb_raw)

        # 讀取重試設定
        retry_enabled = cfg.getboolean("retry", "enabled", fallback=False)
        retry_count = cfg.getint("retry", "count", fallback=3)
        retry_interval = cfg.getint("retry", "interval", fallback=30)

        if not nid or not pwd or not tb_ids:
            raise ValueError("NID / PASS / tbSubIDs 不能為空")
        return nid, pwd, tb_ids, retry_enabled, retry_count, retry_interval
    except Exception as e:
        raise ValueError(f"設定檔內容不完整或格式錯誤：{e}")


def make_session():
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "zh-TW,zh;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Sec-Ch-Ua": '"Chromium";v="139", "Not;A=Brand";v="99"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"macOS"',
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
        "Connection": "keep-alive",
        "Origin": BASE,
        "Referer": f"{BASE}/",
        "Cache-Control": "max-age=0",
        "Content-Type": "application/x-www-form-urlencoded",
        "Priority": "u=0, i",
    }

    session = requests.Session()
    adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=3)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(headers)
    return session


def load_cookies_if_any(session: requests.Session) -> bool:
    if not COOKIE_FILE.exists():
        return False
    try:
        with open(COOKIE_FILE, "rb") as f:
            session.cookies = pickle.load(f)
        return True
    except Exception:
        return False


def load_session_meta():
    if not SESSION_META.exists():
        return None
    try:
        return json.loads(SESSION_META.read_text(encoding="utf-8"))
    except Exception:
        return None


def save_session_meta(guid: str, lang: str, base: str):
    SESSION_META.write_text(
        json.dumps({"guid": guid, "lang": lang, "base": base}, ensure_ascii=False),
        encoding="utf-8",
    )


def is_login_page(html: str) -> bool:
    return ('id="ctl00_Login1_UserName"' in html) or ("Login.aspx" in html)


def is_session_timeout(html: str) -> bool:
    return (
        ("Session 已逾時" in html)
        or ("請重新登入" in html)
        or ("error.aspx?code" in html)
    )


def validate_session(
    session: requests.Session, guid: str, lang: str, base: str
) -> bool:
    """
    驗證當前會話是否仍然有效，使用 lxml 進行快速解析。
    """
    url = f"{base}/AddWithdraw.aspx?guid={guid}&lang={lang}"
    try:
        # 使用一個較短的逾時時間來快速判斷連線問題
        resp = session.get(url, allow_redirects=True, timeout=10)
        resp.raise_for_status()  # 檢查 HTTP 狀態碼

        text = resp.text
        if is_login_page(text) or is_session_timeout(text):
            return False

        # 使用 lxml 進行快速解析
        tree = lxml_html.fromstring(text)

        # 檢查頁面是否包含必要的隱藏欄位和查詢按鈕
        has_viewstate = bool(tree.xpath('//input[@name="__VIEWSTATE"]'))
        has_query_btn = bool(
            tree.xpath(
                '//input[@name="ctl00$MainContent$TabContainer1$tabSelected$btnGetSub"]'
            )
        )

        return has_viewstate and has_query_btn

    except requests.RequestException:
        # 如果請求失敗（例如連線逾時），則視為無效會話
        return False


OCR_ENGINE = ddddocr.DdddOcr()


def do_login(session: requests.Session, nid: str, pwd: str, OCR_ENGINE=OCR_ENGINE):
    session.get(f"{BASE}/")
    cap = session.get(f"{BASE}/validateCode.aspx")
    ocr = OCR_ENGINE
    captcha = ocr.classification(cap.content)  # 直接處理，不存檔
    print("自動識別驗證碼:", captcha)

    r = session.get(f"{BASE}/Login.aspx")
    tree = lxml_html.fromstring(r.text)
    viewstate = tree.xpath('//input[@name="__VIEWSTATE"]/@value')[0]
    viewstategenerator = tree.xpath('//input[@name="__VIEWSTATEGENERATOR"]/@value')[0]
    eventvalidation = tree.xpath('//input[@name="__EVENTVALIDATION"]/@value')[0]

    login_data = {
        "__EVENTTARGET": "ctl00$Login1$LoginButton",
        "__EVENTARGUMENT": "",
        "__LASTFOCUS": "",
        "__VIEWSTATE": viewstate,
        "__VIEWSTATEGENERATOR": viewstategenerator,
        "__VIEWSTATEENCRYPTED": "",
        "__EVENTVALIDATION": eventvalidation,
        "ctl00$Login1$RadioButtonList1": "zh-tw",
        "ctl00$Login1$UserName": nid,
        "ctl00$Login1$Password": pwd,
        "ctl00$Login1$vcode": captcha,
    }
    resp = session.post(f"{BASE}/Login.aspx", data=login_data, allow_redirects=True)
    print("登入後跳轉URL:", resp.url)
    if resp.url.endswith("Login.aspx") or is_login_page(resp.text):
        raise RuntimeError("登入失敗，可能是驗證碼或帳密錯誤")

    parsed = urlparse(resp.url)
    base_after = f"{parsed.scheme}://{parsed.netloc}"
    guid = parse_qs(parsed.query).get("guid", [None])[0]
    lang = parse_qs(parsed.query).get("lang", [None])[0]

    if not guid or not lang:
        test = session.get(f"{base_after}/AddWithdraw.aspx", allow_redirects=True)
        p2 = urlparse(test.url)
        guid = parse_qs(p2.query).get("guid", [guid])[0]
        lang = parse_qs(p2.query).get("lang", [lang])[0]

    if not guid or not lang:
        raise RuntimeError("登入成功但無 guid/lang，流程中止")

    with open(COOKIE_FILE, "wb") as f:
        pickle.dump(session.cookies, f)
    save_session_meta(guid, lang, base_after)
    print("登入後 cookies 已儲存到 cookies.pkl，guid/lang/base 已寫入 session.json")
    return guid, lang, base_after


def get_hidden_fields_fast(page_text):
    tree = lxml_html.fromstring(page_text)
    vs = tree.xpath('//input[@name="__VIEWSTATE"]/@value')
    vg = tree.xpath('//input[@name="__VIEWSTATEGENERATOR"]/@value')
    ev = tree.xpath('//input[@name="__EVENTVALIDATION"]/@value')
    if not (vs and vg and ev):
        raise RuntimeError("頁面缺少必要隱藏欄位")
    return vs[0], vg[0], ev[0]


def find_add_event_args(html: str) -> list[str]:
    """解析頁面中所有 addCourse$N，回傳如 ['addCourse$0','addCourse$1', ...]，依序且去重"""
    nums = re.findall(r"addCourse\$(\d+)", html)
    seen = set()
    ordered = []
    for n in nums:
        arg = f"addCourse${n}"
        if arg not in seen:
            seen.add(arg)
            ordered.append(arg)
    return ordered


def parse_quota_info(quota_info: str) -> int:
    """解析 quota_info 字符串，提取剩餘名額。
    假設格式：'剩餘名額/開放名額：X  /Y'
    回傳 X，如果解析失敗回傳 0
    """
    try:
        match = re.search(r"剩餘名額/開放名額：(\d+)\s*/\d+", quota_info)
        if match:
            return int(match.group(1))
        return 0
    except Exception:
        return 0


def query_course_quota(session, add_withdraw_url, sub_id, vs, vg, ev):
    """
    單獨函數：查詢課程並查詢其餘額。
    回傳 (courseName, quota_info, quota_msg, new_vs, new_vg, new_ev, quota_html)
    如果失敗，quota_info 為 "未知"
    """
    # 🔍 查詢該科
    query_data = {
        "ctl00_ToolkitScriptManager1_HiddenField": "",
        "ctl00_MainContent_TabContainer1_ClientState": '{"ActiveTabIndex":1,"TabState":[true,true]}',
        "__EVENTTARGET": "",
        "__EVENTARGUMENT": "",
        "__LASTFOCUS": "",
        "__VIEWSTATE": vs,
        "__VIEWSTATEGENERATOR": vg,
        "__VIEWSTATEENCRYPTED": "",
        "__EVENTVALIDATION": ev,
        "ctl00$MainContent$TabContainer1$tabSelected$tbSubID": sub_id,
        "ctl00$MainContent$TabContainer1$tabSelected$btnGetSub": "查詢",
        "ctl00$MainContent$TabContainer1$tabSelected$cpeWishList_ClientState": "false",
    }
    r = session.post(add_withdraw_url, data=query_data)
    courseName = text_xpath(r.text, X_COURSE_NAME)

    if is_session_timeout(r.text) or is_login_page(r.text):
        raise RuntimeError("會話失效，需要重新登入")

    # 更新隱藏欄位
    vs, vg, ev = get_hidden_fields_fast(r.text)

    # 查詢餘額（假設查詢後該課程為第一個選項，使用 selquota$0）
    quota_data = {
        "ctl00_ToolkitScriptManager1_HiddenField": "",
        "ctl00_MainContent_TabContainer1_ClientState": '{"ActiveTabIndex":1,"TabState":[true,true]}',
        "__EVENTTARGET": "ctl00$MainContent$TabContainer1$tabSelected$gvToAdd",
        "__EVENTARGUMENT": "selquota$0",  # 假設第一個選項
        "__LASTFOCUS": "",
        "__VIEWSTATE": vs,
        "__VIEWSTATEGENERATOR": vg,
        "__VIEWSTATEENCRYPTED": "",
        "__EVENTVALIDATION": ev,
        "ctl00$MainContent$TabContainer1$tabSelected$tbSubID": sub_id,
        "ctl00$MainContent$TabContainer1$tabSelected$cpeWishList_ClientState": "false",
    }
    quota_r = session.post(add_withdraw_url, data=quota_data)
    quota_msg = text_xpath(quota_r.text, X_MSG)

    # 提取 alert 資訊
    alert_pattern = re.compile(r"alert\s*\(\s*['\"](.*?)['\"]\s*\)", re.DOTALL)
    alert_match = alert_pattern.search(quota_r.text)
    quota_info = alert_match.group(1).strip() if alert_match else "未知"

    # 更新隱藏欄位
    new_vs, new_vg, new_ev = get_hidden_fields_fast(quota_r.text)

    return courseName, quota_info, quota_msg, new_vs, new_vg, new_ev, quota_r.text


def main(stop_check_func=None):
    config_result = load_config()
    NID, PASS, TB_SUB_IDS, RETRY_ENABLED, RETRY_COUNT, RETRY_INTERVAL = config_result
    session = make_session()

    have_cookies = load_cookies_if_any(session)
    meta = load_session_meta() or {}
    guid, lang, base = meta.get("guid"), meta.get("lang"), meta.get("base")

    if (
        have_cookies
        and guid
        and lang
        and base
        and validate_session(session, guid, lang, base)
    ):
        print("✅ 既有 cookies 有效，直接使用現有登入狀態")
    else:
        print("⚠️ 既有 cookies 不可用或缺少 guid/lang/base，執行一般登入")
        guid, lang, base = do_login(session, NID, PASS)

    add_withdraw_url = f"{base}/AddWithdraw.aspx?guid={guid}&lang={lang}"

    # 如果啟用重試，則進行多輪重試
    if RETRY_ENABLED:
        if RETRY_COUNT == 0:
            print(f"✅ 啟用無限重試功能，每次間隔 {RETRY_INTERVAL} 秒")
        else:
            print(
                f"✅ 啟用自動重試功能，將重試 {RETRY_COUNT} 次，每次間隔 {RETRY_INTERVAL} 秒"
            )

        retry_round = 0
        while True:
            if stop_check_func and stop_check_func():
                print("⚠️ 收到停止信號，中斷重試")
                break

            retry_round += 1
            if RETRY_COUNT == 0:
                print(f"\n===== 第 {retry_round} 輪重試 （無限重試模式）=====")
            else:
                print(f"\n===== 第 {retry_round} 輪重試 =====")

            all_success, need_relogin = process_course_selection(
                session, add_withdraw_url, TB_SUB_IDS, stop_check_func
            )
            if need_relogin:
                print("🔄 偵測到『系統偵測異常』，執行重新登入...")
                session = make_session()
                try:
                    guid, lang, base = do_login(session, NID, PASS)
                    add_withdraw_url = (
                        f"{base}/AddWithdraw.aspx?guid={guid}&lang={lang}"
                    )
                except Exception as e:
                    print(f"❌ 重新登入失敗：{e}")
                    break
                if RETRY_INTERVAL > 0:
                    print(f"⏳ 等待 {RETRY_INTERVAL} 秒後再次嘗試...")
                    for i in range(RETRY_INTERVAL):
                        if stop_check_func and stop_check_func():
                            print("⚠️ 收到停止信號，中斷等待")
                            return
                        time.sleep(1)
                continue

            if all_success:
                print(f"🎉 所有課程選課成功！")
                break

            # 檢查是否達到重試次數限制（0 表示無限重試）
            if RETRY_COUNT > 0 and retry_round >= RETRY_COUNT:
                print("❌ 重試次數已達上限")
                break

            # 等待間隔時間
            if RETRY_INTERVAL > 0:
                print(f"⏳ 等待 {RETRY_INTERVAL} 秒後進行下一輪重試...")
                for i in range(RETRY_INTERVAL):
                    if stop_check_func and stop_check_func():
                        print("⚠️ 收到停止信號，中斷等待")
                        return
                    time.sleep(1)
            else:
                # 間隔為 0 秒，但仍需短暫延遲避免過快重試
                if stop_check_func and stop_check_func():
                    print("⚠️ 收到停止信號，中斷重試")
                    return
                time.sleep(0.1)  # 100ms 的最小延遲
    else:
        all_success, need_relogin = process_course_selection(
            session, add_withdraw_url, TB_SUB_IDS, stop_check_func
        )
        if need_relogin:
            print("🔄 偵測到『系統偵測異常』，重新登入後再嘗試一次...")
            session = make_session()
            try:
                guid, lang, base = do_login(session, NID, PASS)
                add_withdraw_url = f"{base}/AddWithdraw.aspx?guid={guid}&lang={lang}"
                process_course_selection(
                    session, add_withdraw_url, TB_SUB_IDS, stop_check_func
                )
            except Exception as e:
                print(f"❌ 重新登入失敗：{e}")

    print("\n===== 選課結束 =====")


def process_course_selection(
    session, add_withdraw_url, TB_SUB_IDS, stop_check_func=None
):
    """處理課程選課
    回傳 (all_success, need_relogin)
    need_relogin: 是否因『系統偵測異常』需要重新登入
    """
    all_success = True
    need_relogin = False

    # 🚀 第一次 GET AddWithdraw.aspx，拿初始隱藏欄位
    r = session.get(add_withdraw_url, allow_redirects=True)
    if is_session_timeout(r.text) or is_login_page(r.text):
        print("⚠️ 初始會話失效，需要重新登入")
        return False, False

    vs, vg, ev = get_hidden_fields_fast(r.text)

    # 逐科處理
    for idx, sub_id in enumerate(TB_SUB_IDS, start=1):
        if stop_check_func and stop_check_func():
            print("⚠️ 收到停止信號，停止選課")
            return False, False

        success = False
        while not success:
            try:
                # 查詢課程並查詢餘額
                courseName, quota_info, quota_msg, vs, vg, ev, quota_html = (
                    query_course_quota(session, add_withdraw_url, sub_id, vs, vg, ev)
                )
                # print(
                #     f'ℹ️ 第 {idx} 科: {sub_id} {courseName} 餘額查詢: "{quota_info}" (訊息: {quota_msg})'
                # )

                # 檢查是否有空位
                remaining = parse_quota_info(quota_info)
                if remaining <= 0:
                    print(
                        f"❌ 第 {idx} 科: {sub_id} {courseName} 無空位 ({quota_info})"
                    )
                    all_success = False
                    break  # 無空位，跳到下一科或結束

                # 找出所有可加選列（使用查詢餘額後的頁面）
                event_args = find_add_event_args(quota_html)
                last_msg = "無加選按鈕"
                if not event_args:
                    msg_txt = text_xpath(quota_html, X_MSG)
                    last_msg = msg_txt or last_msg
                    print(f'❌ 第 {idx} 科: {sub_id} {courseName} "{last_msg}"')
                    all_success = False
                    break

                for ea in event_args:
                    if stop_check_func and stop_check_func():
                        print("⚠️ 收到停止信號，停止選課")
                        return False, False

                    add_data = {
                        "ctl00_ToolkitScriptManager1_HiddenField": "",
                        "ctl00_MainContent_TabContainer1_ClientState": '{"ActiveTabIndex":1,"TabState":[true,true]}',
                        "__EVENTTARGET": "ctl00$MainContent$TabContainer1$tabSelected$gvToAdd",
                        "__EVENTARGUMENT": ea,
                        "__LASTFOCUS": "",
                        "__VIEWSTATE": vs,
                        "__VIEWSTATEGENERATOR": vg,
                        "__VIEWSTATEENCRYPTED": "",
                        "__EVENTVALIDATION": ev,
                        "ctl00$MainContent$TabContainer1$tabSelected$tbSubID": sub_id,
                        "ctl00$MainContent$TabContainer1$tabSelected$cpeWishList_ClientState": "false",
                    }
                    r = session.post(add_withdraw_url, data=add_data)

                    text_msg = text_xpath(r.text, X_MSG)
                    last_msg = text_msg or last_msg

                    if "系統偵測異常" in text_msg:
                        try:
                            if COOKIE_FILE.exists():
                                COOKIE_FILE.unlink()
                                print("🗑️ 已刪除 cookies 檔案 (系統偵測異常)")
                        except Exception as e:
                            print(f"刪除 cookies 失敗: {e}")
                        print(f'❌ 第 {idx} 科: {sub_id} {courseName} "{last_msg}"')
                        return False, True

                    if any(k in text_msg for k in ("成功", "已加選", "完成")):
                        success = True
                        print(f'✅ 第 {idx} 科: {sub_id} {courseName} "{last_msg}"')
                        break

                    # 更新隱藏欄位以便嘗試下一列
                    try:
                        vs, vg, ev = get_hidden_fields_fast(r.text)
                    except Exception:
                        break

                if not success:
                    print(
                        f"❌ 第 {idx} 科: {sub_id} {courseName} 加選失敗，重新查詢..."
                    )

            except RuntimeError as e:
                print(f"⚠️ 查詢餘額失敗：{e}")
                all_success = False
                break

        if not success:
            all_success = False

    return all_success, need_relogin


if __name__ == "__main__":
    main()
