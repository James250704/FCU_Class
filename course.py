import os
import re
import json
import pickle
import time
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import requests
from bs4 import BeautifulSoup
import ddddocr  # uv add ddddocr
from configparser import ConfigParser


BASE = "https://course.fcu.edu.tw"
COOKIE_FILE = Path("cookies.pkl")
SESSION_META = Path("session.json")


def save_response_to_file(filename, content):
    with open(filename, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"✅ 網頁內容已儲存到 {filename}")


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
    s = requests.Session()
    s.headers.update(headers)
    return s


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
    url = f"{base}/AddWithdraw.aspx?guid={guid}&lang={lang}"
    resp = session.get(url, allow_redirects=True)
    text = resp.text
    if (
        resp.url.endswith("Login.aspx")
        or is_login_page(text)
        or is_session_timeout(text)
    ):
        return False
    soup = BeautifulSoup(text, "html.parser")
    vs = soup.find("input", {"name": "__VIEWSTATE"})
    btn = soup.find(
        "input", {"name": "ctl00$MainContent$TabContainer1$tabSelected$btnGetSub"}
    )
    return bool(vs and btn)


def do_login(session: requests.Session, nid: str, pwd: str):
    session.get(f"{BASE}/")
    cap = session.get(f"{BASE}/validateCode.aspx")
    with open("captcha.jpg", "wb") as f:
        f.write(cap.content)
    ocr = ddddocr.DdddOcr()
    with open("captcha.jpg", "rb") as f:
        captcha = ocr.classification(f.read())
    print("自動識別驗證碼:", captcha)

    r = session.get(f"{BASE}/Login.aspx")
    soup = BeautifulSoup(r.text, "html.parser")
    viewstate = soup.find("input", {"name": "__VIEWSTATE"}).get("value", "")
    viewstategenerator = soup.find("input", {"name": "__VIEWSTATEGENERATOR"}).get(
        "value", ""
    )
    eventvalidation = soup.find("input", {"name": "__EVENTVALIDATION"}).get("value", "")

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


def get_hidden_fields(html: str, dump_name: str = "last_page.html"):
    soup = BeautifulSoup(html, "html.parser")

    def val(name):
        el = soup.find("input", {"name": name})
        return el.get("value", "") if el else ""

    vs = val("__VIEWSTATE")
    vg = val("__VIEWSTATEGENERATOR")
    ev = val("__EVENTVALIDATION")
    if not (vs and vg and ev):
        Path(dump_name).write_text(html, encoding="utf-8")
        raise RuntimeError("頁面缺少必要隱藏欄位，已落檔到 " + dump_name)
    return vs, vg, ev


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

            all_success = process_course_selection(
                session, add_withdraw_url, TB_SUB_IDS, stop_check_func
            )

            if all_success:
                print(f"🎉 所有課程選課成功！")
                break

            # 檢查是否達到重試次數限制（0 表示無限重試）
            if RETRY_COUNT > 0 and retry_round >= RETRY_COUNT:
                print("❌ 重試次數已達上限，選課結束")
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
        # 不啟用重試，執行單次選課
        process_course_selection(session, add_withdraw_url, TB_SUB_IDS, stop_check_func)

    print("\n===== 選課結束 =====")


def process_course_selection(
    session, add_withdraw_url, TB_SUB_IDS, stop_check_func=None
):
    """處理課程選課，返回是否全部成功"""
    all_success = True

    # 逐科處理
    for idx, sub_id in enumerate(TB_SUB_IDS, start=1):
        if stop_check_func and stop_check_func():
            print("⚠️ 收到停止信號，停止選課")
            return False

        print(f"\n===== 第 {idx} 科：{sub_id} =====")

        # 進入頁面拿初始隱藏欄位
        r = session.get(add_withdraw_url, allow_redirects=True)
        if is_session_timeout(r.text) or is_login_page(r.text):
            print("⚠️ 會話失效，需要重新登入")
            all_success = False
            continue

        vs, vg, ev = get_hidden_fields(r.text, dump_name=f"aw_{sub_id}_page.html")

        # 查詢該科
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
        vs, vg, ev = get_hidden_fields(r.text, dump_name=f"aw_{sub_id}_query.html")

        # 找出所有可加選列的 __EVENTARGUMENT
        event_args = find_add_event_args(r.text)
        if not event_args:
            print("找不到可加選按鈕，可能查無課或未開放。")
            # 顯示頁面訊息
            soup = BeautifulSoup(r.text, "html.parser")
            msg = soup.find(
                "span",
                {"id": "ctl00_MainContent_TabContainer1_tabSelected_lblMsgBlock"},
            )
            if msg:
                print("訊息：", msg.get_text(strip=True))
            all_success = False
            continue

        success = False
        for ea in event_args:
            if stop_check_func and stop_check_func():
                print("⚠️ 收到停止信號，停止選課")
                return False

            add_data = {
                "ctl00_ToolkitScriptManager1_HiddenField": "",
                "ctl00_MainContent_TabContainer1_ClientState": '{"ActiveTabIndex":1,"TabState":[true,true]}',
                "__EVENTTARGET": "ctl00$MainContent$TabContainer1$tabSelected$gvToAdd",
                "__EVENTARGUMENT": ea,  # 例如 addCourse$0
                "__LASTFOCUS": "",
                "__VIEWSTATE": vs,
                "__VIEWSTATEGENERATOR": vg,
                "__VIEWSTATEENCRYPTED": "",
                "__EVENTVALIDATION": ev,
                "ctl00$MainContent$TabContainer1$tabSelected$tbSubID": sub_id,
                "ctl00$MainContent$TabContainer1$tabSelected$cpeWishList_ClientState": "false",
            }
            r = session.post(add_withdraw_url, data=add_data)

            soup = BeautifulSoup(r.text, "html.parser")
            msg = soup.find(
                "span",
                {"id": "ctl00_MainContent_TabContainer1_tabSelected_lblMsgBlock"},
            )
            text = msg.get_text(strip=True) if msg else "(無訊息)"
            print(f"訊息：{text}")

            # 成功關鍵詞自行調整
            if any(k in text for k in ("成功", "已加選", "完成")):
                success = True
                break

            # 更新隱藏欄位以便嘗試下一列
            try:
                vs, vg, ev = get_hidden_fields(
                    r.text, dump_name=f"aw_{sub_id}_after_{ea}.html"
                )
            except Exception:
                # 若頁面跳離或缺欄位就中止此科
                break

        if not success:
            print(f"→ 科目 {sub_id} 未成功加選。")
            all_success = False

    return all_success


if __name__ == "__main__":
    main()
