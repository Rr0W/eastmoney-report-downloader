#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东方财富网 财报 + 研报 批量下载器（便携版 v2.0）
================================================================
一条命令完成：抓列表 → 下载 PDF → 按规范重命名 → 生成清单 → 校验 → 打包 zip

【可迁移性】
  · 单个文件，无第三方依赖（纯 Python 标准库），拷贝到任意机器即可运行
  · 跨平台：Windows / macOS / Linux
  · 不依赖任何绝对路径、不读固定位置的配置、不写系统目录
  · 断点续传：中断后重跑同一条命令即可，不会重复下载
  · 幂等：重复执行不会产生 _2 副本

【运行环境】
  Python 3.7+（3.9+ 更佳）。无需 pip install，**完整功能开箱可用**。
  可选增强（不装也能正常工作，只是多一层校验冗余）：
      pip install pypdf            # 或系统里装了 pdftotext (poppler-utils)
  研报日期识别顺序：PDF 元数据标题 → 封面正文（可选组件）→ PDF 生成日期 → 东财发布日。
  前两级之外的组合已能覆盖全部主流券商 PDF；可选组件装了则多一层兜底。

【用法】
  python eastmoney_downloader.py --code 600519 --period 2026 --zip
  python eastmoney_downloader.py --name 茅台 --period 2026H1 --types report
  python eastmoney_downloader.py --code 000568 --period 2026-01-01~2026-06-30
  python eastmoney_downloader.py --code 600519 --period 2026 --list-only
  python eastmoney_downloader.py --doctor          # 环境与连通性自检

【参数速查】
  --code    6 位股票代码（600519）
  --name    公司名（茅台），自动解析成代码；与 --code 二选一
  --period  2026 | 2026H1 | 2026H2 | 2026Q1~Q4 | 2026-03 | 2026-01-01~2026-06-30
            缺省为当年
  --types   report=研报 / ann=财报公告 / report,ann（默认两者都下）
  --zip     完成后打包为 UTF-8 文件名 zip

【命名规范（默认值，无需手动指定）】
  研报：<YYMMDD>_<行业>_<公司简称+代码>_<机构>_<研究员>.pdf
        260918_白酒_贵州茅台600519_诚通证券_陈文倩.pdf
  财报：<YYMMDD>_<行业>_<公司简称+代码>_<报告名>.pdf
        260417_白酒_贵州茅台600519_2025年年度报告.pdf
  日期口径：研报取「研报自身标注日期」，财报取「披露日期（公告日）」。

【数据来源】
  东方财富网公开接口，仅供研究参考，不构成投资建议。
"""
from __future__ import annotations

import argparse
import calendar
import csv
import datetime
import html
import json
import os
import platform
import random
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

__version__ = "2.3"

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
REF = "https://data.eastmoney.com/"

# 东方财富「名称搜索」接口的公开 token（非密钥，网页前端也带同一个）。
# 支持用环境变量覆盖，避免在代码里写死：`EM_SEARCH_TOKEN=xxx python ...`
SEARCH_TOKEN = os.environ.get("EM_SEARCH_TOKEN", "D43BF722C8E33BDC906FB84D85E326E8")

# 下载判定的最小有效 PDF 大小（字节）。小于此值视为无效响应（错误页/反爬页）。
MIN_PDF_SIZE = 2000
# 断点续传缓存文件名（位于输出目录内）
CACHE_FILENAME = ".downloaded.json"
# 单次下载重试次数 / 默认超时（秒）
DOWNLOAD_TRIES = 5
DOWNLOAD_TIMEOUT = 60

# 完整的浏览器请求头族。pdf.dfcfw.com 的 CDN 偶尔会对「三无」请求（只有 UA+Referer）
# 返回混淆 JS 反爬页（形如 <script>function a(a){... _0x649a ...），补齐浏览器指纹
# 头可显著降低触发概率。
BROWSER_HEADERS = {
    "User-Agent": UA,
    "Accept": ("text/html,application/xhtml+xml,application/xml;q=0.9,"
               "image/avif,image/webp,image/apng,*/*;q=0.8"),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "sec-ch-ua": ('"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"'),
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
}

# 反爬 challenge 页的特征片段（瑞数/东财自研 WAF 的混淆 JS）
_CHALLENGE_MARKS = (b"function a(a)", b"_0x", b"__jsl", b"jsl_clearance",
                    b"<script", b"challenge")

# 控制台标记：Windows 的 GBK 代码页打不出 ✓/✗，自动降级为 ASCII
def _pick_marks():
    try:
        "✓".encode(sys.stdout.encoding or "utf-8")
        return "✓", "✗"
    except Exception:
        return "OK", "XX"


MARK_OK, MARK_BAD = _pick_marks()

# ---------------- 规则表 ----------------
# 定期报告识别：排除业绩说明会 / 内控 / ESG 等「描述报告的公告」
RE_PERIODIC = re.compile(
    r"(20\d{2}\s*年\s*(年度报告|半年度报告|第[一三]季度报告|季度报告))|(20\d{2}\s*年\s*(年报|中报))")
EXCLUDE_WORDS = ["说明会", "业绩说明", "内部控制", "社会责任", "审计报告", "专项说明",
                 "ESG", "环境、社会", "述职报告", "履职情况", "风险评估报告",
                 "延期披露", "预约披露", "披露时间", "更正公告", "更正后"]

# 研报自身日期：形如 机构-公司(代码)标题-260918.pdf
RE_TITLE_DATE = re.compile(r"-(\d{6})(?:\.pdf)?")
RE_TXT_DATE = re.compile(r"(20\d{2})\s*[年\-/\.]\s*(\d{1,2})\s*[月\-/\.]\s*(\d{1,2})")

SPEC_REPORT = "{date}_{ind}_{company}{code}_{org}_{author}"
SPEC_ANN = "{date}_{ind}_{company}{code}_{title}"


def log(*a):
    print(*a, flush=True)


def warn(*a):
    sys.stderr.write(" ".join(str(x) for x in a) + "\n")
    sys.stderr.flush()


def safe_name(s, limit=80):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(s))
    return re.sub(r"\s+", " ", s).strip().strip(".")[:limit]


def norm_ind(s):
    """白酒Ⅱ -> 白酒"""
    return re.sub(r"[ⅠⅡⅢⅣⅤ]+$", "", str(s or "")).strip()


def ymd_to_yymmdd(s):
    """2026-09-21 -> 260921"""
    return str(s).replace("-", "").replace("/", "")[2:8] if s else ""


# ==================== 时间范围解析 ====================
def _ymd(s):
    """严格解析日期：2026-09-21 / 2026.9.21 / 20260921"""
    t = str(s or "").strip()
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", t)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    elif re.match(r"^\d{8}$", t):
        y, mo, d = int(t[:4]), int(t[4:6]), int(t[6:8])
    else:
        return ""
    if not (1900 <= y <= 2100 and 1 <= mo <= 12 and 1 <= d <= 31):
        return ""
    try:
        datetime.date(y, mo, d)
    except Exception:
        return ""
    return "%04d-%02d-%02d" % (y, mo, d)


def parse_period(s):
    """时间描述 -> (begin, end, label)。支持全年/半年/季度/整月/单日/区间"""
    t = re.sub(r"\s+", "", str(s or ""))
    if not t:
        return "", "", ""

    for sep in ("~", "到", "至", "—", "－", ".."):
        if sep in t:
            a, _, b = t.partition(sep)
            s1, s2 = _ymd(a), _ymd(b.lstrip("-"))
            if s1 and s2:
                return s1, s2, "%s_%s" % (s1.replace("-", ""), s2.replace("-", ""))
    m = re.match(r"^(\d{8})-(\d{8})$", t)
    if m and _ymd(m.group(1)) and _ymd(m.group(2)):
        return _ymd(m.group(1)), _ymd(m.group(2)), "%s_%s" % (m.group(1), m.group(2))

    m = re.match(r"^(\d{4})[Hh]([12])$", t)
    if m:
        y, h = m.group(1), int(m.group(2))
        return (("%s-01-01" % y, "%s-06-30" % y, "%sH1" % y) if h == 1
                else ("%s-07-01" % y, "%s-12-31" % y, "%sH2" % y))
    m = re.match(r"^(\d{4})[Qq]([1-4])$", t)
    if m:
        y, q = m.group(1), int(m.group(2))
        b = ["01-01", "04-01", "07-01", "10-01"][q - 1]
        e = ["03-31", "06-30", "09-30", "12-31"][q - 1]
        return "%s-%s" % (y, b), "%s-%s" % (y, e), "%sQ%d" % (y, q)

    m = re.match(r"^(\d{4})[-/.](\d{1,2})$", t)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            last = calendar.monthrange(y, mo)[1]
            return ("%04d-%02d-01" % (y, mo), "%04d-%02d-%02d" % (y, mo, last),
                    "%04d%02d" % (y, mo))

    d = _ymd(t)
    if d:
        return d, d, d.replace("-", "")

    m = re.match(r"^(\d{4})$", t)
    if m:
        return "%s-01-01" % m.group(1), "%s-12-31" % m.group(1), m.group(1)

    return "", "", ""


# ==================== 异常 ====================
class DownloadError(Exception):
    """下载失败。`retryable` 表示是否值得重试（网络抖动可重试，内容无效不必空转）。"""

    def __init__(self, message, retryable=True):
        super().__init__(message)
        self.retryable = retryable


# ==================== HTTP ====================
def _headers(referer=REF, accept=None, compress=True):
    h = dict(BROWSER_HEADERS)
    h["Referer"] = referer
    if accept:
        h["Accept"] = accept
    if not compress:
        # urllib 不会自动解压 gzip；去掉 Accept-Encoding 让服务端返回明文
        h.pop("Accept-Encoding", None)
    return h


def _is_challenge(data):
    """判断响应是否为反爬 JS challenge 页（而非真实内容）。"""
    if not data:
        return False
    if isinstance(data, str):
        data = data.encode("utf-8", "ignore")
    head = data[:2048]
    if head[:4] == b"%PDF":
        return False
    return any(m in head for m in _CHALLENGE_MARKS)


def _http(url, referer=REF, tries=3, timeout=30, raw=False, accept=None):
    """统一的 HTTP 入口。urllib 失败时若存在 curl 则回退（个别接口会拒绝 Python 的 TLS 指纹）"""
    last = ""
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=_headers(referer, accept, compress=False))
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
            if raw and _is_challenge(data):
                raise ValueError("反爬 challenge 页 len=%d" % len(data))
            return data if raw else data.decode("utf-8", "ignore")
        except Exception as e:
            last = str(e)
        if i == 0:
            alt = _via_curl(url, referer, timeout, raw, accept)
            if alt and not (raw and _is_challenge(alt)):
                return alt
        if i < tries - 1:
            warn("   [retry %d/%d] %s" % (i + 1, tries, last))
            time.sleep(1.5 * (i + 1))
    warn("   请求失败：%s -> %s" % (url[:80], last))
    return b"" if raw else ""


def _via_curl(url, referer=REF, timeout=15, raw=False, accept=None):
    """curl 兜底通道（Windows 10+ 内置 curl.exe）。带完整浏览器头 + gzip 解压。"""
    if not shutil.which("curl"):
        return b"" if raw else ""
    try:
        h = _headers(referer, accept)
        cmd = ["curl", "-4", "-sS", "--compressed", "-m", str(timeout)]
        for k, v in h.items():
            cmd += ["-H", "%s: %s" % (k, v)]
        cmd.append(url)
        r = subprocess.run(cmd, capture_output=True, timeout=timeout + 15)
        return r.stdout if raw else r.stdout.decode("utf-8", "ignore")
    except Exception:
        return b"" if raw else ""


def http_json(url, referer=REF, **kw):
    txt = _http(url, referer=referer, **kw)
    if not txt:
        return None
    try:
        return json.loads(txt)
    except Exception:
        warn("   JSON 解析失败: %s" % str(txt)[:150])
        return None


# ==================== 标的解析 ====================
def resolve_by_name(name, market=None):
    """公司名 -> (代码, 简称)。取 A 股首个匹配；失败返回 ('','')"""
    url = ("https://searchapi.eastmoney.com/api/suggest/get?input=%s&type=14"
           "&token=%s&count=10"
           % (urllib.parse.quote(name), SEARCH_TOKEN))
    js = http_json(url, referer="https://www.eastmoney.com/")
    rows = ((js or {}).get("QuotationCodeTable") or {}).get("Data") or []
    for r in rows:
        if r.get("Classify") != "AStock":
            continue
        code = str(r.get("Code") or "")
        if not re.match(r"^\d{6}$", code):
            continue
        if market == "SH" and r.get("MktNum") != "1":
            continue
        if market == "SZ" and r.get("MktNum") != "0":
            continue
        return code, re.sub(r"\s+", "", r.get("Name") or "")
    return "", ""


def secid(code):
    """代码 -> (secid 前缀, 交易所后缀)"""
    if code[0] == "6":
        return "1", "SH"
    if code[0] in "48":
        return "0", "BJ"
    return "0", "SZ"


# ==================== 公司简称 / 申万行业 ====================
_QUOTE_CACHE = {}
_QUOTE_HOSTS = ["push2.eastmoney.com", "push2delay.eastmoney.com", "push2his.eastmoney.com"]


def _pick_industry(rows):
    """优先二级(Ⅱ)，其次三级(Ⅲ)，最后任意行业板块"""
    cand = [r for r in rows if r.get("BOARD_TYPE") == "行业" and r.get("BOARD_NAME")]
    for suf in ("Ⅱ", "Ⅲ", "Ⅰ"):
        for r in cand:
            if r["BOARD_NAME"].endswith(suf):
                return norm_ind(r["BOARD_NAME"])
    return norm_ind(cand[0]["BOARD_NAME"]) if cand else ""


def quote_f10(code):
    """首选：F10 板块归属接口，一次拿到 简称 + 申万行业（实测最稳）"""
    _, mkt = secid(code)
    url = ("https://datacenter-web.eastmoney.com/api/data/v1/get"
           "?reportName=RPT_F10_CORETHEME_BOARDTYPE&columns=ALL"
           "&filter=(SECUCODE%%3D%%22%s.%s%%22)" % (code, mkt))
    js = http_json(url)
    rows = ((js or {}).get("result") or {}).get("data") or []
    if not rows:
        return "", ""
    return (re.sub(r"\s+", "", rows[0].get("SECURITY_NAME_ABBR") or ""),
            _pick_industry(rows))


def quote_push2(code):
    """兜底：行情接口 f58=简称, f127=行业"""
    pre, _ = secid(code)
    qs = "secid=%s.%s&fields=f57,f58,f127&invt=2&fltt=1" % (pre, code)
    for host in _QUOTE_HOSTS:
        out = _http("https://%s/api/qt/stock/get?%s" % (host, qs),
                    referer="https://quote.eastmoney.com/", tries=1, timeout=12)
        try:
            d = (json.loads(out).get("data") or {}) if out.strip() else {}
            name = re.sub(r"\s+", "", d.get("f58") or "")      # "五 粮 液" -> "五粮液"
            ind = norm_ind(d.get("f127"))
            if name and ind:
                return name, ind
        except Exception:
            pass
    return "", ""


def quote_info(code, tries=3):
    """取 (公司简称, 申万行业)，多来源回退，全失败返回 ('', '')"""
    if code in _QUOTE_CACHE:
        return _QUOTE_CACHE[code]
    for i in range(tries):
        try:
            name, ind = quote_f10(code)
            if name and ind:
                _QUOTE_CACHE[code] = (name, ind)
                return name, ind
        except Exception as e:
            warn("   F10 板块接口失败: %s" % e)
        name, ind = quote_push2(code)
        if name and ind:
            _QUOTE_CACHE[code] = (name, ind)
            return name, ind
        if i < tries - 1:
            warn("   公司简称/行业第 %d/%d 轮未取到，重试" % (i + 1, tries))
            time.sleep(1.2 * (i + 1))
    _QUOTE_CACHE[code] = ("", "")
    return "", ""


# ==================== PDF 自身日期（三级回退，零依赖优先） ====================
def _decode_pdf_string(tok):
    """解析 PDF 字符串对象 (…) 或 <hex>，返回解码后的文本"""
    if tok.startswith(b"<"):
        h = re.sub(rb"\s+", b"", tok[1:-1])
        try:
            b = bytes.fromhex(h.decode("ascii"))
        except Exception:
            return ""
    else:
        b, out, i = tok[1:-1], bytearray(), 0
        esc = {0x6E: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12, 0x28: 40, 0x29: 41, 0x5C: 92}
        while i < len(b):
            c = b[i]
            if c == 0x5C and i + 1 < len(b):
                nx = b[i + 1]
                if nx in esc:
                    out.append(esc[nx]); i += 2; continue
                if 0x30 <= nx <= 0x37:
                    j, o = i + 1, b""
                    while j < len(b) and 0x30 <= b[j] <= 0x37 and len(o) < 3:
                        o += bytes([b[j]]); j += 1
                    out.append(int(o, 8) & 0xFF); i = j; continue
                i += 2; continue
            out.append(c); i += 1
        b = bytes(out)
    if b[:2] == b"\xfe\xff":
        return b[2:].decode("utf-16-be", "ignore")
    if b[:2] == b"\xff\xfe":
        return b[2:].decode("utf-16-le", "ignore")
    for enc in ("utf-8", "gbk", "big5", "latin-1"):
        try:
            return b.decode(enc)
        except Exception:
            continue
    return ""


RE_PDF_TITLE = re.compile(rb"/Title\s*(\((?:\\.|[^\\()])*\)|<[0-9A-Fa-f\s]+>)", re.S)
RE_PDF_CREATE = re.compile(rb"/(?:CreationDate|ModDate)\s*\(?\s*D:\s*(\d{4})(\d{2})(\d{2})")
RE_XMP_TITLE = re.compile(rb"<dc:title>.*?<rdf:li[^>]*>(.*?)</rdf:li>", re.S)
RE_XMP_CREATE = re.compile(rb"(?:xmp:CreateDate|pdf:CreateDate|xmp:ModifyDate)>(\d{4})-(\d{2})-(\d{2})")
RE_STREAM = re.compile(rb"stream\r?\n")


def _valid_yymmdd(s):
    """260918 -> True；同时滤掉 050112 这类脏元数据"""
    if not re.match(r"^\d{6}$", s):
        return False
    y, mo, d = 2000 + int(s[:2]), int(s[2:4]), int(s[4:6])
    try:
        datetime.date(y, mo, d)
    except Exception:
        return False
    return y >= 2010


def _yymmdd_of(y, mo, d):
    s = "%02d%02d%02d" % (int(y) % 100, int(mo), int(d))
    return s if _valid_yymmdd(s) else ""


def _inflate_all(raw, max_stream=200_000, max_total=32 * 1024 * 1024, max_n=60):
    """兜底：把 PDF 里较小的 FlateDecode 流解压出来（有些生成器把 Info 字典塞进对象流）。
    严格限量，避免大文件耗时。"""
    import zlib
    out, total, n = [], 0, 0
    for m in RE_STREAM.finditer(raw):
        if n >= max_n or total >= max_total:
            break
        start = m.end()
        end = raw.find(b"endstream", start)
        if end < 0:
            continue
        blob = raw[start:end]
        if len(blob) > max_stream:
            continue
        for wbits in (15, -15):
            try:
                dec = zlib.decompress(blob, wbits)
                out.append(dec)
                total += len(dec)
                n += 1
                break
            except Exception:
                continue
    return out


def pdf_meta_date(path):
    """第 1 级（纯标准库）：PDF 元数据 /Title 或 XMP dc:title 里的研报自身日期。
    券商标注格式 `机构-公司(代码)标题-260918.pdf` —— 复刻 macOS mdls 的能力但不依赖系统。"""
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except Exception:
        return ""
    cands = [_decode_pdf_string(m.group(1)) for m in RE_PDF_TITLE.finditer(raw)]
    cands += [m.group(1).decode("utf-8", "ignore") for m in RE_XMP_TITLE.finditer(raw)]
    for txt in cands:
        if not txt:
            continue
        m = RE_TITLE_DATE.search(txt)
        if m and _valid_yymmdd(m.group(1)):
            return m.group(1)
    return ""


def pdf_file_date(path, publish="", window=12):
    """第 2 级（纯标准库）：PDF 自身的 /CreationDate /ModDate。
    很多券商的 PDF 是「报告当天生成」，此日期即研报日期，可补上无 /Title 的那部分。
    为防误用（该字段偶尔是文件生成时间而非报告日期），要求落在
    [publishDate - window, publishDate] 区间内，否则弃用。"""
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except Exception:
        return ""
    cands = []
    for m in RE_PDF_CREATE.finditer(raw):
        cands.append("%s%s%s" % (m.group(1)[-2:].decode(), m.group(2).decode(), m.group(3).decode()))
    for m in RE_XMP_CREATE.finditer(raw):
        cands.append("%s%s%s" % (m.group(1)[-2:].decode(), m.group(2).decode(), m.group(3).decode()))
    if not cands:
        for blob in _inflate_all(raw):
            for m in RE_PDF_CREATE.finditer(blob):
                cands.append("%s%s%s" % (m.group(1)[-2:].decode(), m.group(2).decode(), m.group(3).decode()))
            if cands:
                break
    cands = [c for c in cands if _valid_yymmdd(c)]
    if not cands:
        return ""
    if not publish:
        return cands[0]

    def d(s):
        return datetime.date(2000 + int(s[:2]), int(s[2:4]), int(s[4:6]))

    lo = d(publish) - datetime.timedelta(days=window)
    hi = d(publish)
    ok = [c for c in cands if lo <= d(c) <= hi]
    if not ok:
        return ""
    return max(ok)          # 取最接近发布日的那个


def pdf_text_date(path):
    """第 3 级：pypdf（若已安装）或 pdftotext（若在 PATH），从首页正文取日期。
    这两者需要可选组件，但最忠实于「研报封面标注日期」。"""
    txt = ""
    try:
        from pypdf import PdfReader            # type: ignore
        r = PdfReader(path)
        txt = "\n".join((r.pages[i].extract_text() or "")
                        for i in range(min(2, len(r.pages))))
    except Exception:
        txt = ""
    if not txt and shutil.which("pdftotext"):
        try:
            r = subprocess.run(["pdftotext", "-f", "1", "-l", "2", path, "-"],
                               capture_output=True, timeout=30)
            txt = r.stdout.decode("utf-8", "ignore")
        except Exception:
            txt = ""
    for y, mo, d in RE_TXT_DATE.findall(txt or ""):
        got = _yymmdd_of(y, mo, d)
        if got:
            return got
    return ""


def pdf_own_date(path, publish=""):
    """研报自身标注日期 YYMMDD，返回 (日期, 来源)。
    来源: title=PDF元数据标题 / text=封面正文(需可选组件) / pdfdate=PDF生成日期 / ''=未取到
    零依赖时前两级标题+生成日期已能覆盖绝大多数券商的 PDF。"""
    d = pdf_meta_date(path)
    if d:
        return d, "title"
    d = pdf_text_date(path)
    if d:
        return d, "text"
    d = pdf_file_date(path, publish)
    if d:
        return d, "pdfdate"
    return "", ""


# ==================== 命名模板 ====================
def fill(spec, fields):
    """套用模板 -> (文件名, 缺失字段列表)"""
    miss = [k for k, v in fields.items() if ("{%s}" % k) in spec and not str(v).strip()]
    out = spec
    for k, v in fields.items():
        out = out.replace("{%s}" % k, safe_name(v, 60))
    miss += re.findall(r"\{(\w+)\}", out)
    return safe_name(out, 180), miss


# ==================== 列表抓取 ====================
def fetch_reports(code, begin, end):
    """券商研报列表（infoCode 用于拼 H3_ PDF 直链）"""
    out, page = [], 1
    while page <= 50:
        url = ("https://reportapi.eastmoney.com/report/list?cb=&pageSize=100&pageNo=%d"
               "&qType=0&industryCode=*&beginTime=%s&endTime=%s&code=%s"
               "&rcode=&fields=&p=%d&pageNum=%d&pageNumber=%d"
               % (page, begin, end, code, page, page, page))
        js = http_json(url, referer="https://data.eastmoney.com/report/")
        if not js:
            break
        data = js.get("data") or []
        out.extend(data)
        hits = js.get("hits") or 0
        log("  研报 第%d页 +%d（累计 %d/%d）" % (page, len(data), len(out), hits))
        if not data or len(out) >= hits:
            break
        page += 1
        time.sleep(0.8)
    # 同一 infoCode 去重（接口偶发重复返回）
    seen, uniq = set(), []
    for r in out:
        ic = r.get("infoCode")
        if ic and ic not in seen:
            seen.add(ic); uniq.append(r)
    return uniq


def fetch_anns(code, begin, end, max_pages=40):
    """公告列表（按日期倒序返回，翻过下界即停）"""
    out, page = [], 1
    while page <= max_pages:
        url = ("https://np-anotice-stock.eastmoney.com/api/security/ann?page_size=100"
               "&page_index=%d&ann_type=A&client_source=web&stock_list=%s&f_node=0&s_node=0"
               % (page, code))
        js = http_json(url, referer="https://data.eastmoney.com/notices/stock/%s.html" % code)
        if not js:
            break
        lst = (js.get("data") or {}).get("list") or []
        if not lst:
            break
        out.extend(lst)
        dates = [a.get("notice_date", "")[:10] for a in lst]
        log("  公告 第%d页 +%d（累计 %d，最早 %s）" % (page, len(lst), len(out), min(dates)))
        if min(dates) < begin:
            break
        page += 1
        time.sleep(0.8)
    return [a for a in out if begin <= a.get("notice_date", "")[:10] <= end]


def ann_match(a, mode, keyword):
    """periodic 模式：排除一切「描述报告的公告」，只留报告本身"""
    title = a.get("title", "")
    t = title.split(":", 1)[-1]
    if mode == "all":
        return True
    if mode == "keyword":
        return keyword in title
    if any(w in t for w in EXCLUDE_WORDS):
        return False
    if t.startswith("关于") or "的公告" in t:
        return False
    return bool(RE_PERIODIC.search(t))


# ==================== 公告 PDF 地址 ====================
_ATTACH_CACHE = {}
_ATTACH_MULTI = set()


def ann_pdf_info(art_code):
    """查公告详情，返回 (首选地址, 附件个数)。失败返回 ('', 0)"""
    if art_code in _ATTACH_CACHE:
        return _ATTACH_CACHE[art_code]
    url = ("https://np-cnotice-stock.eastmoney.com/api/content/ann?art_code=%s"
           "&client_source=web&page_index=1" % art_code)
    js = http_json(url)
    u, n = "", 0
    if js:
        data = js.get("data") or {}
        lst = data.get("attach_list") or []
        n = len(lst)
        u = data.get("attach_url_web") or data.get("attach_url") or ""
        if not u and lst:
            u = lst[0].get("attach_url") or ""
    if n > 1:
        _ATTACH_MULTI.add(art_code)
    _ATTACH_CACHE[art_code] = (u, n)
    return u, n


def ann_pdf_urls(art_code):
    """公告 PDF 的候选地址列表。

    裸链 https://pdf.dfcfw.com/pdf/H2_<art_code>_1.pdf 会被 CDN 缓存，
    极少数情况下会缓存住一个几百字节的错误页 —— 换个查询参数就绕开了。
    最后再挂上详情接口给出的权威地址作为兜底。
    """
    base = "https://pdf.dfcfw.com/pdf/H2_%s_1.pdf" % art_code
    urls = [base, "%s?%d.pdf" % (base, int(time.time() * 1000))]

    # 详情接口有额外附件（正文+附件分开）时提示，避免漏下
    u, n = ann_pdf_info(art_code)
    if n > 1:
        warn("   提示：该公告有 %d 个附件，当前只取第 1 个（%s）" % (n, art_code))
    if u and u not in urls:
        urls.append(u)
    return urls


# ==================== 下载 ====================
def download(url, path, referer=REF, tries=DOWNLOAD_TRIES, timeout=DOWNLOAD_TIMEOUT,
             urls=None):
    """下载 PDF 到 `path`，成功返回 (True, 字节数)，失败返回 (False, 0)。

    `urls` 可给出多个候选地址，每轮换一个重试，用于绕开 CDN 偶发的错误页缓存。
    遇到反爬 challenge 页时自动切换到 curl 通道并拉长退避（等待抖动恢复），
    而不是在同一批等价地址上空转。
    """
    if os.path.exists(path) and os.path.getsize(path) > MIN_PDF_SIZE:
        return True, os.path.getsize(path)

    candidates = [u for u in (urls or [url]) if u] or [url]
    last_error = ""
    for attempt in range(tries):
        u = candidates[attempt % len(candidates)]
        try:
            # 首选 urllib（带完整浏览器头）；反爬时 _http 内部会先尝试 curl 兜底
            data = _http(u, referer=referer, tries=2, timeout=timeout, raw=True,
                         accept="application/pdf,*/*;q=0.8")
            if not isinstance(data, (bytes, bytearray)) or len(data) < MIN_PDF_SIZE:
                raise DownloadError("非 PDF 内容 len=%s" % (len(data) if data else 0),
                                    retryable=False)
            if data[:4] != b"%PDF":
                raise DownloadError("非 PDF 魔数（len=%d）" % len(data), retryable=False)
            tmp = path + ".part"
            with open(tmp, "wb") as f:
                f.write(data)
            os.replace(tmp, path)
            return True, len(data)
        except DownloadError as e:
            last_error = str(e)
            # 内容无效（错误页/反爬页），不值得在同一批等价地址上空转太久
            challenged = "challenge" in last_error or "反爬" in last_error
            warn("   [retry %d/%d] %s" % (attempt + 1, tries, last_error))
            if attempt < tries - 1:
                time.sleep((6 if challenged else 2) + 2 * attempt + random.random() * 1.5)
        except Exception as e:
            last_error = str(e)
            warn("   [retry %d/%d] %s" % (attempt + 1, tries, last_error))
            if attempt < tries - 1:
                time.sleep(2 + 2 * attempt + random.random() * 1.5)
    return False, 0


def unique_path(d, name):
    dst = os.path.join(d, name)
    if not os.path.exists(dst):
        return dst, name
    base, ext = os.path.splitext(name)
    n = 2
    while os.path.exists(os.path.join(d, "%s_%d%s" % (base, n, ext))):
        n += 1
    return os.path.join(d, "%s_%d%s" % (base, n, ext)), "%s_%d%s" % (base, n, ext)


def save_cache(path, cache):
    """增量落盘：中断后重跑能正确续传"""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)
    except Exception as e:
        warn("   缓存写入失败: %s" % e)


def verify_pdfs(outdir):
    """校验 PDF 魔数 + 结束标记（能识别截断）。返回 (总数, 损坏数, 坏文件)"""
    tot, bad, names = 0, 0, []
    for root, _, files in os.walk(outdir):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            tot += 1
            p = os.path.join(root, f)
            try:
                with open(p, "rb") as fh:
                    if fh.read(4) != b"%PDF":
                        raise ValueError("bad magic")
                    fh.seek(max(0, os.path.getsize(p) - 2048))
                    if b"%%EOF" not in fh.read():
                        raise ValueError("no EOF")
            except Exception:
                bad += 1
                names.append(f)
    return tot, bad, names


def make_zip(outdir, zip_basename=""):
    """打包 zip，强制 UTF-8 文件名标志位（避免 Windows 解压乱码）"""
    import zipfile
    parent = os.path.dirname(os.path.abspath(outdir.rstrip(os.sep + "/")))
    base = zip_basename or os.path.basename(outdir.rstrip(os.sep + "/"))
    zpath = os.path.join(parent, base + ".zip")
    if os.path.exists(zpath):
        os.remove(zpath)
    n = 0
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(outdir):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for f in sorted(files):
                if f == ".DS_Store" or f.endswith((".py", ".part")) \
                        or f == "manifest.json" or f.startswith(".dl_") \
                        or f.startswith(".downloaded"):
                    continue
                p = os.path.join(root, f)
                zi = zipfile.ZipInfo.from_file(p, os.path.relpath(p, parent))
                zi.flag_bits |= 0x800
                zi.compress_type = zipfile.ZIP_DEFLATED
                with z.open(zi, "w") as w, open(p, "rb") as fh:
                    w.write(fh.read())
                n += 1
    return zpath, n


# ==================== 清单输出 ====================
def write_index(outdir, manifest, code, begin, end):
    if not manifest:
        return
    with open(os.path.join(outdir, "清单.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["类别", "日期", "机构", "标题", "评级", "研究员",
                                          "文件", "大小KB", "来源链接"], extrasaction="ignore")
        w.writeheader()
        for m in manifest:
            w.writerow(m)

    L = ["# %s 财报与研报下载清单\n" % code,
         "> 数据来源：东方财富网　·　区间：%s ~ %s\n" % (begin, end)]
    for cat, tag in (("定期报告", "一"), ("公告", "二"), ("研报", "三")):
        items = [m for m in manifest if m["类别"] == cat]
        if not items:
            continue
        L.append("\n## %s、%s 共 %d 份\n" % (tag, cat, len(items)))
        L.append("| 日期 | 机构 | 标题 | 大小 |")
        L.append("|---|---|---|---|")
        for m in sorted(items, key=lambda x: x["日期"]):
            L.append("| %s | %s | [%s](%s) | %.2f MB |" % (
                m["日期"], m.get("机构", ""), m["标题"],
                urllib.parse.quote(m["文件"]), m["大小KB"] / 1024))
    with open(os.path.join(outdir, "清单.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(L))

    def rows(items, isrep):
        r = []
        for m in sorted(items, key=lambda x: x["日期"]):
            who = (m.get("研究员", "") or "").replace(",", "、")
            r.append("<tr><td class='d'>%s</td><td>%s</td><td><a href=\"%s\">%s</a></td>"
                     "<td class='s'>%.2f MB</td></tr>"
                     % (m["日期"], html.escape(who if isrep else ""),
                        urllib.parse.quote(m["文件"]), html.escape(m["标题"]), m["大小KB"] / 1024))
        return "\n".join(r)

    n_r = len([m for m in manifest if m["类别"] == "研报"])
    n_a = len([m for m in manifest if m["类别"] != "研报"])
    orgs = len(set(m.get("机构", "") for m in manifest if m.get("机构")))
    sz = sum(m["大小KB"] for m in manifest) / 1024
    H = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><title>%s 财报与研报清单</title>
<style>:root{--bg:#f7f8fa;--card:#fff;--ink:#1a1d24;--sub:#6b7280;--line:#e8eaef;--red:#c8102e}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;line-height:1.6}
.wrap{max-width:1080px;margin:0 auto;padding:40px 24px 64px}h1{font-size:26px;margin:0 0 6px}
.sub{color:var(--sub);font-size:13px;margin-bottom:26px}.kpis{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:28px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 22px;min-width:150px}
.kpi b{display:block;font-size:28px;color:var(--red);font-weight:700}.kpi span{font-size:12.5px;color:var(--sub)}
h2{font-size:17px;margin:34px 0 12px;padding-left:10px;border-left:4px solid var(--red)}
table{width:100%%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
th{background:#f1f3f7;font-size:12.5px;color:#4b5563;text-align:left;padding:10px 14px;font-weight:600}
td{padding:9px 14px;border-top:1px solid var(--line);font-size:13.5px;vertical-align:top}
td.d{color:var(--sub);white-space:nowrap;font-variant-numeric:tabular-nums}
td.s{color:var(--sub);white-space:nowrap;text-align:right;font-variant-numeric:tabular-nums}
a{color:#0b5fff;text-decoration:none}a:hover{text-decoration:underline}
.foot{margin-top:30px;font-size:12px;color:var(--sub);border-top:1px solid var(--line);padding-top:14px}</style></head>
<body><div class="wrap"><h1>%s 财报与研报</h1>
<div class="sub">数据来源：东方财富网　·　区间：%s ~ %s</div>
<div class="kpis"><div class="kpi"><b>%d</b><span>财报/公告份数</span></div>
<div class="kpi"><b>%d</b><span>券商研报篇数</span></div>
<div class="kpi"><b>%d</b><span>覆盖券商家数</span></div>
<div class="kpi"><b>%.0f</b><span>文件总大小 MB</span></div></div>
<h2>财报 / 公告</h2><table><thead><tr><th>日期</th><th>研究员</th><th>文件</th><th style="text-align:right">大小</th></tr></thead>
<tbody>%s</tbody></table>
<h2>券商研报</h2><table><thead><tr><th>日期</th><th>研究员</th><th>标题</th><th style="text-align:right">大小</th></tr></thead>
<tbody>%s</tbody></table>
<div class="foot">本清单由 eastmoney_downloader.py v%s 自动生成，PDF 均来自东方财富网公开数据，仅供研究参考，不构成投资建议。</div>
</div></body></html>""" % (code, code, begin, end, n_a, n_r, orgs, sz,
                          rows([m for m in manifest if m["类别"] != "研报"], False),
                          rows([m for m in manifest if m["类别"] == "研报"], True), __version__)
    with open(os.path.join(outdir, "清单.html"), "w", encoding="utf-8") as f:
        f.write(H)


# ==================== 自检 ====================
def doctor():
    log("eastmoney_downloader.py v%s  自检" % __version__)
    log("-" * 56)
    log("Python      : %s (%s)" % (platform.python_version(), sys.executable))
    log("平台        : %s %s" % (platform.system(), platform.release()))
    log("控制台编码  : %s" % (sys.stdout.encoding or "?"))
    log("curl        : %s" % (shutil.which("curl") or "未找到（仅影响极端情况下的兜底通道）"))
    try:
        import pypdf                                  # noqa: F401
        log("pypdf       : 已安装（作为研报日期的额外兜底）")
    except ImportError:
        log("pypdf       : 未安装（可选，不影响功能；核心解析已内置）")
    log("pdftotext   : %s" % (shutil.which("pdftotext") or "未找到（可选）"))
    log("-" * 56)
    for label, url, ref in (
            ("研报列表接口", "https://reportapi.eastmoney.com/report/list?cb=&pageSize=1&pageNo=1"
                             "&qType=0&industryCode=*&beginTime=2026-01-01&endTime=2026-12-31"
                             "&code=600519&rcode=&fields=&p=1&pageNum=1&pageNumber=1",
             "https://data.eastmoney.com/report/"),
            ("公告列表接口", "https://np-anotice-stock.eastmoney.com/api/security/ann?page_size=1"
                             "&page_index=1&ann_type=A&client_source=web&stock_list=600519&f_node=0&s_node=0",
             REF),
            ("F10 板块接口", "https://datacenter-web.eastmoney.com/api/data/v1/get"
                             "?reportName=RPT_F10_CORETHEME_BOARDTYPE&columns=ALL"
                             "&filter=(SECUCODE%3D%22600519.SH%22)", REF),
            ("PDF 下载节点", "https://pdf.dfcfw.com/pdf/H3_AP202609211829719676_1.pdf", REF),
            ("名称查询接口", "https://searchapi.eastmoney.com/api/suggest/get?input=%E8%8C%85%E5%8F%B0"
                             "&type=14&token=" + SEARCH_TOKEN + "&count=3",
             "https://www.eastmoney.com/")):
        t0 = time.time()
        r = _http(url, referer=ref, tries=1, timeout=20, raw=("PDF" in label))
        ok = bool(r)
        if isinstance(r, bytes):
            ok = r[:4] == b"%PDF"
        log("  %-12s %s  %.2fs  %s" % (label, MARK_OK if ok else MARK_BAD,
                                       time.time() - t0, "" if ok else "不可用"))
    n, i = quote_info("600519")
    log("  %-12s %s  %s" % ("简称/行业", MARK_OK if n else MARK_BAD,
                            ("%s / %s" % (n, i)) if n else "取不到"))
    log("-" * 56)
    log("自检结束。")


# ==================== 主流程 ====================
def _build_parser():
    p = argparse.ArgumentParser(
        prog="eastmoney_downloader.py",
        description="东方财富网 财报+研报 批量下载器 v%s（单文件 · 零依赖 · 跨平台）" % __version__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例:\n"
               "  python %(prog)s --code 600519 --period 2026 --zip\n"
               "  python %(prog)s --name 茅台 --period 2026H1 --types report\n"
               "  python %(prog)s --code 000568 --period 2026-01-01~2026-06-30\n")
    p.add_argument("--code", help="6 位股票代码，如 600519")
    p.add_argument("--name", help="公司名，如 茅台（自动解析成代码）")
    p.add_argument("--period", help="时间范围：2026 / 2026H1 / 2026H2 / 2026Q1-Q4 / "
                                    "2026-03 / 2026-01-01~2026-06-30")
    p.add_argument("--year", help="年份快捷方式（等价 --period 2026）")
    p.add_argument("--begin", help="起始日期 YYYY-MM-DD")
    p.add_argument("--end", help="结束日期 YYYY-MM-DD")
    p.add_argument("--types", default="report,ann", help="report=研报，ann=财报/公告")
    p.add_argument("--ann-filter", default="periodic", choices=["periodic", "all", "keyword"])
    p.add_argument("--ann-keyword", default="", help="ann-filter=keyword 时的标题关键词")
    p.add_argument("--name-spec", default=SPEC_REPORT, help="研报命名模板")
    p.add_argument("--name-spec-ann", default=SPEC_ANN, help="财报/公告命名模板")
    p.add_argument("--date-from", default="pdf", choices=["api", "pdf"],
                   help="api=东财发布日；pdf=优先研报自身标注日期（默认）")
    p.add_argument("--author-sep", default="_", help="多作者连接符")
    p.add_argument("--out", help="输出目录（默认 ./<公司简称><代码>_<时间>_财报研报）")
    p.add_argument("--list-only", action="store_true", help="只抓清单，不下载 PDF")
    p.add_argument("--zip", dest="do_zip", action="store_true", help="完成后打包 zip")
    p.add_argument("--sleep", type=float, default=0.6, help="下载间隔秒数（勿低于 0.5）")
    p.add_argument("--doctor", action="store_true", help="环境与接口连通性自检")
    p.add_argument("--version", action="version", version="eastmoney_downloader.py v" + __version__)
    return p


def _resolve_target(args):
    """解析股票代码，返回 (code, 公司名或简称)。解析失败直接退出。"""
    code, plain = (args.code or "").strip(), ""
    if not code and args.name:
        code, plain = resolve_by_name(args.name.strip())
        if not code:
            sys.exit("错误：未能把 %r 解析成 A 股代码，请直接提供 6 位代码（如 600519）" % args.name)
        log("已解析：%s -> %s" % (args.name, code))
    if not re.match(r"^\d{6}$", code):
        sys.exit("错误：股票代码必须是 6 位数字（如 600519），当前为 %r\n"
                 "      也可以用 --name 茅台 让程序自动解析。" % code)
    return code, plain


def _resolve_period(args):
    """解析时间范围，返回 (begin, end, label)。解析失败直接退出。"""
    begin = end = label = ""
    if args.period:
        begin, end, label = parse_period(args.period)
        if not begin:
            sys.exit("错误：无法解析时间范围 %r\n"
                     "支持：2026 / 2026H1 / 2026H2 / 2026Q3 / 2026-03 / 2026-01-01~2026-06-30"
                     % args.period)
    if args.year and not args.period:
        begin, end, label = parse_period(str(args.year))
    begin, end = args.begin or begin, args.end or end
    if not begin or not end:
        year = str(datetime.date.today().year)
        begin, end = begin or "%s-01-01" % year, end or "%s-12-31" % year
        label = label or year
    if begin > end:
        sys.exit("错误：起始日期晚于结束日期（%s > %s）" % (begin, end))
    return begin, end, label


def _resolve_types(args):
    """解析 --types，返回去重后的类型列表。非法输入直接退出。"""
    types = [t.strip() for t in args.types.split(",") if t.strip()]
    if not types:
        sys.exit("错误：--types 不能为空（report / ann / report,ann）")
    bad = [t for t in types if t not in ("report", "ann")]
    if bad:
        sys.exit("错误：--types 不支持 %s（仅 report / ann）" % ",".join(bad))
    return types


def _download_announcements(code, begin, end, company, ind, outdir, cache,
                            cache_path, args):
    """下载财报/公告，返回 manifest 条目列表。"""
    log("\n[1/2] 抓取公告列表 ...")
    category = "定期报告" if args.ann_filter == "periodic" else "公告"
    picked = [x for x in fetch_anns(code, begin, end)
              if ann_match(x, args.ann_filter, args.ann_keyword)]
    log("  按 [%s] 筛出 %d 条" % (args.ann_filter, len(picked)))

    subdir = os.path.join(outdir, "财报")
    if picked and not args.list_only:
        os.makedirs(subdir, exist_ok=True)

    entries = []
    for i, item in enumerate(sorted(picked, key=lambda z: z["notice_date"]), 1):
        title = item["title"].split(":", 1)[-1].strip()
        if company and title.startswith(company):
            title = title[len(company):].strip()
        day = item["notice_date"][:10]
        fields = {"date": ymd_to_yymmdd(day), "apdate": day, "ind": ind,
                  "company": company, "code": code, "title": title}
        name, miss = fill(args.name_spec_ann, fields)
        if miss:
            warn("  警告：财报命名缺字段 %s，标题=%s" % (",".join(miss), title))
        name += ".pdf"

        ok, size = True, 0
        if not args.list_only:
            target = os.path.join(subdir, name)
            cached = cache.get(item["art_code"])
            if cached and os.path.exists(os.path.join(subdir, cached)):
                name = cached
                ok, size = True, os.path.getsize(os.path.join(subdir, name))
            elif os.path.exists(target) and os.path.getsize(target) > MIN_PDF_SIZE:
                ok, size = True, os.path.getsize(target)
                cache[item["art_code"]] = name
            else:
                target, name = unique_path(subdir, name)
                ok, size = download("", target, urls=ann_pdf_urls(item["art_code"]))
                if ok:
                    cache[item["art_code"]] = name
            save_cache(cache_path, cache)
            time.sleep(args.sleep)

        log("  [%d/%d] %s %s" % (i, len(picked), MARK_OK if ok else MARK_BAD, name[:66]))
        entries.append({
            "类别": category, "日期": day, "标题": item["title"], "_报告名": title,
            "文件": os.path.join("财报", name) if ok else "",
            "大小KB": round(size / 1024),
            "来源链接": "https://data.eastmoney.com/notices/detail/%s/%s.html"
                        % (code, item["art_code"]),
        })
    return entries


def _download_reports(code, begin, end, company, ind, outdir, cache, cache_path,
                      date_stat, args):
    """下载研报，返回 manifest 条目列表；`date_stat` 就地累积日期来源统计。"""
    log("\n[2/2] 抓取研报列表 ...")
    reports = fetch_reports(code, begin, end)
    subdir = os.path.join(outdir, "研报")
    if reports and not args.list_only:
        os.makedirs(subdir, exist_ok=True)

    entries = []
    for i, item in enumerate(sorted(reports, key=lambda z: z["publishDate"]), 1):
        org = item.get("orgSName") or item.get("orgName") or ""
        authors = [x.strip() for x in re.split(r"[,，、;；/]", item.get("researcher") or "")
                   if x.strip()]
        publish_day = item["publishDate"][:10]
        fields = {"apdate": publish_day, "date": ymd_to_yymmdd(publish_day),
                  "ind": norm_ind(item.get("indvInduName") or item.get("industryName")) or ind,
                  "company": item.get("stockName") or company,
                  "code": item.get("stockCode") or code, "org": org,
                  "author": args.author_sep.join(authors), "title": item["title"]}
        name, miss = fill(args.name_spec, fields)
        if miss:
            warn("  警告：研报命名缺字段 %s，标题=%s" % (",".join(miss), item["title"][:30]))
        name += ".pdf"

        ok, size, final_name = False, 0, name
        if args.list_only:
            ok = True
        else:
            reuse = cache.get(item["infoCode"]) or ""
            if reuse and not os.path.exists(os.path.join(subdir, reuse)):
                reuse = ""
            if not reuse and args.date_from == "api" \
                    and os.path.exists(os.path.join(subdir, name)) \
                    and os.path.getsize(os.path.join(subdir, name)) > MIN_PDF_SIZE:
                reuse = name
            if reuse:
                final_name = reuse
                ok, size = True, os.path.getsize(os.path.join(subdir, final_name))
                cache[item["infoCode"]] = final_name
            else:
                tmp = os.path.join(subdir, ".dl_%s.pdf" % item["infoCode"])
                base_url = "https://pdf.dfcfw.com/pdf/H3_%s_1.pdf" % item["infoCode"]
                ok, size = download(base_url, tmp, urls=[
                    base_url, "%s?%d.pdf" % (base_url, int(time.time() * 1000))])
                if ok:
                    if args.date_from == "pdf":
                        own_date, source = pdf_own_date(tmp, ymd_to_yymmdd(publish_day))
                        if own_date:
                            fields["date"] = own_date
                            name, _ = fill(args.name_spec, fields)
                            name += ".pdf"
                            date_stat[source] = date_stat.get(source, 0) + 1
                        else:
                            date_stat["fallback"] = date_stat.get("fallback", 0) + 1
                    target = os.path.join(subdir, name)
                    # 缓存丢失重跑：同名且同大小 => 同一份报告，直接复用，避免 _2 副本
                    if os.path.exists(target) and os.path.getsize(target) == os.path.getsize(tmp):
                        os.remove(tmp)
                        final_name = name
                    else:
                        target, final_name = unique_path(subdir, name)
                        os.replace(tmp, target)
                    cache[item["infoCode"]] = final_name
                elif os.path.exists(tmp):
                    os.remove(tmp)
            save_cache(cache_path, cache)
            time.sleep(args.sleep)

        log("  [%d/%d] %s %s" % (i, len(reports), MARK_OK if ok else MARK_BAD, final_name[:66]))
        entries.append({
            "类别": "研报", "日期": publish_day, "机构": org, "标题": item["title"],
            "评级": item.get("emRatingName", ""), "研究员": item.get("researcher", ""),
            "文件": os.path.join("研报", final_name) if ok else "",
            "大小KB": round(size / 1024),
            "来源链接": "https://data.eastmoney.com/report/info/%s.html" % item["infoCode"],
        })
    return entries


def main():
    args = _build_parser().parse_args()

    if args.doctor:
        doctor()
        return

    code, plain = _resolve_target(args)
    begin, end, label = _resolve_period(args)
    types = _resolve_types(args)

    tag = label or begin
    outdir = os.path.abspath(args.out) if args.out else \
        os.path.abspath(os.path.join(os.getcwd(), "%s_%s_财报研报" % (code, tag)))

    # 命名模板需要公司简称/行业时先取；财报模板取不到必须中止（避免残缺文件名）
    ann_need = ("ann" in types) and any(k in args.name_spec_ann for k in ("{ind}", "{company}"))
    rep_need = ("report" in types) and any(k in args.name_spec for k in ("{ind}", "{company}"))
    company, ind0 = quote_info(code) if (ann_need or rep_need) else ("", "")
    if not company and plain:
        company = plain
    if ann_need and (not company or not ind0):
        sys.exit("错误：未能获取公司简称/行业，无法套用财报命名模板。\n"
                 "      为避免产生残缺文件名，已中止，未创建任何文件或目录。\n"
                 "      请先确认代码是否存在（沪市 6 开头 / 深市 0、3 开头 / 北交所 4、8 开头）；\n"
                 "      若确为网络抖动，稍后重试即可；或改用不含 {ind}/{company} 的 --name-spec-ann。")
    if rep_need and not company:
        warn("  提示：未取到公司简称，研报文件名将回退使用研报接口自带字段。")

    if not args.out and company:
        outdir = os.path.abspath(os.path.join(os.path.dirname(outdir),
                                              "%s%s_%s_财报研报" % (company, code, tag)))
    os.makedirs(outdir, exist_ok=True)
    cache_path = os.path.join(outdir, CACHE_FILENAME)
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, encoding="utf-8") as f:
                cache = json.load(f)
        except Exception:
            cache = {}

    log("=" * 62)
    log("标的: %s %s   区间: %s ~ %s   类型: %s" % (company or code, code, begin, end, "/".join(types)))
    log("输出: %s%s" % (outdir, "   [仅清单模式]" if args.list_only else ""))
    if "report" in types and not args.list_only:
        log("研报命名: %s   (日期口径: %s)" % (args.name_spec, args.date_from))
    log("=" * 62)

    manifest, date_stat = [], {}

    # ---- 财报 / 公告 ----
    if "ann" in types:
        manifest += _download_announcements(
            code, begin, end, company, ind0, outdir, cache, cache_path, args)

    # ---- 研报 ----
    if "report" in types:
        manifest += _download_reports(
            code, begin, end, company, ind0, outdir, cache, cache_path,
            date_stat, args)

    if not args.list_only:
        save_cache(cache_path, cache)
    with open(os.path.join(outdir, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    write_index(outdir, manifest, code, begin, end)

    okc = sum(1 for m in manifest if m["文件"])
    log("\n" + "=" * 62)
    log("完成：%d/%d 成功   输出目录：%s" % (okc, len(manifest), outdir))
    if "report" in types and not args.list_only and args.date_from == "pdf" and date_stat:
        log("研报日期来源：元数据标题 %d / 封面正文 %d / PDF生成日 %d / 回退发布日 %d"
            % (date_stat.get("title", 0), date_stat.get("text", 0),
               date_stat.get("pdfdate", 0), date_stat.get("fallback", 0)))

    if not args.list_only and manifest:
        tot, badn, badnames = verify_pdfs(outdir)
        log("校验：%d 个 PDF，损坏 %d 个%s"
            % (tot, badn, ("：" + "、".join(badnames[:5])) if badn else ""))
        if args.do_zip:
            zname = "%s%s_%s_财报研报" % (company, code, tag) if company else os.path.basename(outdir)
            zpath, n = make_zip(outdir, zname)
            log("打包：%s（%d 个文件，%.1f MB）" % (zpath, n, os.path.getsize(zpath) / 1048576))
        log("清单：%s" % os.path.join(outdir, "清单.html"))

    failed = [m for m in manifest if not m["文件"]]
    if failed and not args.list_only:
        log("")
        warn("!" * 62)
        warn("警告：有 %d/%d 个文件没下下来（重试后仍失败，通常是服务端临时故障）。"
             % (len(failed), len(manifest)))
        for m in failed[:12]:
            warn("   - %s  %s" % (m["日期"], m["标题"][:58]))
        if len(failed) > 12:
            warn("   ... 另有 %d 个" % (len(failed) - 12))
        warn("   直接重跑同一条命令即可补齐：已成功的会跳过，不会重复下载。")
        warn("!" * 62)
    log("=" * 62)

    return 2 if (failed and not args.list_only) else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit("\n已中断。重跑同一条命令即可续传。")
