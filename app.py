import io
import os
import re
import string
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


COLUMNS = [
    "비엘", "쉬퍼", "쉬퍼주소", "컨사이니", "컨사이니 사업자번호", "컨사이니 주소",
    "선명", "항차", "출발지", "도착지", "품명", "마크", "수량", "중량", "CBM",
    "원본파일명", "확인필요"
]


def group_words_to_lines(words, y_tol=3):
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    for w in words:
        cy = (w["top"] + w["bottom"]) / 2
        for line in lines:
            if abs(line["cy"] - cy) <= y_tol:
                line["words"].append(w)
                line["cy"] = (line["cy"] * line["n"] + cy) / (line["n"] + 1)
                line["n"] += 1
                break
        else:
            lines.append({"cy": cy, "n": 1, "words": [w]})

    result = []
    for line in sorted(lines, key=lambda x: x["cy"]):
        ws = sorted(line["words"], key=lambda w: w["x0"])
        result.append({
            "top": min(w["top"] for w in ws),
            "bottom": max(w["bottom"] for w in ws),
            "x0": min(w["x0"] for w in ws),
            "x1": max(w["x1"] for w in ws),
            "text": " ".join(w["text"] for w in ws),
            "words": ws,
        })
    return result


def text_from_words(words, y_tol=3):
    return "\n".join(line["text"] for line in group_words_to_lines(words, y_tol=y_tol)).strip()


def words_in_region(page, x0, top, x1, bottom, *, mode="inside"):
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    selected = []
    for w in words:
        if w["top"] < top or w["bottom"] > bottom:
            continue
        if mode == "start":
            ok = x0 <= w["x0"] < x1
        else:
            ok = w["x0"] >= x0 and w["x1"] <= x1
        if ok:
            selected.append(w)
    return selected


def text_in_region(page, x0, top, x1, bottom, *, mode="inside"):
    return text_from_words(words_in_region(page, x0, top, x1, bottom, mode=mode))


def extract_business_no(text):
    """컨사이니명 안의 000-00-00000 형식 사업자번호만 추출합니다."""
    m = re.search(r"(\d{3}-\d{2}-\d{5})", text or "")
    return m.group(1) if m else ""


def clean_company_with_paren(text):
    """컨사이니명 정리.
    - 000-00-00000 형식 사업자번호만 제거
    - 사업자번호 제거 후 남는 빈 괄호 (), （）만 제거
    - 일반 괄호 문구는 임의로 삭제하지 않음
    예: ZZZIP GUESTHOUSE（105-20-88541） -> ZZZIP GUESTHOUSE
        GOGOSS(452-64-00260) -> GOGOSS
        HOMLUX CO., LTD. 563-87-03514 -> HOMLUX CO., LTD.
    """
    text = text or ""
    text = re.sub(r"[\(（]?\s*\b\d{3}-\d{2}-\d{5}\b\s*[\)）]?", " ", text)
    text = re.sub(r"[\(（]\s*[\)）]", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" .,-")
    return text


def clean_consignee_address_and_bizno(address_text, existing_bizno=""):
    """컨사이니 주소 안에 000-00-00000 형식 사업자번호가 있으면 사업자번호 칸으로 분리합니다.
    전화번호(예: 010-3970-7762)는 000-00-00000 형식이 아니므로 주소에 그대로 유지합니다.
    """
    address_text = address_text or ""
    biz_no = existing_bizno or ""

    if not biz_no:
        found = extract_business_no(address_text)
        if found:
            biz_no = found

    if biz_no:
        # 사업자번호만 주소에서 제거. 일반 전화번호는 제거하지 않음.
        address_text = re.sub(rf"[\(（]?\s*{re.escape(biz_no)}\s*[\)）]?", " ", address_text)

    # 사업자번호 제거 후 생긴 빈 괄호만 제거
    address_text = re.sub(r"[\(（]\s*[\)）]", " ", address_text)
    address_text = re.sub(r"[ \t]+", " ", address_text)
    address_text = "\n".join(line.strip(" .,-") for line in address_text.splitlines() if line.strip(" .,-"))
    return address_text.strip(), biz_no


def clean_mark(text):
    lines = []
    for ln in (text or "").splitlines():
        ln = ln.strip()
        if not ln:
            continue
        if re.search(r"marks\s*&?\s*nos|container\s*seal", ln, re.I):
            continue
        lines.append(ln)
    return "\n".join(lines).strip()


def clean_description(text):
    stop_patterns = [
        r"^FREIGHT\b",
        r"^SHIPPED\s+ON\s+BOARD\b",
        r"^Above\s+Particulars\b",
        r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}$",
    ]
    lines = []
    seen_said_to_contain = False

    for ln in (text or "").splitlines():
        ln = re.sub(r"\b\d+(?:\.\d+)?\s*KGS\b", "", ln, flags=re.I)
        ln = re.sub(r"\b\d+(?:\.\d+)?\s*CBM\b", "", ln, flags=re.I)
        ln = ln.strip()
        if not ln:
            continue
        if any(re.search(p, ln, re.I) for p in stop_patterns):
            break

        # 품명은 SAID TO CONTAIN 아래부터 시작. 안내 문구 자체는 제외.
        if re.search(r"SHIPPER['’]?S\s+LOAD\s+COUNT", ln, re.I):
            continue
        if re.search(r"SAID\s+TO\s+CONTAIN", ln, re.I):
            seen_said_to_contain = True
            # 같은 줄 뒤에 품명이 붙는 예외가 있으면 뒤쪽만 살림
            tail = re.split(r"SAID\s+TO\s+CONTAIN", ln, flags=re.I)[-1].strip(" :-")
            if tail:
                lines.append(tail)
            continue

        lines.append(ln)

    return "\n".join(lines).strip()


def clean_description_words(words):
    clean = []
    for w in words:
        t = w["text"].strip()
        if re.fullmatch(r"\d+(?:\.\d+)?\s*KGS", t, re.I):
            continue
        if re.fullmatch(r"\d+(?:\.\d+)?\s*CBM", t, re.I):
            continue
        if re.fullmatch(r"\d+\s*(?:PKGS?|CTNS?|CARTONS?|PCS)", t, re.I):
            continue
        clean.append(w)
    return clean_description(text_from_words(clean))


def safe_search(pattern, text, group=1, flags=re.I):
    m = re.search(pattern, text or "", flags)
    return m.group(group).strip() if m else ""


def normalize_port_code(text):
    """Convert common port text to requested customs-style code."""
    raw = (text or "").strip()
    key = re.sub(r"\s+", " ", raw.upper().replace("，", ","))
    key = key.replace(" ", "")
    mapping = {
        "SHIDAO": "CNSHD",
        "SHIDAOCHINA": "CNSHD",
        "SHIDAO,CHINA": "CNSHD",
        "YANTAICHINA": "CNYNT",
        "YANTAI,CHINA": "CNYNT",
        "WEIHAICHINA": "CNWEI",
        "WEIHAI,CHINA": "CNWEI",
        "GUNSAN,KOREA": "KRKUV",
        "GUNSANKOREA": "KRKUV",
        "INCHEON,KOREA": "KRINC",
        "INCHEONKOREA": "KRINC",
        "INCHON": "KRINC",
        "INCHON,KOREA": "KRINC",
        "INCHONKOREA": "KRINC",
    }
    return mapping.get(key, raw)



def split_vessel_voyage_text(text):
    """선명/항차 보정.
    - HANSUNG으로 시작하는 경우 HANSUNG INCHEON까지 선명으로 고정
    - 뒤의 숫자 3~4자리(+E 선택)를 항차로 분리
    """
    raw = re.sub(r"\s+", " ", text or " ").strip()
    m = re.search(r"\bHANSUNG\s+INCHEON\s+(\d{3,4}E?)\b", raw, re.I)
    if m:
        voyage = m.group(1).upper()
        if re.fullmatch(r"\d{3,4}", voyage):
            voyage = f"{voyage}E"
        return "HANSUNG INCHEON", voyage

    m = re.search(r"\b([A-Z][A-Z0-9]*(?:\s+[A-Z0-9]+)*?)\s+(\d{3,4}E?)\b", raw, re.I)
    if m:
        vessel = re.sub(r"\s+", " ", m.group(1)).strip()
        voyage = m.group(2).upper()
        if re.fullmatch(r"\d{3,4}", voyage):
            voyage = f"{voyage}E"
        return vessel, voyage
    return "", ""

def extract_qty_from_mark(mark_text):
    """Fallback quantity recognition from mark ranges, e.g. C/T:1-71, FR-xxx-001 - FR-xxx-037, LJ1-63."""
    txt = (mark_text or "").replace("\n", " ")
    # C/T:1-71, CT: 1 ~ 71
    m = re.search(r"C\s*/?\s*T\s*[:：]?\s*(\d+)\s*[-~]\s*(\d+)", txt, re.I)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        return str(b - a + 1) if b >= a else str(b)
    # FR-1519594-001 - FR-1519594-037 / GR-...001 - ...005
    nums = re.findall(r"(?:^|[^A-Z0-9])(\d{1,4})(?=\s*(?:$|[^A-Z0-9]))", txt, re.I)
    if len(nums) >= 2:
        # use the last two small sequence numbers if it looks like a range
        a, b = int(nums[-2]), int(nums[-1])
        if 0 <= a <= b <= 9999:
            return str(b - a + 1)
    # LJ1-63, JZF9-1-9: use last number as count when no explicit range exists
    tail = re.search(r"[-\s](\d{1,4})\s*$", txt)
    if tail:
        return str(int(tail.group(1)))
    return ""






def extract_wooyoung_pdf(page, filename):
    """WOOYOUNG / INUF 텍스트형 BL 양식 전용 추출.
    해당 양식은 칸이 명확하므로 일반 BL 로직을 건드리지 않고 좌측/우측 구역 기준으로만 보정합니다.
    """
    warnings = []
    w, h = page.width, page.height
    text = page.extract_text() or ""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    def region_lines(x0, top, x1, bottom):
        txt = text_in_region(page, x0, top, x1, bottom, mode="inside")
        return [re.sub(r"\s+", " ", ln).strip() for ln in txt.splitlines() if ln.strip()]

    def remove_labels(block, *labels):
        out = []
        for ln in block:
            if any(re.fullmatch(lab, ln, re.I) for lab in labels):
                continue
            out.append(ln)
        return out

    # B/L No.: 우측 상단 또는 좌측 하단의 Bill of Lading No. 아래 숫자
    bl = ""
    bl_blocks = [
        region_lines(w * 0.44, h * 0.125, w * 0.60, h * 0.165),
        region_lines(w * 0.07, h * 0.865, w * 0.26, h * 0.900),
    ]
    for block in bl_blocks:
        cand = " ".join(block)
        m = re.search(r"\b(\d{10,20})\b", cand)
        if m:
            bl = m.group(1)
            break
    if not bl:
        # fallback: 파일 전체에서 긴 숫자 후보 중 B/L에 가까운 값 사용
        nums = re.findall(r"\b\d{12,20}\b", text)
        if nums:
            bl = nums[-1]

    # SHIPPER: 좌측 상단 SHIPPER 칸. 첫 줄=쉬퍼, 나머지=쉬퍼주소
    shipper_block = remove_labels(
        region_lines(w * 0.07, h * 0.085, w * 0.47, h * 0.164),
        r"SHIPPER"
    )
    shipper = shipper_block[0] if shipper_block else ""
    shipper_addr = "\n".join(shipper_block[1:]).strip()

    # CONSIGNEE: 좌측 CONSIGNEE 칸. 첫 줄=컨사이니, 나머지=컨사이니주소
    consignee_block = remove_labels(
        region_lines(w * 0.07, h * 0.155, w * 0.47, h * 0.235),
        r"CONSIGNEE", r"NOTIFY PARTY", r"SAME AS CONSIGNEE.*"
    )
    consignee_raw = consignee_block[0] if consignee_block else ""
    consignee_no = extract_business_no(consignee_raw)
    consignee = clean_company_with_paren(consignee_raw)
    consignee_addr = "\n".join(consignee_block[1:]).strip()
    consignee_addr, consignee_no = clean_consignee_address_and_bizno(consignee_addr, consignee_no)

    # Vessel/Voyage No. 칸 아래 값: 마지막 3~4자리+E = 항차, 앞부분 = 선명
    vessel_region = " ".join(region_lines(w * 0.07, h * 0.335, w * 0.34, h * 0.365))
    vessel_region = re.sub(r"Vessel\s*/\s*Voyage\s+No\.?", " ", vessel_region, flags=re.I)
    vessel, voyage = split_vessel_voyage_text(vessel_region)

    # PORT OF LOADING / DISCHARGE 칸 아래 값
    pol_raw = " ".join(region_lines(w * 0.07, h * 0.374, w * 0.21, h * 0.402))
    pod_raw = " ".join(region_lines(w * 0.24, h * 0.374, w * 0.45, h * 0.402))
    pol_raw = re.sub(r"PORT\s+OF\s+LOADING", " ", pol_raw, flags=re.I).strip()
    pod_raw = re.sub(r"PORT\s+OF\s+DISCHARGE", " ", pod_raw, flags=re.I).strip()
    pol = normalize_port_code(pol_raw)
    pod = normalize_port_code(pod_raw)

    # 기존 표 영역 인식은 유지: N/M 줄부터 수량/품명/중량/CBM 추출
    start_idx = None
    for i, ln in enumerate(lines):
        if re.search(r"^N/M\b", ln, re.I):
            start_idx = i
            break

    mark = ""
    pkg = ""
    weight = ""
    cbm = ""
    description_lines = []

    if start_idx is not None:
        first = lines[start_idx]
        mark = "N/M"

        m_pkg = re.search(r"\b(\d+)\s*(?:CTNS?|PKGS?|PACKAGES?|CARTONS?|PCS)\b", first, re.I)
        if m_pkg:
            pkg = m_pkg.group(1)

        desc_first = re.sub(r"^N/M\s*", "", first, flags=re.I)
        desc_first = re.sub(r"\b\d+\s*(?:CTNS?|PKGS?|PACKAGES?|CARTONS?|PCS)\b", "", desc_first, flags=re.I).strip()
        desc_first = re.sub(r"\b\d+(?:\.\d+)?\s*KGS?\b", "", desc_first, flags=re.I).strip()
        desc_first = re.sub(r"\b\d+(?:\.\d+)?\s*CBM\b", "", desc_first, flags=re.I).strip()
        if desc_first:
            description_lines.append(desc_first)

        for ln in lines[start_idx + 1:]:
            if re.search(r"SHIPPER['’]?S\s+LOAD|S\.T\.C|SHIPPED\s+ON\s+BOARD|FREIGHT\s+PREPAID|Total Number|SAY TOTAL", ln, re.I):
                break
            ln_clean = re.sub(r"\b\d+(?:\.\d+)?\s*KGS?\b", "", ln, flags=re.I)
            ln_clean = re.sub(r"\b\d+(?:\.\d+)?\s*CBM\b", "", ln_clean, flags=re.I).strip()
            if ln_clean:
                description_lines.append(ln_clean)

    m_w = re.search(r"\b(\d+(?:\.\d+)?)\s*KGS?\b", text, re.I)
    if m_w:
        weight = m_w.group(1)
    m_c = re.search(r"\b(\d+(?:\.\d+)?)\s*CBM\b", text, re.I)
    if m_c:
        cbm = m_c.group(1)

    description = "\n".join(description_lines).strip()

    if not bl:
        warnings.append("비엘 미인식")
    if not shipper:
        warnings.append("쉬퍼 미인식")
    if not consignee:
        warnings.append("컨사이니 미인식")
    if not vessel:
        warnings.append("선명 미인식")
    if not voyage:
        warnings.append("항차 미인식")
    if not pol:
        warnings.append("출발지 미인식")
    if not pod:
        warnings.append("도착지 미인식")
    if not mark:
        warnings.append("마크 미인식")
    if not description:
        warnings.append("품명 미인식")
    if not weight:
        warnings.append("중량 확인필요")
    if not cbm:
        warnings.append("CBM 확인필요")
    if not pkg:
        warnings.append("수량 확인필요")

    return {
        "비엘": bl,
        "쉬퍼": shipper,
        "쉬퍼주소": shipper_addr,
        "컨사이니": consignee,
        "컨사이니 사업자번호": consignee_no,
        "컨사이니 주소": consignee_addr,
        "선명": vessel,
        "항차": voyage,
        "출발지": pol,
        "도착지": pod,
        "품명": description,
        "마크": mark,
        "수량": pkg,
        "중량": weight,
        "CBM": cbm,
        "원본파일명": filename,
        "확인필요": ", ".join(warnings),
    }


def _strip_label_text(text):
    """라벨 문구가 같이 잡힌 경우 값만 남기기."""
    text = text or ""
    text = re.sub(r"^.*?\b(?:Consignee|Shipper|Notify\s+party|Ocean\s+vessel|Port\s+of\s+loading|Port\s+of\s+discharge|Place\s+of\s+receipt|Place\s+of\s+delivery)\b.*?:?", "", text, flags=re.I).strip()
    return text


def _normalize_space_keep_lines(lines):
    cleaned = []
    for ln in lines:
        ln = re.sub(r"\s+", " ", (ln or "")).strip()
        if ln:
            cleaned.append(ln)
    return cleaned


def _block_by_left_labels(page, start_pattern, end_pattern, *, x_limit_ratio=0.55):
    """SILKROAD 특수 BL용: 왼쪽 상단 영역의 제목 기준으로 블록 추출."""
    w = page.width
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    lines = group_words_to_lines(words, y_tol=4)
    left_lines = [ln for ln in lines if ln["x0"] < w * x_limit_ratio]

    start_y = None
    end_y = None
    for ln in left_lines:
        if start_y is None and re.search(start_pattern, ln["text"], re.I):
            start_y = ln["bottom"]
            continue
        if start_y is not None and re.search(end_pattern, ln["text"], re.I):
            end_y = ln["top"]
            break
    if start_y is None:
        return []
    if end_y is None:
        end_y = page.height

    block = []
    for ln in left_lines:
        if ln["top"] >= start_y - 1 and ln["bottom"] <= end_y + 1:
            # 라벨 줄이 같이 걸리는 경우 제외
            if re.search(start_pattern, ln["text"], re.I) or re.search(end_pattern, ln["text"], re.I):
                continue
            block.append(ln["text"])
    return _normalize_space_keep_lines(block)


def _extract_after_label_near(page, label_pattern, *, x0_ratio=0, x1_ratio=0.55, y_after=None, y_before=None):
    """라벨과 같은 줄 또는 바로 아래 줄에 있는 값을 추출."""
    w, h = page.width, page.height
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    lines = group_words_to_lines(words, y_tol=4)
    candidates = []
    for i, ln in enumerate(lines):
        if not (w * x0_ratio <= ln["x0"] <= w * x1_ratio):
            continue
        if y_after is not None and ln["top"] < y_after:
            continue
        if y_before is not None and ln["top"] > y_before:
            continue
        if re.search(label_pattern, ln["text"], re.I):
            # 같은 줄 뒤쪽에 값이 같이 있으면 우선 사용
            same = re.sub(label_pattern + r".*?", "", ln["text"], flags=re.I).strip(" :")
            if same:
                candidates.append(same)
            # 다음 1~2줄 값도 확인
            for nxt in lines[i+1:i+3]:
                if nxt["top"] - ln["bottom"] < h * 0.045:
                    candidates.append(nxt["text"])
            break
    return " ".join(_normalize_space_keep_lines(candidates)).strip()


def extract_silkroad_pdf(page, filename):
    """SREJY / SILKROAD LOGISTICS 특수 BL 전용 추출.
    일반 BL 로직은 건드리지 않고, 해당 양식일 때만 보정합니다.
    """
    warnings = []
    w, h = page.width, page.height
    text = page.extract_text() or ""
    flat = re.sub(r"\s+", " ", text)

    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    lines_obj = group_words_to_lines(words, y_tol=4)
    line_texts = [ln["text"] for ln in lines_obj]

    # B/L No.
    bl = safe_search(r"\b(SREJY[A-Z0-9]{6,}|[A-Z]{3,}[A-Z0-9]{8,})\b", text)

    # 상단 좌측 실제 값 기준 추출: pdfplumber에서 라벨은 누락될 수 있어 값의 위치/패턴을 기준으로 보정
    left_top = [ln for ln in lines_obj if ln["x0"] < w * 0.45 and ln["top"] < h * 0.31]

    # Shipper: 상단 첫 회사명부터 사업자번호 포함 컨사이니 시작 전까지
    shipper_lines = []
    for ln in left_top:
        t = ln["text"].strip()
        if not t or t == bl:
            continue
        if extract_business_no(t):
            break
        # B/L No. 우측 텍스트 방지
        if ln["x0"] < w * 0.10:
            shipper_lines.append(t)
    shipper = shipper_lines[0] if shipper_lines else ""
    shipper_addr = "\n".join(shipper_lines[1:]).strip()

    # Consignee: 사업자번호 있는 줄부터 Notify(SILKROAD LOGISTICS CO., LTD) 전까지
    consignee_lines = []
    collecting = False
    for ln in left_top:
        t = ln["text"].strip()
        if not collecting and extract_business_no(t):
            collecting = True
        if collecting:
            # Notify Party 값 시작 전까지만
            if t.upper().startswith("SILKROAD LOGISTICS CO") and not extract_business_no(t):
                break
            consignee_lines.append(t)
    consignee_raw = consignee_lines[0] if consignee_lines else ""
    consignee_no = extract_business_no(consignee_raw)
    consignee = clean_company_with_paren(consignee_raw)
    consignee_addr = "\n".join(consignee_lines[1:]).strip()
    consignee_addr, consignee_no = clean_consignee_address_and_bizno(consignee_addr, consignee_no)

    # 선명/항차: HUADONG PEARL8 7264E 처럼 한 줄에 붙거나 뒤에 항구가 붙어도 분리
    vessel = ""
    voyage = ""
    for src in line_texts + [flat]:
        m = re.search(r"\b([A-Z][A-Z0-9]*(?:\s+[A-Z0-9]+)*?)\s+([0-9]{3,5}E)\b", src, re.I)
        if m:
            cand_v = re.sub(r"\s+", " ", m.group(1)).strip()
            cand_y = m.group(2).strip()
            if re.search(r"HUADONG|PEARL|HANSUNG|GRAND|PEACE|INCHEON", cand_v, re.I):
                vessel, voyage = cand_v, cand_y
                break
    if re.fullmatch(r"\d+", voyage or ""):
        voyage = f"{voyage}E"

    # 출발지/도착지: 해당 특수 양식은 SHIDAO / INCHEON 값을 코드로 변환
    pol_raw = "SHIDAO,CHINA" if re.search(r"SHIDAO\s*,?\s*CHINA", text, re.I) else ""
    pod_raw = "INCHEON,KOREA" if re.search(r"INCHEON\s*,?\s*KOREA", text, re.I) else ""
    pol = normalize_port_code(pol_raw)
    pod = normalize_port_code(pod_raw)

    # 마크: 표 왼쪽 마크 값만. 수량/품명 안내문구가 같은 줄에 붙으면 제거
    mark_candidates = []
    for ln in lines_obj:
        t = ln["text"].strip()
        if ln["top"] > h * 0.43 and ln["top"] < h * 0.58 and ln["x0"] < w * 0.20:
            if re.match(r"^(SILK|SMARTGO|[A-Z]{2,}[-A-Z0-9]*)", t, re.I):
                t = re.split(r"\b\d+\s*(?:PACKAGES?|PKGS?|CTNS?|CARTONS?|CARTON|PCS|BOXES|BOX)\b", t, flags=re.I)[0].strip()
                t = re.split(r"\bSAID\s+TO\s+CONTAIN\b|SHIPPER", t, flags=re.I)[0].strip()
                if t:
                    mark_candidates.append(t)
    mark = clean_mark("\n".join(mark_candidates))

    # 수량/중량/CBM: 단위 앞 숫자만
    m_pkg = re.search(r"\b(\d+(?:\.\d+)?)\s*(?:PACKAGES?|PKGS?|CTNS?|CARTONS?|CARTON|PCS|BOXES|BOX)\b", text, re.I)
    pkg = m_pkg.group(1) if m_pkg else ""
    m_w = re.search(r"\b(\d+(?:\.\d+)?)\s*KGS?\b", text, re.I)
    weight = m_w.group(1) if m_w else ""
    m_c = re.search(r"\b(\d+(?:\.\d+)?)\s*CBM\b", text, re.I)
    cbm = m_c.group(1) if m_c else ""

    # 품명: SAID TO CONTAIN 아래 실제 품명줄만. 좌표/문구 두 방식 병행
    desc_lines = []
    # 1순위: 표 품명칸 좌표 기준. 안내문구, 중량/CBM, SURRENDERED 제외
    for ln in lines_obj:
        t = ln["text"].strip()
        if ln["top"] > h * 0.48 and ln["top"] < h * 0.57 and w * 0.25 < ln["x0"] < w * 0.60:
            if re.search(r"SHIPPER|SAID\s+TO\s+CONTAIN|SURRENDERED|LOADED|FREIGHT|KGS?|CBM|PACKAGES?", t, re.I):
                continue
            if t:
                desc_lines.append(t)
    description = "\n".join(desc_lines).strip()
    # 2순위: 텍스트 기준
    if not description:
        m = re.search(r"SAID\s+TO\s+CONTAIN\s*:?\s*(.*?)(?:\n\s*(?:SURRENDERED|LOADED\s+ON\s+BOARD|FREIGHT\s+COLLECT|CFS/CFS|DESTINATION|AS\s+ARRANGED)\b)", text, flags=re.I | re.S)
        if m:
            description = "\n".join(
                ln.strip()
                for ln in m.group(1).splitlines()
                if ln.strip() and not re.search(r"SHIPPER|SAID\s+TO\s+CONTAIN", ln, re.I)
            ).strip()

    if not bl:
        warnings.append("비엘 미인식")
    if not consignee_addr:
        warnings.append("컨사이니 주소 확인필요")
    if not vessel:
        warnings.append("선명 확인필요")
    if not voyage:
        warnings.append("항차 확인필요")
    if not pod:
        warnings.append("도착지 확인필요")
    if not mark:
        warnings.append("마크 미인식")
    if not description:
        warnings.append("품명 미인식")
    if not weight:
        warnings.append("중량 확인필요")
    if not cbm:
        warnings.append("CBM 확인필요")
    if not pkg:
        warnings.append("수량 확인필요")

    return {
        "비엘": bl,
        "쉬퍼": shipper,
        "쉬퍼주소": shipper_addr,
        "컨사이니": consignee,
        "컨사이니 사업자번호": consignee_no,
        "컨사이니 주소": consignee_addr,
        "선명": vessel,
        "항차": voyage,
        "출발지": pol,
        "도착지": pod,
        "품명": description,
        "마크": mark,
        "수량": pkg,
        "중량": weight,
        "CBM": cbm,
        "원본파일명": filename,
        "확인필요": ", ".join(warnings),
    }



def fix_split_business_no_text(text):
    """452-\n81-03361처럼 줄바꿈으로 끊어진 사업자번호를 452-81-03361로 보정합니다."""
    text = text or ""
    text = re.sub(r"(\d{3})-\s*\n\s*(\d{2}-\d{5})", r"\1-\2", text)
    text = re.sub(r"(\d{3})-\s+(\d{2}-\d{5})", r"\1-\2", text)
    return text


def extract_first_bl_number(text):
    text = text or ""
    # HKD 계열은 영문+숫자 혼합 BL 우선
    m = re.search(r"\b(HKD[A-Z0-9]{8,25})\b", text, re.I)
    if m:
        return m.group(1).upper()
    m = re.search(r"Bill\s*of\s*Lading\s*No\.?\s*\n\s*([A-Z0-9]{8,30})", text, re.I)
    if m:
        return m.group(1).strip()
    m = re.search(r"\b([A-Z]{2,}[A-Z0-9]{6,})\b", text, re.I)
    return m.group(1).strip() if m else ""


def normalize_broken_unit_text(text):
    text = text or ""
    # PDF 텍스트 겹침으로 2C TN S / 2 BO XE S처럼 끊긴 단위를 복구
    text = re.sub(r"(\d+)\s*C\s*T\s*N\s*S", r"\1CTNS", text, flags=re.I)
    text = re.sub(r"(\d+)\s*P\s*K\s*G\s*S", r"\1PKGS", text, flags=re.I)
    text = re.sub(r"(\d+)\s*P\s*L\s*T\s*S", r"\1PLTS", text, flags=re.I)
    text = re.sub(r"(\d+)\s+BO\s+\d+\s*C\s*XE\s*TN\s*S\s*S", r"\1 BOXES", text, flags=re.I)
    text = re.sub(r"(\d+)\s*BO\s*XE\s*S", r"\1BOXES", text, flags=re.I)
    return text


def clean_line_remove_bl(line, bl):
    line = line or ""
    if bl:
        line = line.replace(bl, " ")
    return re.sub(r"\s+", " ", line).strip(" ,")


def extract_xingwen_hkd_pdf(page, filename):
    """HKD / XINGWEN / HAOKUNDA 계열 특수 BL 전용 추출.
    일반 BL은 건드리지 않고, 이 계열만 좌상단 Shipper/Consignee 칸을 줄 흐름 기준으로 읽습니다.
    """
    warnings = []
    raw_text = page.extract_text() or ""
    text = normalize_broken_unit_text(fix_split_business_no_text(raw_text))
    w, h = page.width, page.height
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    lines_obj = group_words_to_lines(words, y_tol=3)

    bl = extract_first_bl_number(text)

    # 좌측 상단 라인들: 오른쪽으로 조금 넘어간 텍스트도 같은 줄이면 포함하되, BL 번호는 제거
    geo_lines = []
    for ln in lines_obj:
        t = normalize_broken_unit_text(clean_line_remove_bl(ln["text"], bl))
        if not t:
            continue
        geo_lines.append({**ln, "text": t})

    # Shipper: 상단 첫 구역. x 좌측 시작 라인만 사용, BL 번호가 같은 줄에 붙으면 제거됨.
    shipper_lines = []
    for ln in geo_lines:
        if ln["top"] < h * 0.105 and ln["x0"] < w * 0.18:
            t = ln["text"].strip()
            if re.fullmatch(r"SHIPPER", t, re.I):
                continue
            shipper_lines.append(t)
    shipper = shipper_lines[0] if shipper_lines else ""
    shipper_addr = "\n".join(shipper_lines[1:]).strip()

    # Consignee: 선으로 나뉜 컨사이니 구역. 첫 줄=컨사이니, 나머지=주소.
    consignee_lines = []
    for ln in geo_lines:
        if h * 0.105 <= ln["top"] < h * 0.215 and ln["x0"] < w * 0.18:
            t = ln["text"].strip()
            if re.fullmatch(r"CONSIGNEE", t, re.I):
                continue
            # Notify 라벨/값은 컨사이니 주소에서 제외
            if re.fullmatch(r"NOTIFY\s+PARTY", t, re.I) or t.upper().startswith("SAME AS ABOVE"):
                continue
            consignee_lines.append(t)

    consignee_raw = ""
    addr_start = 1
    if consignee_lines:
        if len(consignee_lines) >= 2 and re.search(r"\d{3}-$", consignee_lines[0]) and re.fullmatch(r"\d{2}-\d{5}", consignee_lines[1]):
            consignee_raw = consignee_lines[0] + consignee_lines[1]
            addr_start = 2
        else:
            consignee_raw = consignee_lines[0]
            addr_start = 1
    consignee_zone_text = "\n".join(consignee_lines)
    # 컨사이니 첫 줄에는 사업자번호가 있으면 그대로 유지
    consignee = re.sub(r"\s+", " ", consignee_raw).strip()
    consignee_no = extract_business_no(consignee_zone_text) or extract_business_no(consignee)
    consignee_addr = "\n".join(consignee_lines[addr_start:]).strip()
    consignee_addr, consignee_no = clean_consignee_address_and_bizno(consignee_addr, consignee_no)

    # 선명/항차
    vessel, voyage = split_vessel_voyage_text(text)

    # 출발지/도착지: Ocean Vessel 라인 기준. 같은 줄에 WEIHAI CHINA이 붙은 경우도 처리.
    pol_raw = ""
    pod_raw = ""
    for i, ln in enumerate(geo_lines):
        t = ln["text"]
        if re.search(r"HANSUNG\s+INCHEON\s+\d{3,4}E?", t, re.I):
            after = re.split(r"HANSUNG\s+INCHEON\s+\d{3,4}E?", t, flags=re.I)[-1].strip()
            if after:
                pol_raw = after
            # 다음 1~2개 라인에서 POL/POD 후보 찾기
            for nxt in geo_lines[i+1:i+4]:
                nt = nxt["text"].strip()
                if not pol_raw and re.search(r"WEIHAI|SHIDAO|YANTAI", nt, re.I):
                    pol_raw = nt
                if not pod_raw and re.search(r"INCHEON|INCHON", nt, re.I):
                    pod_raw = nt
            break
    if not pol_raw:
        m = re.search(r"\b(WEIHAI\s+CHINA|WEIHAI,\s*CHINA|SHIDAO(?:,?\s*CHINA)?)\b", text, re.I)
        pol_raw = m.group(1) if m else ""
    if not pod_raw:
        m = re.search(r"\b(INCHEON\s*,?\s*KOREA|INCHON\s*,?\s*KOREA)\b", text, re.I)
        pod_raw = m.group(1) if m else ""
    pol = normalize_port_code(pol_raw)
    pod = normalize_port_code(pod_raw)

    # 수량/중량/CBM: 단위 앞 숫자만. 겹쳐 깨진 2C TN S도 보정된 text에서 추출.
    m_pkg = re.search(r"\b(\d+(?:\.\d+)?)\s*(?:CTNS?|CARTONS?|PKGS?|PACKAGES?|PLTS?|BOXES?|PCS)\b", text, re.I)
    pkg = m_pkg.group(1) if m_pkg else ""
    m_w = re.search(r"\b(\d+(?:\.\d+)?)\s*KGS?\b", text, re.I)
    weight = m_w.group(1) if m_w else ""
    m_c = re.search(r"\b(\d+(?:\.\d+)?)\s*CBM\b", text, re.I)
    cbm = m_c.group(1) if m_c else ""

    # 품명: SAID TO CONTAIN 아래 실제 품명만. SURRENDERED 제외.
    description_lines = []
    in_desc = False
    for ln in geo_lines:
        if ln["top"] < h * 0.38:
            continue
        t = normalize_broken_unit_text(ln["text"].strip())
        up = t.upper()
        if re.search(r"SAID\s+TO\s+CONTAIN", up):
            in_desc = True
            tail = re.split(r"SAID\s+TO\s+CONTAIN", t, flags=re.I)[-1].strip(" :-")
            if tail:
                description_lines.append(tail)
            continue
        if not in_desc:
            continue
        if re.search(r"^(SURRENDERED|COPY|FREIGHT\b|SHIPPED\b|LOADED\b|ONE\b|SAY\b|TOTAL\b|YST\b|E-MAIL|TEL\b|FAX\b|\d{4}[-./]\d{1,2}[-./]\d{1,2})", up):
            break
        if re.search(r"\b(?:KGS?|CBM|CTNS?|PKGS?|PLTS?|BOXES?)\b", up):
            continue
        # E04617 TELESCOPE처럼 마크/실번호와 품명이 같은 줄에 겹친 경우 앞 코드는 제외하고 뒤 품명만 살림
        m_code_desc = re.match(r"^(?:[A-Z]{4}\d{7}|[A-Z]{1,4}\d{1,5}[-~]\d{1,5}|E\d{3,}|\d{4,6})\s+(.+)$", t, re.I)
        if m_code_desc:
            tail_desc = m_code_desc.group(1).strip()
            if tail_desc and not re.search(r"SAID\s+TO\s+CONTAIN", tail_desc, re.I):
                description_lines.append(tail_desc)
            continue
        if re.match(r"^(OL-|[A-Z]{4}\d{7}\b|[A-Z]{1,4}\d{1,5}[-~]\d{1,5}|E\d{3,}|\d{4,6}$)", t, re.I):
            continue
        description_lines.append(t)
    description = "\n".join(description_lines).strip()

    # 마크: 표 왼쪽/연속 텍스트. OL- 한 줄은 구역을 넘어가도 한 마크로 유지.
    mark_lines = []
    for ln in geo_lines:
        if ln["top"] < h * 0.39:
            continue
        t = normalize_broken_unit_text(ln["text"].strip())
        up = t.upper()
        if re.search(r"^(FREIGHT\b|SURRENDERED|COPY|SHIPPED\b|ONE\b|SAY\b|YST\b|E-MAIL|TEL\b|FAX\b)", up):
            break
        if up.startswith("OL-"):
            t = re.split(r"SHIPPER", t, flags=re.I)[0].strip()
            t = re.sub(r"(\d+)BOXES", r"\1 BOXES", t, flags=re.I)
            mark_lines.append(t)
            continue
        # HHXU3325532 SAID TO CONTAIN처럼 마크와 안내문구가 같은 줄이면 앞 마크는 살림
        if re.search(r"SAID\s+TO\s+CONTAIN", up):
            prefix_said = re.split(r"SAID\s+TO\s+CONTAIN", t, flags=re.I)[0].strip()
            if prefix_said and re.match(r"^([A-Z]{4}\d{7}|[A-Z]{1,4}\d{1,5}[-~]\d{1,5}|E\d{3,}|\d{4,6})$", prefix_said, re.I):
                mark_lines.append(prefix_said)
            continue
        if ln["x0"] < w * 0.18:
            prefix = re.split(r"\b\d+(?:\.\d+)?\s*(?:CTNS?|CARTONS?|PKGS?|PACKAGES?|PLTS?|BOXES?|PCS|KGS?|CBM)\b|SHIPPER", t, flags=re.I)[0].strip()
            if prefix and not re.search(r"HANSUNG|WEIHAI|INCHEON|KOREA", prefix, re.I):
                mark_lines.append(prefix)
        # 품명 뒤쪽의 컨테이너/실번호 계열. E04617 TELESCOPE면 E04617만 마크로 사용.
        else:
            m_mark = re.match(r"^([A-Z]{4}\d{7}|[A-Z]{1,4}\d{1,5}[-~]\d{1,5}|E\d{3,}|\d{4,6})(?:\s+.+)?$", t, re.I)
            if m_mark:
                mark_lines.append(m_mark.group(1))
    mark = clean_mark("\n".join(dict.fromkeys([m for m in mark_lines if m]).keys()))

    if not bl:
        warnings.append("비엘 미인식")
    if not shipper:
        warnings.append("쉬퍼 확인필요")
    if not consignee:
        warnings.append("컨사이니 확인필요")
    if not consignee_addr:
        warnings.append("컨사이니 주소 확인필요")
    if not vessel:
        warnings.append("선명 확인필요")
    if not voyage:
        warnings.append("항차 확인필요")
    if not description:
        warnings.append("품명 미인식")
    if not pkg:
        warnings.append("수량 확인필요")
    if not weight:
        warnings.append("중량 확인필요")
    if not cbm:
        warnings.append("CBM 확인필요")

    return {
        "비엘": bl,
        "쉬퍼": shipper,
        "쉬퍼주소": shipper_addr,
        "컨사이니": consignee,
        "컨사이니 사업자번호": consignee_no,
        "컨사이니 주소": consignee_addr,
        "선명": vessel,
        "항차": voyage,
        "출발지": pol,
        "도착지": pod,
        "품명": description,
        "마크": mark,
        "수량": pkg,
        "중량": weight,
        "CBM": cbm,
        "원본파일명": filename,
        "확인필요": ", ".join(warnings),
    }


def extract_one_pdf(file_bytes, filename, desc_right_ratio=0.89, table_bottom_ratio=0.69):
    warnings = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        if not pdf.pages:
            return {**{c: "" for c in COLUMNS}, "원본파일명": filename, "확인필요": "PDF 페이지 없음"}

        page = pdf.pages[0]
        all_text = page.extract_text() or ""


        # HKD / XINGWEN / HAOKUNDA 계열 특수 양식
        if re.search(r"\bHKD[A-Z0-9]{8,25}\b", all_text.upper()) and ("HANSUNG INCHEON" in all_text.upper()):
            return extract_xingwen_hkd_pdf(page, filename)

        # SILKROAD / SREJY 특수 양식은 WOOYOUNG 조건보다 먼저 분기해야 함
        # (일부 PDF 텍스트에 MUTIMODAL 문구가 같이 잡혀도 SILKROAD로 처리)
        if ("SILKROAD LOGISTICS" in all_text.upper()) or re.search(r"\bSREJY[A-Z0-9]{6,}\b", all_text.upper()):
            return extract_silkroad_pdf(page, filename)

        # WOOYOUNG / INUF 특수 양식
        if ("WOOYOUNG (CHINA) LOGISTICS" in all_text.upper()) or ("MUTIMODAL TRANSPORT BILL OF LADING" in all_text.upper()):
            return extract_wooyoung_pdf(page, filename)

        w, h = page.width, page.height

        bl_region = text_in_region(page, w * 0.58, 0, w, h * 0.18)
        bl = safe_search(r"\b([A-Z0-9]{8,})\b", bl_region)

        shipper_block = text_in_region(page, 0, h * 0.02, w * 0.57, h * 0.10)
        shipper_lines = [x.strip() for x in shipper_block.splitlines() if x.strip()]
        shipper = shipper_lines[0] if shipper_lines else ""
        shipper_addr = "\n".join(shipper_lines[1:]).strip()
        if shipper and not shipper_addr:
            shipper_addr = "CHINA"

        consignee_block = text_in_region(page, 0, h * 0.09, w * 0.57, h * 0.16)
        consignee_lines = [x.strip() for x in consignee_block.splitlines() if x.strip()]
        consignee_raw = consignee_lines[0] if consignee_lines else ""
        consignee_no = extract_business_no(consignee_raw)
        consignee = clean_company_with_paren(consignee_raw)
        consignee_addr = "\n".join(consignee_lines[1:])
        consignee_addr, consignee_no = clean_consignee_address_and_bizno(consignee_addr, consignee_no)

        vessel_region = text_in_region(page, 0, h * 0.31, w * 0.25, h * 0.37).replace("\n", " ")
        vessel = ""
        voyage = ""
        m = re.search(r"([A-Z]+(?:\s+[A-Z0-9]+)*)\s+([0-9A-Z]+)$", vessel_region)
        if m:
            vessel, voyage = m.group(1).strip(), m.group(2).strip()
        else:
            vessel = vessel_region.strip()
        if re.fullmatch(r"\d+", voyage or ""):
            voyage = f"{voyage}E"

        # HANSUNG 양식 보정: HANSUNG INCHEON까지 선명, 오른쪽 3~4자리(+E)는 항차
        hv, hy = split_vessel_voyage_text(all_text)
        if hv and hv.upper().startswith("HANSUNG INCHEON"):
            vessel, voyage = hv, hy

        pol_raw = text_in_region(page, w * 0.25, h * 0.31, w * 0.48, h * 0.37).replace("\n", " ").strip()
        pod_raw = text_in_region(page, 0, h * 0.35, w * 0.28, h * 0.40).replace("\n", " ").strip()
        pol = normalize_port_code(pol_raw)
        pod = normalize_port_code(pod_raw)

        table_top = h * 0.39
        table_bottom = h * table_bottom_ratio

        mark_words = words_in_region(page, 0, table_top, w * 0.28, table_bottom, mode="start")
        mark = clean_mark(text_from_words(mark_words))

        pkg_region = text_in_region(page, w * 0.27, table_top, w * 0.39, table_top + h * 0.08, mode="start")
        pkg = safe_search(r"(\d+)\s*(?:PKGS?|PACKAGES?|CTNS?|CARTONS?|PCS)", pkg_region)
        if not pkg:
            pkg = extract_qty_from_mark(mark)

        desc_left = w * 0.37
        desc_right = w * desc_right_ratio
        desc_words = words_in_region(page, desc_left, table_top, desc_right, table_bottom, mode="start")
        description = clean_description_words(desc_words)

        weight_region = text_in_region(page, w * 0.76, table_top, w * 0.90, table_top + h * 0.08, mode="start").replace("\n", " ")
        weight = safe_search(r"(\d+(?:\.\d+)?)\s*KGS", weight_region)

        cbm_region = text_in_region(page, w * 0.88, table_top, w, table_top + h * 0.08, mode="start").replace("\n", " ")
        cbm = safe_search(r"(\d+(?:\.\d+)?)\s*CBM", cbm_region)

        if consignee_raw and "(" in consignee_raw and consignee.endswith("()"):
            consignee = consignee.replace("()", "").strip()
        if not mark:
            warnings.append("마크 미인식")
        if not description:
            warnings.append("품명 미인식")
        if not weight:
            warnings.append("중량 확인필요")
        if not cbm:
            warnings.append("CBM 확인필요")
        if not pkg:
            warnings.append("수량 확인필요")

        return {
            "비엘": bl,
            "쉬퍼": shipper,
            "쉬퍼주소": shipper_addr,
            "컨사이니": consignee,
            "컨사이니 사업자번호": consignee_no,
            "컨사이니 주소": consignee_addr,
            "선명": vessel,
            "항차": voyage,
            "출발지": pol,
            "도착지": pod,
            "품명": description,
            "마크": mark,
            "수량": pkg,
            "중량": weight,
            "CBM": cbm,
            "원본파일명": filename,
            "확인필요": ", ".join(warnings),
        }

def make_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="변환결과")
        ws = writer.book["변환결과"]
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
            cell.alignment = cell.alignment.copy(horizontal="center", vertical="center")
        widths = {
            "A": 22, "B": 30, "C": 42, "D": 24, "E": 20, "F": 45,
            "G": 20, "H": 12, "I": 18, "J": 18, "K": 55, "L": 35,
            "M": 10, "N": 12, "O": 10, "P": 34, "Q": 28,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = cell.alignment.copy(wrap_text=True, vertical="top")
        for r in range(2, ws.max_row + 1):
            ws.row_dimensions[r].height = 95
    output.seek(0)
    return output.getvalue()

st.set_page_config(
    page_title="TY LOGIS 업무 자동화 시스템",
    layout="wide",
    initial_sidebar_state="collapsed"
)

import base64

def img_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

if "login" not in st.session_state:
    st.session_state.login = False
if "user" not in st.session_state:
    st.session_state.user = ""
if "page" not in st.session_state:
    st.session_state.page = "main"

logo_html = ""
if os.path.exists("ty_logo.png"):
    logo_b64 = img_to_base64("ty_logo.png")
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" class="logo-img">'
else:
    logo_html = '<div class="logo-missing">ty_logo.png 파일이 없습니다.</div>'

st.markdown("""
<style>
:root {
    --bg: #f4f6f9; --panel: #ffffff; --text: #111827; --muted: #64748b;
    --line: #e5e7eb; --gold-light: #f3dfad; --dark: #2f2f33; --dark-hover: #1f1f23;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] {display: none;}
#MainMenu {visibility: hidden;} footer {visibility: hidden;}
.stApp { background: var(--bg); }
.block-container { max-width: 1200px; padding-top: 2.5rem; padding-bottom: 2rem; }
.badge { display:inline-block; background:var(--gold-light); color:var(--dark); padding:8px 18px; border-radius:999px; font-size:13px; font-weight:800; margin-bottom:20px; }
.title { font-size:42px; line-height:1.15; font-weight:900; color:var(--text); margin:0 0 18px 0; letter-spacing:-1.2px; }
.desc { color:var(--muted); font-size:16px; line-height:1.75; margin-bottom:30px; }
.logo-box { background:#ffffff; border-radius:14px; padding:18px; width:420px; max-width:100%; box-shadow:0 12px 32px rgba(15,23,42,0.06); border:1px solid var(--line); text-align:center; box-sizing:border-box; }
.logo-img { width:360px; max-width:100%; display:block; margin:0 auto; }
.version { margin-top:18px; background:#f8fafc; color:#94a3b8; text-align:center; font-size:12px; padding:10px; width:420px; max-width:100%; border-radius:10px; border:1px solid var(--line); box-sizing:border-box; }
.login-title { font-size:34px; font-weight:900; color:var(--text); margin-bottom:10px; }
.login-sub { color:var(--muted); font-size:15px; margin-bottom:28px; }
[data-testid="stTextInput"] label, [data-testid="stFileUploader"] label, [data-testid="stRadio"] label, [data-testid="stSlider"] label { color:var(--text); font-weight:600; }
[data-testid="stTextInput"] input { background:#f8fafc; border:1px solid #d9dee7; color:var(--text); border-radius:10px; }
div.stButton > button, div.stDownloadButton > button { background:var(--dark); color:white; border:none; border-radius:10px; height:48px; font-weight:800; }
div.stButton > button:hover, div.stDownloadButton > button:hover { background:var(--dark-hover); color:white; border:none; }
.small-text { color:#94a3b8; font-size:12px; margin-top:20px; }
.topbar { background:#ffffff; border:1px solid var(--line); border-radius:18px; padding:22px 28px; display:flex; justify-content:space-between; align-items:center; box-shadow:0 12px 32px rgba(15,23,42,0.05); margin-bottom:26px; }
.topbar-title { font-size:26px; font-weight:900; color:var(--text); }
.topbar-sub { font-size:14px; color:var(--muted); margin-top:5px; }
.dashboard-card { background:#ffffff; border:1px solid var(--line); border-radius:18px; padding:26px; box-shadow:0 12px 32px rgba(15,23,42,0.05); min-height:150px; }
.card-icon { font-size:28px; margin-bottom:12px; }
.card-title { font-size:20px; font-weight:900; color:var(--text); margin-bottom:8px; }
.card-desc { font-size:14px; color:var(--muted); line-height:1.6; }
.section-title { font-size:20px; font-weight:900; color:var(--text); margin:18px 0 14px 0; }
.content-card { background:#ffffff; border:1px solid var(--line); border-radius:18px; padding:28px; box-shadow:0 12px 32px rgba(15,23,42,0.05); margin-bottom:18px; }
.page-title { font-size:30px; font-weight:900; color:var(--text); margin-bottom:8px; }
.page-sub { color:var(--muted); font-size:15px; }
</style>
""", unsafe_allow_html=True)

def login_page():
    left, right = st.columns([1.05, 1], gap="large")
    with left:
        st.markdown(f"""
        <div class="badge">TY · KY · YST 통합 업무 포털</div>
        <div class="title">업무 자동화 시스템</div>
        <div class="desc">전자상거래 · 3PL · 씨앤에어 업무를 한 곳에서 처리합니다.<br>파일 변환, 검증, 현장 운영 자료를 빠르게 자동화합니다.</div>
        <div class="logo-box">{logo_html}</div>
        <div class="version">TY LOGIS Internal System · v22.4</div>
        """, unsafe_allow_html=True)
    with right:
        st.markdown('<div class="login-title">로그인</div><div class="login-sub">계정과 비밀번호를 입력하세요.</div>', unsafe_allow_html=True)
        user = st.text_input("사용자 계정", placeholder="예: admin")
        pw = st.text_input("비밀번호", type="password", placeholder="기본 테스트 비밀번호: 1234")
        if st.button("로그인"):
            if (user == "admin" and pw == "1234") or (user == "ty" and pw == "1234") or (user == "yst" and pw == "1234"):
                st.session_state.login = True; st.session_state.user = user; st.session_state.page = "main"; st.rerun()
            else:
                st.error("아이디 또는 비밀번호가 틀렸습니다.")
        st.markdown('<div class="small-text">테스트 계정: admin / 1234, ty / 1234, yst / 1234</div>', unsafe_allow_html=True)

def topbar():
    page_map = {
        "main": "메인 대시보드",
        "ecommerce": "전자상거래",
        "seaair": "SEA & AIR",
        "threepl": "3PL",
        "bl_convert": "3PL BL PDF 변환",
        "kyungdong": "전자상 경동리스트",
        "meni_convert": "메니변환",
    }
    page_name = page_map.get(st.session_state.page, "메인 대시보드")
    st.markdown(f"""
    <div class="topbar"><div><div class="topbar-title">TY LOGIS 업무 자동화 시스템</div><div class="topbar-sub">접속 계정: {st.session_state.user} · {page_name}</div></div><div class="badge">v22.4</div></div>
    """, unsafe_allow_html=True)


def main_page():
    topbar()

    st.markdown('<div class="section-title">부서별 업무 메뉴</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3, gap="large")

    with c1:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">🛒</div>'
            '<div class="card-title">전자상거래</div>'
            '<div class="card-desc">전자상 통관·택배·경동리스트 관련 업무를 처리합니다.</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("전자상거래 들어가기", use_container_width=True, key="go_ecom"):
            st.session_state.page = "ecommerce"
            st.rerun()

    with c2:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">🚢</div>'
            '<div class="card-title">SEA & AIR</div>'
            '<div class="card-desc">해상·항공 포워딩 관련 업무 메뉴를 구성할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("SEA & AIR 들어가기", use_container_width=True, key="go_seaair"):
            st.session_state.page = "seaair"
            st.rerun()

    with c3:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">🏭</div>'
            '<div class="card-title">3PL</div>'
            '<div class="card-desc">BL/PDF 변환, 현장 운영, 적재 관련 업무를 처리합니다.</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("3PL 들어가기", use_container_width=True, key="go_3pl"):
            st.session_state.page = "threepl"
            st.rerun()

    st.divider()

    if st.button("로그아웃"):
        st.session_state.login = False
        st.session_state.user = ""
        st.session_state.page = "main"
        st.rerun()


def ecommerce_page():
    topbar()
    if st.button("← 메인으로 돌아가기", key="ecom_back"):
        st.session_state.page = "main"
        st.rerun()

    st.markdown(
        '<div class="content-card"><div class="page-title">🛒 전자상거래</div>'
        '<div class="page-sub">전자상 통관·택배·경동리스트 업무 메뉴입니다.</div></div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">📦</div>'
            '<div class="card-title">전자상 경동리스트</div>'
            '<div class="card-desc">멀티건 송장 매칭, 동춘경동, 학익경동 자동변환을 처리합니다.</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("전자상 경동리스트 열기", use_container_width=True, key="open_kd_from_ecom"):
            st.session_state.page = "kyungdong"
            st.rerun()

    with c2:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">📍</div>'
            '<div class="card-title">주소 / 통관 검증</div>'
            '<div class="card-desc">주소 정리, 우편번호 확인, 개인통관부호 검증 메뉴를 추가할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        st.button("준비중", use_container_width=True, disabled=True, key="ecom_addr_ready")

    with c3:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">🧾</div>'
            '<div class="card-title">메니변환</div>'
            '<div class="card-desc">HDFC 메니 파일의 HS코드, 용도구분, 중량, FTA 관련 처리를 자동 변환합니다.</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("메니변환 열기", use_container_width=True, key="open_meni_from_ecom"):
            st.session_state.page = "meni_convert"
            st.rerun()


def seaair_page():
    topbar()
    if st.button("← 메인으로 돌아가기", key="seaair_back"):
        st.session_state.page = "main"
        st.rerun()

    st.markdown(
        '<div class="content-card"><div class="page-title">🚢 SEA & AIR</div>'
        '<div class="page-sub">해상·항공 포워딩 관련 업무 메뉴입니다.</div></div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">📄</div>'
            '<div class="card-title">서류 변환</div>'
            '<div class="card-desc">해상·항공 서류 자동화 메뉴를 추가할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        st.button("준비중", use_container_width=True, disabled=True, key="seaair_doc_ready")

    with c2:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">✈️</div>'
            '<div class="card-title">항공 업무</div>'
            '<div class="card-desc">항공 관련 업무 메뉴를 추가할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        st.button("준비중", use_container_width=True, disabled=True, key="seaair_air_ready")

    with c3:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">🚢</div>'
            '<div class="card-title">해상 업무</div>'
            '<div class="card-desc">해상 관련 업무 메뉴를 추가할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        st.button("준비중", use_container_width=True, disabled=True, key="seaair_sea_ready")


def threepl_page():
    topbar()
    if st.button("← 메인으로 돌아가기", key="threepl_back"):
        st.session_state.page = "main"
        st.rerun()

    st.markdown(
        '<div class="content-card"><div class="page-title">🏭 3PL</div>'
        '<div class="page-sub">3PL BL/PDF 변환 및 현장 운영 업무 메뉴입니다.</div></div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">📄</div>'
            '<div class="card-title">3PL BL / PDF 변환</div>'
            '<div class="card-desc">PDF BL 파일을 업로드하면 엑셀 변환 결과를 생성합니다.</div></div>',
            unsafe_allow_html=True,
        )
        if st.button("BL 변환 열기", use_container_width=True, key="open_bl_from_3pl"):
            st.session_state.page = "bl_convert"
            st.rerun()

    with c2:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">📦</div>'
            '<div class="card-title">팔레트 적재</div>'
            '<div class="card-desc">현장 적재 및 팔레트 계산 메뉴를 추가할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        st.button("준비중", use_container_width=True, disabled=True, key="threepl_pallet_ready")

    with c3:
        st.markdown(
            '<div class="dashboard-card"><div class="card-icon">🏷️</div>'
            '<div class="card-title">현장 운영</div>'
            '<div class="card-desc">입출고, 현장 작업자료 메뉴를 추가할 수 있습니다.</div></div>',
            unsafe_allow_html=True,
        )
        st.button("준비중", use_container_width=True, disabled=True, key="threepl_ops_ready")


def bl_convert_page():
    topbar()
    if st.button("← 메인으로 돌아가기"):
        st.session_state.page = "main"; st.rerun()
    st.markdown('<div class="content-card"><div class="page-title">📄 3PL BL PDF → Excel 변환</div><div class="page-sub">여러 개의 BL PDF 파일을 업로드하면 엑셀 파일로 변환합니다.</div></div>', unsafe_allow_html=True)
    with st.expander("추출 설정", expanded=False):
        desc_mode = st.radio("품명 오른쪽 범위", ["안전모드", "넓게 인식"], help="안전모드는 중량칸 혼입을 더 강하게 막고, 넓게 인식은 길게 넘어간 품명을 더 많이 잡습니다.")
        desc_right_ratio = 0.80 if desc_mode == "안전모드" else 0.89
        table_bottom_ratio = st.slider("MARK/품명 하단 범위", min_value=0.55, max_value=0.78, value=0.69, step=0.01, help="마크가 아래로 길게 이어지는 특수 PDF면 조금 올려주세요. 너무 올리면 하단 문구가 포함될 수 있습니다.")
    uploaded = st.file_uploader("BL PDF 파일 업로드", type=["pdf"], accept_multiple_files=True)
    if uploaded:
        rows = []
        progress = st.progress(0)
        for idx, file in enumerate(uploaded, start=1):
            try:
                rows.append(extract_one_pdf(file.getvalue(), file.name, desc_right_ratio, table_bottom_ratio))
            except Exception as e:
                rows.append({**{c: "" for c in COLUMNS}, "원본파일명": file.name, "확인필요": f"오류: {e}"})
            progress.progress(idx / len(uploaded))
        df = pd.DataFrame(rows, columns=COLUMNS)
        st.subheader("미리보기")
        st.dataframe(df, use_container_width=True, hide_index=True)
        excel_bytes = make_excel(df)
        file_name = f"BL_PDF_변환결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button("엑셀 다운로드", data=excel_bytes, file_name=file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with st.expander("반영된 변환 기능"):
            st.write("- SREJY / SILKROAD 특수 BL 라벨 기준 인식 보정")
            st.write("- HANSUNG INCHEON 선명/항차 분리 보정")
            st.write("- WOOYOUNG / INUF 화주 전용 BL 양식 라벨 기준 인식 보정")
            st.write("- 항차 패턴 보정: 숫자 4자리, 숫자4자리+E, 숫자3자리+E")
            st.write("- 컨사이니 주소, 선명/항차, 마크, 수량, 중량, CBM 인식 범위 보정")
            st.write("- 품명은 SAID TO CONTAIN 아래 내용 기준으로 추출")
    else:
        st.info("PDF를 업로드하면 자동으로 엑셀 변환 결과가 생성됩니다.")


# ==============================
# 전자상 경동리스트
# ==============================
def kd_parse_list(text: str):
    if not text or not text.strip():
        return []
    items = re.split(r"[,\s;]+", text.strip())
    out, seen = [], set()
    for x in items:
        x = x.strip()
        if x and x not in seen:
            out.append(x)
            seen.add(x)
    return out

def kd_norm_str(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.lower() == "nan":
        return ""
    return s

def kd_to_float(x):
    try:
        if pd.isna(x):
            return None
        s = str(x).replace(",", "").strip()
        if s == "" or s.lower() == "nan":
            return None
        return float(s)
    except:
        return None

def kd_set_number_like_original(value):
    s = kd_norm_str(value)
    if s == "":
        return ""
    try:
        return int(s)
    except:
        return s

def kd_replace_delivery_terms(df):
    replace_map = {
        "경착": "착택",
        "착불": "착택",
        "신용": "현택",
        "대납": "현택",
        "선불": "현택",
    }

    def convert_cell(x):
        s = kd_norm_str(x)
        return replace_map.get(s, x)

    return df.applymap(convert_cell)

def kyungdong_page():
    topbar()
    if st.button("← 메인으로 돌아가기", key="kd_back"):
        st.session_state.page = "main"
        st.rerun()

    st.markdown(
        '<div class="content-card"><div class="page-title">📦 전자상 경동리스트</div>'
        '<div class="page-sub">멀티건 송장 매칭 · 동춘경동 자동변환 · 학익경동 자동변환</div></div>',
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3 = st.tabs(["① 멀티건 송장 매칭", "② 동춘경동 자동변환", "③ 학익경동 자동변환"])

    # ---------------- TAB 1 ----------------
    with tab1:
        st.subheader("멀티건 송장 매칭 (조회불가 포함)")
        uploaded = st.file_uploader("멀티건모음 엑셀 업로드", type=["xlsx", "xls"], key="kd_multi")

        e_text = st.text_area(
            "조회할 택배송장(E) 목록 (선택)",
            height=120,
            placeholder="예)\n507144526830\n507144526841\n...",
            key="kd_e_text",
        )

        if uploaded:
            xls = pd.ExcelFile(uploaded)
            sheet = st.selectbox("시트 선택", xls.sheet_names, key="kd_sheet")
            header_mode = st.radio("첫 줄이 컬럼명(헤더)인가요?", ["네", "아니요"], horizontal=True, key="kd_header")

            if header_mode == "네":
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                st.caption(f"행 {len(df):,} / 컬럼 {len(df.columns)}")

                e_col = st.selectbox(
                    "E열(택배송장) 컬럼",
                    df.columns,
                    index=df.columns.get_loc("CJ单号") if "CJ单号" in df.columns else 0,
                    key="kd_e_col",
                )

                f_col = st.selectbox(
                    "F열(세관신고송장) 컬럼",
                    df.columns,
                    index=df.columns.get_loc("主单号") if "主单号" in df.columns else 0,
                    key="kd_f_col",
                )

            else:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str, header=None)
                df.columns = [
                    string.ascii_uppercase[i] if i < 26 else f"COL{i+1}"
                    for i in range(len(df.columns))
                ]

                st.caption(f"행 {len(df):,} / 컬럼 {len(df.columns)} (헤더 없음)")

                e_col = st.selectbox(
                    "E열 문자",
                    df.columns,
                    index=df.columns.get_loc("E") if "E" in df.columns else 0,
                    key="kd_e_col_nohead",
                )

                f_col = st.selectbox(
                    "F열 문자",
                    df.columns,
                    index=df.columns.get_loc("F") if "F" in df.columns else 0,
                    key="kd_f_col_nohead",
                )

            with st.expander("데이터 미리보기(상위 30행)"):
                st.dataframe(df.head(30), use_container_width=True)

            if st.button("✅ 매칭 생성", type="primary", use_container_width=True, key="kd_make_match"):
                query = kd_parse_list(e_text)

                base = df.copy()
                base[e_col] = base[e_col].apply(kd_norm_str)
                base[f_col] = base[f_col].apply(kd_norm_str)
                base = base[base[e_col] != ""]

                not_in_file = []

                if query:
                    all_e = set(base[e_col].tolist())
                    not_in_file = [e for e in query if e not in all_e]
                    base = base[base[e_col].isin(query)].copy()

                valid = base[base[f_col] != ""][[e_col, f_col]].copy()
                invalid = base[base[f_col] == ""][[e_col]].copy()

                out1 = valid.rename(columns={e_col: "택배송장(E)", f_col: "세관신고송장(F)"})
                out1["비고"] = "정상 매칭"

                out2 = invalid.rename(columns={e_col: "택배송장(E)"})
                out2["비고"] = "조회안되는송장(F 없음)"

                if not_in_file:
                    add = pd.DataFrame({
                        "택배송장(E)": not_in_file,
                        "비고": ["조회안되는송장(파일에 없음)"] * len(not_in_file),
                    })
                    out2 = pd.concat([out2, add], ignore_index=True)

                out1 = out1.drop_duplicates().reset_index(drop=True)
                out2 = out2.drop_duplicates().reset_index(drop=True)

                st.success(f"정상매칭 {len(out1):,}건 / 조회안되는송장 {len(out2):,}건")

                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("### 정상매칭")
                    st.dataframe(out1.head(200), use_container_width=True)
                with c2:
                    st.markdown("### 조회안되는송장")
                    st.dataframe(out2.head(200), use_container_width=True)

                bio = io.BytesIO()
                with pd.ExcelWriter(bio, engine="openpyxl") as w:
                    out1.to_excel(w, index=False, sheet_name="정상매칭")
                    out2.to_excel(w, index=False, sheet_name="조회안되는송장")
                bio.seek(0)

                st.download_button(
                    "⬇️ 엑셀 다운로드 (정상매칭/조회안되는송장)",
                    bio.getvalue(),
                    file_name="멀티건_E-F_매칭_조회불가포함.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="kd_match_download",
                )

    # ---------------- TAB 2 ----------------
    with tab2:
        st.subheader("동춘경동 자동변환 (멀티건 분할 + 중량/BOX 처리 + A열 숫자)")
        st.write("**멀티건 매칭파일(정상매칭 시트)** + **동춘경동 신규파일**을 업로드하세요.")

        match_file = st.file_uploader(
            "멀티건 매칭파일(.xlsx) — 정상매칭 시트 포함",
            type=["xlsx"],
            key="kd_dc_match",
        )

        hak_file = st.file_uploader(
            "동춘경동 신규파일(.xlsx)",
            type=["xlsx"],
            key="kd_dc_file",
        )

        if match_file and hak_file:
            match_df = pd.read_excel(match_file, sheet_name="정상매칭", dtype=str)
            hak_df = pd.read_excel(hak_file, dtype=str)

            if "세관신고송장(F)" not in match_df.columns or "택배송장(E)" not in match_df.columns:
                st.error("매칭파일의 '정상매칭' 시트 컬럼명이 예상과 달라요. (택배송장(E), 세관신고송장(F))")
                st.stop()

            required_dc_cols = ["HBL NO", "BOX 수량", "중량"]
            missing_dc_cols = [c for c in required_dc_cols if c not in hak_df.columns]
            if missing_dc_cols:
                st.error(
                    "동춘경동 신규파일 양식이 아니에요. "
                    f"필요 컬럼: {', '.join(required_dc_cols)} / 없는 컬럼: {', '.join(missing_dc_cols)}\n\n"
                    "업로드한 파일이 학익경동 양식이면 위의 '③ 학익경동 자동변환' 탭에서 처리해주세요."
                )
                st.stop()

            map_group = match_df.groupby("세관신고송장(F)")["택배송장(E)"].apply(list).to_dict()
            rows = []

            for _, r in hak_df.iterrows():
                hbl = kd_norm_str(r.get("HBL NO"))
                e_list = map_group.get(hbl)

                box = kd_to_float(r.get("BOX 수량"))
                wt = kd_to_float(r.get("중량"))

                denom = box if (box and box > 0) else (len(e_list) if e_list else None)
                per = (wt / denom) if (wt is not None and denom) else wt

                def base_row(rr):
                    rr = rr.copy()
                    if "BOX 수량" in rr.index:
                        rr.loc["BOX 수량"] = 1
                    if per is not None and "중량" in rr.index:
                        rr.loc["중량"] = round(per, 3)
                    return rr

                if e_list:
                    for e in e_list:
                        rr = base_row(r)
                        try:
                            rr["HBL NO"] = int(kd_norm_str(e))
                        except:
                            rr["HBL NO"] = kd_norm_str(e)
                        rows.append(rr)
                else:
                    rr = base_row(r)
                    rows.append(rr)

            out = pd.DataFrame(rows)
            if set(hak_df.columns).issubset(set(out.columns)):
                out = out[hak_df.columns]

            out = kd_replace_delivery_terms(out)

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as w:
                out.to_excel(w, index=False, sheet_name="변환결과")
            bio.seek(0)

            st.success(f"변환 완료: {len(out):,}행")
            st.dataframe(out.head(50), use_container_width=True)

            st.download_button(
                "⬇️ 동춘경동 변환파일 다운로드",
                bio.getvalue(),
                file_name="동춘경동_택배송장변환.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="kd_dc_download",
            )

    # ---------------- TAB 3 ----------------
    with tab3:
        st.subheader("학익경동 자동변환 (멀티건 분할 + C/T·W/T 처리)")
        st.write("**멀티건 매칭파일(정상매칭 시트)** + **학익경동 신규파일**을 업로드하세요.")
        st.caption("학익경동 양식 기준: HBL NO(또는 송장번호) / C/T / W/T 컬럼을 사용합니다.")

        hakik_match_file = st.file_uploader(
            "멀티건 매칭파일(.xlsx) — 정상매칭 시트 포함",
            type=["xlsx"],
            key="kd_hakik_match",
        )

        hakik_file = st.file_uploader(
            "학익경동 신규파일(.xlsx)",
            type=["xlsx"],
            key="kd_hakik_file",
        )

        if hakik_match_file and hakik_file:
            match_df = pd.read_excel(hakik_match_file, sheet_name="정상매칭", dtype=str)
            hakik_df = pd.read_excel(hakik_file, dtype=str)

            required_match_cols = ["세관신고송장(F)", "택배송장(E)"]
            missing_match_cols = [c for c in required_match_cols if c not in match_df.columns]

            if missing_match_cols:
                st.error("매칭파일의 '정상매칭' 시트 컬럼명이 예상과 달라요. (택배송장(E), 세관신고송장(F))")
                st.stop()

            if "HBL NO" in hakik_df.columns:
                waybill_col = "HBL NO"
            elif "송장번호" in hakik_df.columns:
                waybill_col = "송장번호"
            else:
                st.error("학익경동 신규파일에 필요한 컬럼이 없어요: HBL NO 또는 송장번호")
                st.stop()

            required_hakik_cols = [waybill_col, "C/T", "W/T"]
            missing_hakik_cols = [c for c in required_hakik_cols if c not in hakik_df.columns]

            if missing_hakik_cols:
                st.error(f"학익경동 신규파일에 필요한 컬럼이 없어요: {', '.join(missing_hakik_cols)}")
                st.stop()

            match_df["세관신고송장(F)"] = match_df["세관신고송장(F)"].apply(kd_norm_str)
            match_df["택배송장(E)"] = match_df["택배송장(E)"].apply(kd_norm_str)

            map_group = (
                match_df[match_df["세관신고송장(F)"] != ""]
                .groupby("세관신고송장(F)")["택배송장(E)"]
                .apply(list)
                .to_dict()
            )

            rows = []
            unmatched_rows = []

            for _, r in hakik_df.iterrows():
                waybill = kd_norm_str(r.get(waybill_col))
                e_list = map_group.get(waybill)

                ct = kd_to_float(r.get("C/T"))
                wt = kd_to_float(r.get("W/T"))

                if e_list:
                    repeat_count = len(e_list)
                elif ct and ct > 0:
                    repeat_count = int(ct)
                else:
                    repeat_count = 1

                weight_divisor = ct if (ct and ct > 0) else repeat_count
                per_wt = (wt / weight_divisor) if (wt is not None and weight_divisor) else wt

                def make_row(rr):
                    rr = rr.copy()
                    if "C/T" in rr.index:
                        rr.loc["C/T"] = 1
                    if per_wt is not None and "W/T" in rr.index:
                        rr.loc["W/T"] = round(per_wt, 3)
                    return rr

                if e_list:
                    for e in e_list:
                        rr = make_row(r)
                        rr[waybill_col] = kd_set_number_like_original(e)
                        rows.append(rr)
                else:
                    for _ in range(repeat_count):
                        rr = make_row(r)
                        rows.append(rr)

                    if ct and ct > 1:
                        unmatched_rows.append({
                            waybill_col: waybill,
                            "C/T": ct,
                            "W/T": wt,
                            "비고": "매칭 송장 없이 C/T 기준으로 자동 분할됨",
                        })

            out = pd.DataFrame(rows)
            if set(hakik_df.columns).issubset(set(out.columns)):
                out = out[hakik_df.columns]

            out = kd_replace_delivery_terms(out)

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as w:
                out.to_excel(w, index=False, sheet_name="변환결과")
                if unmatched_rows:
                    pd.DataFrame(unmatched_rows).to_excel(w, index=False, sheet_name="확인필요")
            bio.seek(0)

            st.success(f"변환 완료: {len(out):,}행")

            if unmatched_rows:
                st.warning(f"확인필요 {len(unmatched_rows):,}건이 있습니다. 다운로드 파일의 '확인필요' 시트를 확인해주세요.")

            st.dataframe(out.head(50), use_container_width=True)

            st.download_button(
                "⬇️ 학익경동 변환파일 다운로드",
                bio.getvalue(),
                file_name="학익경동_택배송장변환.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="kd_hakik_download",
            )



# ==============================
# 메니변환
# ==============================
MENI_KEYWORDS_V = [
    "mini PC", "Vibrator", "smart watch", "Earphones",
    "tablet android", "tablet pc", "Speakers", "smartphone",
    "lenovo xiaoxin pad", "IPAD TABLET", "TABLET", "DRAWING BOARD",
    "BLUETOOTH SPEAKER", "BLUETOOTH EARPHONE CLIP",
    "BLUETOOTH EARBUDS", "BLUETOOTH",
]

MENI_WIRELESS_TERMS = [
    "Wireless Earphones",
    "Wireless Headphones",
    "Wireless Earbuds",
    "Wireless Bluetooth Earphones",
    "Wireless Bluetooth Headphones",
]

MENI_WEIGHT_CYCLE = [1.5, 1.6, 1.7, 1.8, 1.9]
MENI_STEP_WEIGHT = 0.1
MENI_HS_EXCEPT_FOR_V3 = {"900290", "900410", "902920"}

MENI_FTA_HS_CODES = {
    "330430", "420292", "630622", "630629", "640219", "640299",
    "640419", "732112", "846729", "847130", "847141", "847150",
    "847160", "847180", "850440", "850760", "850811", "851511",
    "851671", "851762", "851810", "851830", "851840", "852351",
    "852499", "852589", "852691", "852852", "852862", "852869",
    "854231", "854370", "870829", "870894", "870899", "871200",
    "871496", "871499", "900290", "900410", "902920", "910212",
    "910591", "940169", "950450", "950490", "950691",
}

def meni_find_column_name(columns, keyword, startswith=False):
    for c in columns:
        s = str(c)
        if startswith:
            if s.startswith(keyword):
                return c
        else:
            if keyword in s:
                return c
    raise ValueError(f"'{keyword}' 컬럼명을 찾을 수 없습니다.")

def meni_distribute_to_target(df, col_af, target_total):
    w_series = pd.to_numeric(df[col_af], errors="coerce")
    current_total = float(w_series.sum())
    remaining = float(target_total) - current_total

    if remaining <= 0:
        return df, current_total, 0.0

    candidates = df.index[pd.to_numeric(df[col_af], errors="coerce") >= 2].tolist()
    if not candidates:
        return df, current_total, 0.0

    distributed = 0.0
    i = 0
    n = len(candidates)
    max_loops = 2_000_000

    while remaining >= MENI_STEP_WEIGHT - 1e-12 and max_loops > 0:
        idx = candidates[i % n]
        cur = float(w_series.loc[idx])
        new_val = cur + MENI_STEP_WEIGHT

        if abs(new_val - 30.0) < 1e-12:
            i += 1
            max_loops -= 1
            continue

        w_series.loc[idx] = new_val
        df.at[idx, col_af] = new_val

        distributed += MENI_STEP_WEIGHT
        remaining -= MENI_STEP_WEIGHT
        i += 1
        max_loops -= 1

    if remaining > 0:
        w_series = pd.to_numeric(df[col_af], errors="coerce")
        max_idx = w_series.idxmax()
        cur = float(w_series.loc[max_idx])
        new_val = cur + remaining

        if abs(new_val - 30.0) >= 1e-12:
            df.at[max_idx, col_af] = new_val
            distributed += remaining
            w_series.loc[max_idx] = new_val
            remaining = 0.0

    new_total = float(w_series.sum())
    return df, new_total, distributed

def meni_process_excel_to_bytes(uploaded_file, target_total=None):
    df = pd.read_excel(uploaded_file)

    col_hs    = meni_find_column_name(df.columns, "허용품목코드")
    col_zip   = meni_find_column_name(df.columns, "ZIP CODE")
    col_v     = meni_find_column_name(df.columns, "용도구분")
    col_desc1 = meni_find_column_name(df.columns, "1.DESCRIPTION", startswith=True)
    col_desc2 = meni_find_column_name(df.columns, "2.DESCRIPTION", startswith=True)
    col_af    = meni_find_column_name(df.columns, "Total W/T")
    col_hawb  = meni_find_column_name(df.columns, "HAWB NO")
    col_tel   = meni_find_column_name(df.columns, "C/TEL")
    col_total = meni_find_column_name(df.columns, "总金额")

    w_orig = pd.to_numeric(df[col_af], errors="coerce")
    count_le2_orig = int(((w_orig <= 2) & w_orig.notna()).sum())

    def convert_hs_row(row):
        val_hs = row[col_hs]
        v_val = row[col_v]
        if pd.isna(val_hs):
            return val_hs
        s = str(val_hs).strip()

        if str(v_val).strip() == "3" and s in MENI_HS_EXCEPT_FOR_V3:
            return s

        if s.startswith(("1", "2", "30", "90")):
            return "960719"
        return s

    df[col_hs] = df.apply(convert_hs_row, axis=1)

    def fix_zip(z):
        if pd.isna(z):
            return z
        s = str(z).strip()
        if len(s) == 4:
            return "0" + s
        return s

    df[col_zip] = df[col_zip].apply(fix_zip)

    v_before_str = df[col_v].astype(str).str.strip()
    wireless_mask = (v_before_str == "1") & df[col_desc1].astype(str).apply(
        lambda x: any(term.upper() in str(x).upper() for term in MENI_WIRELESS_TERMS)
    )
    df.loc[wireless_mask, col_v] = 3
    rows_v_blue = df.index[wireless_mask].tolist()
    wireless_changed_cnt = int(wireless_mask.sum())
    wireless_ratio_all = (wireless_changed_cnt / len(df) * 100) if len(df) else 0.0

    def match_v(desc, v):
        if pd.isna(desc) or pd.isna(v):
            return False
        if str(v).strip() != "1":
            return False
        t = str(desc).upper()
        return any(kw.upper() in t for kw in MENI_KEYWORDS_V)

    mask_v = df.apply(lambda r: match_v(r[col_desc1], r[col_v]), axis=1)
    df.loc[mask_v, col_v] = 3
    rows_v_red = df.index[mask_v].tolist()

    w = pd.to_numeric(df[col_af], errors="coerce")
    mask_range = (w >= 2) & (w <= 5)
    bp_empty = df[col_desc2].isna() | (df[col_desc2].astype(str).str.strip() == "")
    bh_no_elec = ~df[col_desc1].astype(str).str.upper().str.contains("ELECTRIC", na=False)
    mask_target = mask_range & bp_empty & bh_no_elec
    t_idx = df.index[mask_target].tolist()

    for i, idx in enumerate(t_idx):
        df.at[idx, col_af] = MENI_WEIGHT_CYCLE[i % len(MENI_WEIGHT_CYCLE)]

    distributed_total = None
    new_total = None
    if target_total is not None:
        df, new_total, distributed_total = meni_distribute_to_target(df, col_af, target_total)

    w_after = pd.to_numeric(df[col_af], errors="coerce")
    count_le2_after = int(((w_after <= 2) & w_after.notna()).sum())

    tel = df[col_tel].astype(str).str.strip()
    v_str = df[col_v].astype(str).str.strip()
    amt = pd.to_numeric(df[col_total], errors="coerce")

    mask_v1 = (v_str == "1") & tel.notna() & (tel != "")
    df_v1 = pd.DataFrame({"TEL": tel.where(mask_v1), "AMT": amt.where(mask_v1)})
    tel_sum = df_v1.groupby("TEL", dropna=True)["AMT"].sum()
    bad_tels = set(tel_sum[tel_sum >= 150].index.tolist())

    mask_phone_rule = mask_v1 & tel.isin(bad_tels)
    rows_v_orange = df.index[mask_phone_rule].tolist()
    df.loc[mask_phone_rule, col_v] = 3
    hawb_list = df.loc[mask_phone_rule, col_hawb].astype(str).tolist()

    v_after_str = df[col_v].astype(str).str.strip()
    hs_str = df[col_hs].astype(str).str.strip()
    amt_after = pd.to_numeric(df[col_total], errors="coerce")

    mask_fta = (v_after_str == "3") & (amt_after >= 150) & hs_str.isin(MENI_FTA_HS_CODES)
    fta_hawb_list = df.loc[mask_fta, col_hawb].astype(str).tolist()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="처리결과")
        wb = writer.book
        ws = writer.sheets["처리결과"]

        header = next(ws.iter_rows(min_row=1, max_row=1))
        col_idx = {cell.value: cell.column for cell in header}
        v_idx = col_idx[col_v]
        af_idx = col_idx[col_af]

        blue   = PatternFill(start_color="FF0070C0", end_color="FF0070C0", fill_type="solid")
        red    = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        green  = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")
        orange = PatternFill(start_color="FFFF9900", end_color="FFFF9900", fill_type="solid")

        rows_v_blue_excel   = {i + 2 for i in rows_v_blue}
        rows_v_red_excel    = {i + 2 for i in rows_v_red}
        rows_af_excel       = {i + 2 for i in t_idx}
        rows_v_orange_excel = {i + 2 for i in rows_v_orange}

        for r in range(2, ws.max_row + 1):
            if r in rows_v_blue_excel:
                ws.cell(row=r, column=v_idx).fill = blue
            if r in rows_v_red_excel:
                ws.cell(row=r, column=v_idx).fill = red
            if r in rows_af_excel:
                ws.cell(row=r, column=af_idx).fill = green
            if r in rows_v_orange_excel:
                ws.cell(row=r, column=v_idx).fill = orange

        memo = wb.create_sheet("메모")
        memo["A1"] = "AF ≤ 2 (원본)"
        memo["B1"] = count_le2_orig

        memo["A2"] = "목표 총중량(kg)"
        memo["B2"] = "" if target_total is None else float(target_total)

        memo["A3"] = "분배 후 총중량(kg)"
        memo["B3"] = "" if new_total is None else float(new_total)

        memo["A4"] = "2차 분배량 합계(kg)"
        memo["B4"] = "" if distributed_total is None else float(distributed_total)

        memo["A5"] = "AF ≤ 2 (분배 후)"
        memo["B5"] = count_le2_after

        memo["A6"] = "WIRELESS 변경 건수(V=1→3)"
        memo["B6"] = wireless_changed_cnt

        memo["A7"] = "전화번호 중복 + 총금액합계≥150 행 수"
        memo["B7"] = len(rows_v_orange)

        memo["A8"] = "해당 전화번호 수"
        memo["B8"] = len(bad_tels)

        memo["A9"] = "WIRELESS 전체 대비 변경 비율(%)"
        memo["B9"] = round(wireless_ratio_all, 2)

        memo["A10"] = "해당 HAWB NO 리스트"
        row = 11
        for h in hawb_list:
            memo[f"A{row}"] = h
            row += 1

        row += 1
        memo[f"A{row}"] = "FTA적용건 HAWB 리스트"
        memo[f"B{row}"] = len(fta_hawb_list)
        row += 1
        for h in fta_hawb_list:
            memo[f"A{row}"] = h
            row += 1

    output.seek(0)

    summary = {
        "원본 AF≤2": count_le2_orig,
        "분배 후 AF≤2": count_le2_after,
        "WIRELESS 변경": wireless_changed_cnt,
        "키워드 V변경": len(rows_v_red),
        "중량 1차 재분배": len(t_idx),
        "전화번호/총금액 변경": len(rows_v_orange),
        "FTA 적용건": len(fta_hawb_list),
    }

    return output.getvalue(), summary

def meni_convert_page():
    topbar()
    if st.button("← 전자상거래로 돌아가기", key="meni_back"):
        st.session_state.page = "ecommerce"
        st.rerun()

    st.markdown(
        '<div class="content-card"><div class="page-title">🧾 메니변환</div>'
        '<div class="page-sub">HDFC 메니 파일 자동 처리 · HS코드/용도구분/중량/FTA/메모시트 생성</div></div>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader("메니 엑셀 파일 업로드", type=["xlsx", "xls"], key="meni_file")
    target_text = st.text_input("목표 총중량(kg) (선택)", placeholder="없으면 비워두세요", key="meni_target")

    target_total = None
    if target_text.strip():
        try:
            target_total = float(target_text.replace(",", "").strip())
        except Exception:
            st.warning("목표 총중량은 숫자로 입력해주세요. 숫자가 아니면 목표 중량 없이 처리됩니다.")
            target_total = None

    if uploaded:
        if st.button("✅ 메니변환 실행", type="primary", use_container_width=True, key="meni_run"):
            try:
                result_bytes, summary = meni_process_excel_to_bytes(uploaded, target_total=target_total)
                st.success("메니변환 완료")

                st.write("처리 요약")
                st.dataframe(pd.DataFrame([summary]), use_container_width=True, hide_index=True)

                st.download_button(
                    "⬇️ 메니변환 결과 다운로드",
                    result_bytes,
                    file_name="메니변환_중량조정_최종.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="meni_download",
                )
            except Exception as e:
                st.error(f"처리 중 오류가 발생했습니다: {e}")
    else:
        st.info("엑셀 파일을 업로드하면 메니변환을 실행할 수 있습니다.")


if not st.session_state.login:
    login_page()
else:
    if st.session_state.page == "bl_convert":
        bl_convert_page()
    elif st.session_state.page == "kyungdong":
        kyungdong_page()
    elif st.session_state.page == "meni_convert":
        meni_convert_page()
    elif st.session_state.page == "ecommerce":
        ecommerce_page()
    elif st.session_state.page == "seaair":
        seaair_page()
    elif st.session_state.page == "threepl":
        threepl_page()
    else:
        main_page()
