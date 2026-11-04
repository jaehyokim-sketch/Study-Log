#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
학습일지 파일명 일괄 표준화 도구 (Study-Log File Renamer)
규정 형식: [차수]_[학생이름]_[학번]_[주제(경맥명)].확장자
"""

import os
import sys
import re
import argparse
import unicodedata

try:
    import fitz
except ImportError:
    pass

sys.stdout.reconfigure(encoding='utf-8')


def get_file_info(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    text = ""
    try:
        if ext == '.pdf':
            doc = fitz.open(filepath)
            for p in doc[:3]:
                text += p.get_text() + "\n"
    except Exception:
        pass
    return text


def rename_logs(folder_path, default_order="6차", default_topic="독맥", apply=False):
    files = sorted(os.listdir(folder_path))
    plan = []
    target_names = {}

    for f in files:
        if not f.endswith(('.pdf', '.docx', '.hwpx')):
            continue
        filepath = os.path.join(folder_path, f)
        f_norm = unicodedata.normalize('NFC', f)
        ext = os.path.splitext(f)[1]

        # Extract info
        text = get_file_info(filepath)

        # 1. Student ID
        m_sids = re.findall(r'20\d{6,7}', text)
        fn_sids = re.findall(r'20\d{6,7}', f_norm)
        valid_sids = [s for s in m_sids if s != '20241410']
        if valid_sids:
            sid = valid_sids[0]
        elif m_sids:
            sid = m_sids[0]
        elif fn_sids:
            sid = fn_sids[0]
        else:
            sid = "학번미상"

        # Special typo fix
        if sid == '202411448':
            sid = '20241448'

        # 2. Student Name
        m_name = re.search(r'(?:성\s*명|이\s*름)[\s:：]+([가-힣]{2,4})', text)
        if not m_name:
            m_name = re.search(r'(?:성\s*명|이\s*름)\s*\n\s*([가-힣]{2,4})', text)
        if not m_name and sid != "학번미상":
            m_name = re.search(rf'{sid}\s*[/_,]?\s*([가-힣]{{2,4}})', text)
        if not m_name and sid != "학번미상":
            m_name = re.search(rf'([가-힣]{{2,4}})\s*[/_,]?\s*{sid}', text)

        if m_name:
            name = m_name.group(1).replace(' ', '')
        else:
            clean_fn = f_norm
            for kw in ['01-09주', '06차', '06주', '07주', '08-09주', '08차', '09주차', '09주', '09차', '10주차', '6차', '6주차', '9주차', '9주', '9차', '학습일지', '경혈학2', '경혈학', 'Google Docs', '독맥', '소속 경혈', '개괄', '경혈']:
                clean_fn = clean_fn.replace(kw, '')
            names = re.findall(r'[가-힣]{2,4}', clean_fn)
            name = names[-1] if names else "이름미상"

        # 3. Topic
        if '족태음비경' in text[:1000] or '족태음비경' in f_norm:
            topic = "족태음비경"
        else:
            topic = default_topic

        base_name = f"{default_order}_{name}_{sid}_{topic}"
        cand = f"{base_name}{ext}"

        if cand in target_names:
            cnt = target_names[cand]
            target_names[cand] += 1
            new_name = f"{base_name} ({cnt}){ext}"
        else:
            target_names[cand] = 1
            new_name = cand

        plan.append((filepath, os.path.join(folder_path, new_name), f, new_name))

    print(f"=== 파일명 변경 계획 (총 {len(plan)}개) ===")
    for old_path, new_path, old_name, new_name in plan:
        print(f"{old_name}  ➔  {new_name}")

    if apply:
        print("\n[일괄 변경 실행 중...]")
        success = 0
        for old_path, new_path, old_name, new_name in plan:
            if old_path == new_path:
                success += 1
                continue
            try:
                os.rename(old_path, new_path)
                success += 1
            except Exception as e:
                print(f"Error renaming {old_name}: {e}")
        print(f"완료: {success}/{len(plan)} 파일 변경 성공.")
    else:
        print("\n미리보기 완료. 실제로 변경하려면 --apply 옵션을 추가하여 실행하세요.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="학습일지 파일명 일괄 변경 도구")
    parser.add_argument("--target", required=True, help="대상 폴더 경로")
    parser.add_argument("--order", default="6차", help="차수 표기 (기본값: 6차)")
    parser.add_argument("--topic", default="독맥", help="경맥/주제명 (기본값: 독맥)")
    parser.add_argument("--apply", action="store_true", help="실제 파일명 변경 실행")
    args = parser.parse_args()

    rename_logs(args.target, args.order, args.topic, args.apply)

