#!/usr/bin/env python3
"""
풀필먼트 송장 → 사방넷 ord_no 매칭 (shmaOrdNo 단일 키).
=======================================================
2026-04-28 v0.2.0 — 사용자 결정 룰 반영.

원칙:
  ★ 받는분 이름은 매칭 키로 사용하지 않는다. ★
  쇼핑몰주문번호(shmaOrdNo)만이 유일한 정확 매칭 키.
  동성동명(이더 김채은 vs 뉴트리 김채은) 사고 방지.

매칭 로직:
  1. 풀필먼트 송장의 쇼핑몰주문번호 ← shmaOrdNo
  2. 사방넷 ord 의 shmaOrdNo
  3. 두 키가 일치하면 매칭 (1:N 합배송 자동 처리)

합배송 처리:
  같은 shmaOrdNo에 사방넷 ord_no가 여러 개 있으면 (합배송) → 모두 매칭
  같은 shmaOrdNo에 풀필먼트 송장이 여러 개 있으면 (분할 출고) → 사방넷 ord_no에 1:1 매칭

미매칭 처리:
  - 풀필먼트 송장에 shmaOrdNo가 없거나 사방넷에 없는 경우 → unmatched 보고
  - 사용자가 직접 풀필먼트·쇼핑몰에서 확인하여 처리

Usage:
  python3 match_waybills.py \\
    --fulfillment ether_with_wb.json nutri_with_wb.json \\
    --sabang sabang_orders.json \\
    --output matched.json
"""

import argparse
import json
import sys
from collections import defaultdict


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def normalize_shma(s):
    """쇼핑몰주문번호 정규화 (공백/문자열 변환)."""
    if s is None:
        return ""
    return str(s).strip()


def build_sabang_index_by_shma(sabang):
    """사방넷 주문을 shmaOrdNo로 인덱싱. 합배송이면 list."""
    by_shma = defaultdict(list)
    for o in sabang:
        if not isinstance(o, dict):
            continue
        shma = normalize_shma(o.get("shmaOrdNo"))
        if shma:
            by_shma[shma].append(o)
    return by_shma


def build_ff_index_by_shma(ff_rows):
    """풀필먼트 송장을 shmaOrdNo로 인덱싱."""
    by_shma = defaultdict(list)
    for r in ff_rows:
        if not isinstance(r, dict):
            continue
        shma = normalize_shma(r.get("쇼핑몰ord") or r.get("쇼핑몰주문번호") or r.get("shmaOrdNo"))
        if shma:
            by_shma[shma].append(r)
    return by_shma


def match_by_shma(ff_by_shma, sabang_by_shma):
    """
    shmaOrdNo 단일 키 매칭.
    반환: matched_rows, unmatched_ff, unmatched_sabang
    """
    matched = []
    unmatched_ff = []
    seen_shma = set()

    for shma, ff_rows in ff_by_shma.items():
        sb_rows = sabang_by_shma.get(shma, [])
        if not sb_rows:
            for r in ff_rows:
                unmatched_ff.append({
                    "shmaOrdNo": shma,
                    "wybl": str(r.get("운송장번호") or r.get("wybl") or ""),
                    "recv_label": r.get("수취인") or r.get("받는분") or "",
                    "ff_account": r.get("계정") or "",
                    "reason": "사방넷에 shmaOrdNo 없음 (사방넷 미수집 가능성)",
                })
            continue
        seen_shma.add(shma)

        # 합배송 / 분할 매칭
        # 풀필 송장 N개 vs 사방넷 ord M개
        # 단순 케이스: M개 ord 모두에 동일 풀필 송장(들) 매핑
        # 분할 케이스: 풀필 송장이 여러개면 ord와 1:1 매핑 (순서 휴리스틱)
        if len(ff_rows) == 1:
            wybl = str(ff_rows[0].get("운송장번호") or ff_rows[0].get("wybl") or "")
            for sb in sb_rows:
                matched.append({
                    "shmaOrdNo": shma,
                    "sabang_ord": str(sb.get("ordNo") or sb.get("ord") or ""),
                    "wybl": wybl,
                    "ff_account": ff_rows[0].get("계정") or "",
                    "match_type": "single_wybl_to_N_ord" if len(sb_rows) > 1 else "1to1",
                })
        else:
            # 분할 케이스 — 우선 1:1 짝짓기 (앞에서부터)
            for i, sb in enumerate(sb_rows):
                wb = ff_rows[i] if i < len(ff_rows) else ff_rows[-1]
                wybl = str(wb.get("운송장번호") or wb.get("wybl") or "")
                matched.append({
                    "shmaOrdNo": shma,
                    "sabang_ord": str(sb.get("ordNo") or sb.get("ord") or ""),
                    "wybl": wybl,
                    "ff_account": wb.get("계정") or "",
                    "match_type": "split_paired",
                })
            # 풀필 송장이 사방넷 ord 보다 많으면 잔여 송장 별도 보고
            if len(ff_rows) > len(sb_rows):
                for j in range(len(sb_rows), len(ff_rows)):
                    extra = ff_rows[j]
                    unmatched_ff.append({
                        "shmaOrdNo": shma,
                        "wybl": str(extra.get("운송장번호") or extra.get("wybl") or ""),
                        "recv_label": extra.get("수취인") or "",
                        "ff_account": extra.get("계정") or "",
                        "reason": "분할 출고 잔여 송장 (사방넷 ord_no 부족)",
                    })

    # 사방넷에는 있으나 풀필먼트에 없는 shmaOrdNo (정상: 빈박스·미출고)
    unmatched_sb = []
    for shma, sb_rows in sabang_by_shma.items():
        if shma not in seen_shma:
            for sb in sb_rows:
                unmatched_sb.append({
                    "shmaOrdNo": shma,
                    "sabang_ord": str(sb.get("ordNo") or sb.get("ord") or ""),
                    "ord_status": sb.get("ordStsCd") or sb.get("sts") or "",
                })

    return matched, unmatched_ff, unmatched_sb


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fulfillment", nargs="+", required=True,
                    help="풀필먼트 송장 JSON 파일들 (이더/뉴트리). 각 row에 쇼핑몰주문번호 필수.")
    ap.add_argument("--sabang", required=True,
                    help="사방넷 주문 JSON 파일. 각 row에 shmaOrdNo 필수.")
    ap.add_argument("--output", required=True, help="매칭 결과 JSON 출력")
    args = ap.parse_args()

    # 1. 사방넷 적재
    sabang = load_json(args.sabang)
    if isinstance(sabang, dict):
        sabang = sabang.get("orders") or sabang.get("data") or []
    sabang_by_shma = build_sabang_index_by_shma(sabang)

    # 2. 풀필먼트 적재 (여러 파일 합침)
    ff_rows = []
    for ff_path in args.fulfillment:
        ff = load_json(ff_path)
        if isinstance(ff, list):
            ff_rows.extend(ff)
        elif isinstance(ff, dict):
            for v in ff.values():
                if isinstance(v, list):
                    ff_rows.extend(v)
    ff_by_shma = build_ff_index_by_shma(ff_rows)

    # 3. 매칭
    matched, unmatched_ff, unmatched_sb = match_by_shma(ff_by_shma, sabang_by_shma)

    out = {
        "summary": {
            "total_ff_rows": len(ff_rows),
            "unique_ff_shmaOrdNo": len(ff_by_shma),
            "total_sabang_rows": len(sabang),
            "unique_sabang_shmaOrdNo": len(sabang_by_shma),
            "matched": len(matched),
            "unmatched_ff": len(unmatched_ff),
            "unmatched_sabang": len(unmatched_sb),
        },
        "matched": matched,
        "unmatched_ff": unmatched_ff,
        "unmatched_sabang": unmatched_sb,
    }
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"[송장 매칭 완료 — shmaOrdNo 단일 키]")
    print(f"  매칭: {len(matched)}건")
    print(f"  풀필 측 미매칭: {len(unmatched_ff)}건 (사방넷에 shmaOrdNo 없음 = 사방넷 미수집)")
    print(f"  사방넷 측 미매칭: {len(unmatched_sb)}건 (풀필 송장 없음 = 빈박스/미출고)")
    if unmatched_ff:
        print("\n⚠️ 풀필 측 미매칭 — 사방넷 미수집 의심:")
        for u in unmatched_ff[:10]:
            print(f"  shma={u['shmaOrdNo']} wybl={u['wybl']} recv={u['recv_label']} acct={u['ff_account']}")
        if len(unmatched_ff) > 10:
            print(f"  ... 외 {len(unmatched_ff) - 10}건")
        print("→ 사방넷 관리자에서 해당 쇼핑몰주문번호 확인 + 수동 주문수집 트리거")


if __name__ == "__main__":
    main()
