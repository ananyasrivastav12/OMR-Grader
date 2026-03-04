from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import cv2
import numpy as np
import pandas as pd


@dataclass(frozen=True)
class Template:
    # Reference image size used when defining absolute ROIs below.
    REF_W: int = 1654
    REF_H: int = 2256

    # Absolute ROIs measured on REF_W x REF_H. They are scaled per-image.
    NAME_ROI: Tuple[int, int, int, int] = (80, 290, 1030, 1035)
    ANSWER_ROI: Tuple[int, int, int, int] = (60, 1120, 1595, 2015)

    # Name bubble grid (A-Z x fixed columns)
    NAME_ROWS: int = 26
    NAME_COLS: int = 20
    NAME_BLANK_THRESHOLD: float = 0.07
    NAME_DELTA_THRESHOLD: float = 0.03
    NAME_DISK_RADIUS_FRAC: float = 0.35

    # Answer layout
    BLOCKS: int = 5
    QUESTIONS_PER_BLOCK: int = 20
    OPTIONS: str = "ABCD"

    # Hough params for filled answer bubbles
    H_DP: float = 1.2
    H_MIN_DIST: int = 18
    H_PARAM1: int = 120
    H_PARAM2: int = 16
    H_MIN_RADIUS: int = 7
    H_MAX_RADIUS: int = 16

    # Scoring
    SCORE_CORRECT: float = 1.0
    SCORE_WRONG: float = -0.25


T = Template()

ROOT = Path(__file__).resolve().parent
BMP_DIR = ROOT / "omr_scanned"
KEY_DIR = ROOT / "data"
OUT_DIR = ROOT / "out"
DEBUG_DIR = ROOT / "debug_out"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def read_gray(path: Path) -> np.ndarray:
    img = cv2.imread(str(path), cv2.IMREAD_GRAYSCALE)
    if img is None:
        raise ValueError(f"Failed to read image: {path}")
    return img


def scale_roi(roi: Tuple[int, int, int, int], width: int, height: int) -> Tuple[int, int, int, int]:
    x1, y1, x2, y2 = roi
    sx = width / T.REF_W
    sy = height / T.REF_H
    return (
        int(round(x1 * sx)),
        int(round(y1 * sy)),
        int(round(x2 * sx)),
        int(round(y2 * sy)),
    )


def crop(img: np.ndarray, roi: Tuple[int, int, int, int]) -> np.ndarray:
    x1, y1, x2, y2 = roi
    return img[y1:y2, x1:x2]


def pick_bmps(folder: Path) -> List[Path]:
    bmps = sorted(folder.glob("*.bmp"))
    if not bmps:
        raise FileNotFoundError(f"No .bmp files found in {folder}")
    return bmps


def pick_first_xlsx(folder: Path) -> Path:
    xs = sorted(folder.glob("*.xlsx"))
    if not xs:
        raise FileNotFoundError(f"No .xlsx files found in {folder}")
    return xs[0]


def load_answer_key_xlsx(path: Path) -> Dict[int, str]:
    df = pd.read_excel(path, header=None)
    key: Dict[int, str] = {}

    for r in range(df.shape[0]):
        for c in range(df.shape[1] - 1):
            q = df.iat[r, c]
            a = df.iat[r, c + 1]
            if pd.isna(q) or pd.isna(a):
                continue

            try:
                qn = int(q)
            except (TypeError, ValueError):
                continue

            ans = str(a).strip().upper()
            if 1 <= qn <= 500 and ans in {"A", "B", "C", "D"}:
                key[qn] = ans

    if not key:
        raise ValueError(f"Could not parse any answers from {path}")

    return key


def save_answer_key_debug(key: Dict[int, str], xlsx_path: Path) -> None:
    ensure_dir(OUT_DIR)
    rows = [{"q": q, "answer": key[q]} for q in sorted(key)]
    pd.DataFrame(rows).to_csv(OUT_DIR / "answer_key_parsed.csv", index=False)

    key_dump = {
        "source_xlsx": xlsx_path.name,
        "count": len(key),
        "first_20": {str(q): key[q] for q in sorted(key)[:20]},
        "all": {str(q): key[q] for q in sorted(key)},
    }
    with (OUT_DIR / "answer_key_parsed.json").open("w", encoding="utf-8") as f:
        json.dump(key_dump, f, indent=2)


def kmeans_1d(values: Sequence[float], k: int) -> Tuple[np.ndarray, np.ndarray]:
    v = np.asarray(values, dtype=np.float32).reshape(-1, 1)
    crit = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 100, 1e-3)
    _, labels, centers = cv2.kmeans(v, k, None, crit, 20, cv2.KMEANS_PP_CENTERS)

    centers_flat = centers.flatten()
    order = np.argsort(centers_flat)
    sorted_centers = centers_flat[order]

    remap = {int(old): int(new) for new, old in enumerate(order)}
    sorted_labels = np.array([remap[int(lbl)] for lbl in labels.flatten()], dtype=np.int32)
    return sorted_centers, sorted_labels


def detect_answer_circles(answer_patch: np.ndarray) -> Optional[np.ndarray]:
    blur = cv2.GaussianBlur(answer_patch, (5, 5), 0)
    circles = cv2.HoughCircles(
        blur,
        cv2.HOUGH_GRADIENT,
        dp=T.H_DP,
        minDist=T.H_MIN_DIST,
        param1=T.H_PARAM1,
        param2=T.H_PARAM2,
        minRadius=T.H_MIN_RADIUS,
        maxRadius=T.H_MAX_RADIUS,
    )
    if circles is None:
        return None
    return np.round(circles[0]).astype(np.int32)


def decode_answers_from_circles(circles: np.ndarray) -> Tuple[Dict[int, str], Dict]:
    if circles.shape[0] < 70:
        return {}, {"error": f"Too few answer circles: {circles.shape[0]}"}

    xs = circles[:, 0].astype(np.float32)
    ys = circles[:, 1].astype(np.float32)

    _, row_labels = kmeans_1d(ys, T.QUESTIONS_PER_BLOCK)
    _, block_labels = kmeans_1d(xs, T.BLOCKS)

    answers: Dict[int, str] = {}
    assign_debug: List[Dict] = []

    for b in range(T.BLOCKS):
        idxs = np.where(block_labels == b)[0]
        if idxs.size == 0:
            continue

        x_block = xs[idxs]

        # In normal cases each block has all A-D chosen at least once.
        # If not, fall back to nearest quantization around the block span.
        if idxs.size >= 12:
            opt_centers, opt_labels = kmeans_1d(x_block, len(T.OPTIONS))
            for j, idx in enumerate(idxs):
                row_i = int(row_labels[idx])
                q = b * T.QUESTIONS_PER_BLOCK + row_i + 1
                answers[q] = T.OPTIONS[int(opt_labels[j])]
                assign_debug.append(
                    {
                        "q": q,
                        "block": b,
                        "row": row_i,
                        "x": float(xs[idx]),
                        "y": float(ys[idx]),
                        "option": answers[q],
                        "mode": "kmeans",
                    }
                )
        else:
            lo, hi = float(np.min(x_block)), float(np.max(x_block))
            if hi <= lo:
                continue
            span = hi - lo
            centers = np.array([lo + span * (i + 0.5) / 4.0 for i in range(4)], dtype=np.float32)
            for idx in idxs:
                row_i = int(row_labels[idx])
                q = b * T.QUESTIONS_PER_BLOCK + row_i + 1
                oi = int(np.argmin(np.abs(centers - xs[idx])))
                answers[q] = T.OPTIONS[oi]
                assign_debug.append(
                    {
                        "q": q,
                        "block": b,
                        "row": row_i,
                        "x": float(xs[idx]),
                        "y": float(ys[idx]),
                        "option": answers[q],
                        "mode": "fallback",
                    }
                )

    return answers, {"detected_circles": int(circles.shape[0]), "assigned": len(answers), "rows": 20, "blocks": 5, "assignments": assign_debug}


def disk_mask(h: int, w: int, radius_frac: float) -> np.ndarray:
    r = max(1, int(radius_frac * min(h, w)))
    cy, cx = h // 2, w // 2
    yy, xx = np.ogrid[:h, :w]
    return ((yy - cy) ** 2 + (xx - cx) ** 2 <= r * r)


def compute_name_scores(name_patch: np.ndarray) -> np.ndarray:
    # Otsu helps stabilize across different scans while staying binary.
    _, bw = cv2.threshold(name_patch, 0, 1, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    rows, cols = T.NAME_ROWS, T.NAME_COLS
    h, w = bw.shape
    cell_h = h / rows
    cell_w = w / cols

    scores = np.zeros((rows, cols), dtype=np.float32)

    for r in range(rows):
        y1 = int(round(r * cell_h))
        y2 = int(round((r + 1) * cell_h))
        for c in range(cols):
            x1 = int(round(c * cell_w))
            x2 = int(round((c + 1) * cell_w))
            cell = bw[y1:y2, x1:x2]
            if cell.size == 0:
                continue
            m = disk_mask(cell.shape[0], cell.shape[1], T.NAME_DISK_RADIUS_FRAC)
            vals = cell[m]
            scores[r, c] = float(vals.mean()) if vals.size else 0.0

    return scores


def decode_name_from_scores(scores: np.ndarray) -> Tuple[str, Dict]:
    rows, cols = scores.shape
    chars: List[str] = []
    per_col: List[Dict] = []

    for c in range(cols):
        col = scores[:, c]
        order = np.argsort(-col)
        best_i = int(order[0])
        best = float(col[best_i])
        second = float(col[int(order[1])]) if rows > 1 else 0.0

        if best < T.NAME_BLANK_THRESHOLD or (best - second) < T.NAME_DELTA_THRESHOLD:
            ch = ""
            conf = 0.0
        else:
            ch = chr(ord("A") + best_i)
            conf = best - second

        chars.append(ch)
        per_col.append({"col": c, "best_row": best_i, "best": best, "second": second, "confidence": conf, "char": ch})

    raw = "".join(ch if ch else "_" for ch in chars)
    decoded = "".join(chars).rstrip("_").strip()
    return decoded, {"raw": raw, "per_col": per_col}


def decode_name_with_offset(scores: np.ndarray, row_offset: int) -> str:
    rows, cols = scores.shape
    chars: List[str] = []
    for c in range(cols):
        col = scores[:, c]
        order = np.argsort(-col)
        best_i = int(order[0])
        best = float(col[best_i])
        second = float(col[int(order[1])]) if rows > 1 else 0.0
        if best < T.NAME_BLANK_THRESHOLD or (best - second) < T.NAME_DELTA_THRESHOLD:
            chars.append("")
            continue

        adj = max(0, min(25, best_i + row_offset))
        chars.append(chr(ord("A") + adj))
    return "".join(chars).rstrip("_").strip()


def score_sheet(student_answers: Dict[int, str], key: Dict[int, str]) -> Tuple[float, Dict]:
    total_q = T.BLOCKS * T.QUESTIONS_PER_BLOCK
    correct = wrong = blank = 0

    for q in range(1, total_q + 1):
        k = key.get(q)
        if k is None:
            continue
        s = student_answers.get(q)
        if s is None:
            blank += 1
        elif s == k:
            correct += 1
        else:
            wrong += 1

    score = correct * T.SCORE_CORRECT + wrong * T.SCORE_WRONG
    return float(score), {"correct": correct, "wrong": wrong, "blank": blank}


def save_answer_overlay(answer_patch: np.ndarray, circles: Optional[np.ndarray], out_path: Path) -> None:
    vis = cv2.cvtColor(answer_patch, cv2.COLOR_GRAY2BGR)
    if circles is not None:
        for x, y, r in circles:
            cv2.circle(vis, (int(x), int(y)), int(r), (0, 0, 255), 2)
            cv2.circle(vis, (int(x), int(y)), 1, (0, 255, 0), -1)
    cv2.imwrite(str(out_path), vis)


def save_answer_overlay_labeled(
    answer_patch: np.ndarray, assignments: List[Dict], key: Dict[int, str], out_path: Path
) -> None:
    vis = cv2.cvtColor(answer_patch, cv2.COLOR_GRAY2BGR)

    for item in assignments:
        q = int(item["q"])
        opt = str(item["option"])
        x = int(round(float(item["x"])))
        y = int(round(float(item["y"])))
        k = key.get(q)

        if k is None:
            color = (255, 200, 0)
        elif k == opt:
            color = (0, 170, 0)
        else:
            color = (0, 0, 255)

        cv2.circle(vis, (x, y), 11, color, 2)
        cv2.putText(
            vis,
            f"{q}:{opt}",
            (x + 8, y - 6),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.34,
            color,
            1,
            cv2.LINE_AA,
        )

    cv2.imwrite(str(out_path), vis)


def save_name_heatmap(scores: np.ndarray, out_path: Path) -> None:
    s = np.clip(scores, 0.0, 1.0)
    img = (s * 255).astype(np.uint8)
    img = cv2.resize(img, None, fx=10, fy=10, interpolation=cv2.INTER_NEAREST)
    cv2.imwrite(str(out_path), img)


def save_name_overlay(name_patch: np.ndarray, name_dbg: Dict, out_path: Path) -> None:
    vis = cv2.cvtColor(name_patch, cv2.COLOR_GRAY2BGR)
    h, w = name_patch.shape
    rows, cols = T.NAME_ROWS, T.NAME_COLS
    cell_h = h / rows
    cell_w = w / cols

    for c in range(cols + 1):
        x = int(round(c * cell_w))
        cv2.line(vis, (x, 0), (x, h - 1), (90, 90, 90), 1)
    for r in range(rows + 1):
        y = int(round(r * cell_h))
        cv2.line(vis, (0, y), (w - 1, y), (90, 90, 90), 1)

    for cinfo in name_dbg.get("per_col", []):
        c = int(cinfo["col"])
        r = int(cinfo["best_row"])
        ch = str(cinfo["char"])
        if not ch:
            continue
        x = int(round((c + 0.5) * cell_w))
        y = int(round((r + 0.5) * cell_h))
        cv2.circle(vis, (x, y), 8, (0, 180, 0), 2)
        cv2.putText(vis, ch, (x - 5, y + 4), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (0, 0, 255), 1, cv2.LINE_AA)

    cv2.imwrite(str(out_path), vis)


def save_decoded_answers_table(
    out_img_dir: Path, answers: Dict[int, str], key: Dict[int, str], assignments: List[Dict]
) -> None:
    pos_by_q: Dict[int, Tuple[float, float]] = {}
    for item in assignments:
        q = int(item["q"])
        pos_by_q[q] = (float(item["x"]), float(item["y"]))

    rows = []
    max_q = T.BLOCKS * T.QUESTIONS_PER_BLOCK
    for q in range(1, max_q + 1):
        pred = answers.get(q)
        k = key.get(q)
        if pred is None:
            status = "blank"
        elif k is None:
            status = "no_key"
        elif pred == k:
            status = "correct"
        else:
            status = "wrong"
        xy = pos_by_q.get(q)
        rows.append(
            {
                "q": q,
                "predicted": pred,
                "key": k,
                "status": status,
                "x": None if xy is None else round(xy[0], 2),
                "y": None if xy is None else round(xy[1], 2),
            }
        )

    pd.DataFrame(rows).to_csv(out_img_dir / "decoded_answers.csv", index=False)


def process_one_image(img_path: Path, key: Dict[int, str]) -> Dict:
    img = read_gray(img_path)
    h, w = img.shape[:2]

    name_roi = scale_roi(T.NAME_ROI, w, h)
    answer_roi = scale_roi(T.ANSWER_ROI, w, h)

    name_patch = crop(img, name_roi)
    answer_patch = crop(img, answer_roi)

    name_scores = compute_name_scores(name_patch)
    decoded_name, name_dbg = decode_name_from_scores(name_scores)

    circles = detect_answer_circles(answer_patch)
    answers: Dict[int, str] = {}
    ans_dbg: Dict = {}

    if circles is not None:
        answers, ans_dbg = decode_answers_from_circles(circles)
    else:
        ans_dbg = {"error": "No circles detected in answer ROI"}

    score, score_dbg = score_sheet(answers, key)

    out_img_dir = DEBUG_DIR / img_path.stem
    ensure_dir(out_img_dir)

    cv2.imwrite(str(out_img_dir / "roi_name.png"), name_patch)
    cv2.imwrite(str(out_img_dir / "roi_answers.png"), answer_patch)
    save_name_heatmap(name_scores, out_img_dir / "heatmap_name.png")
    save_name_overlay(name_patch, name_dbg, out_img_dir / "overlay_name_grid.png")
    save_answer_overlay(answer_patch, circles, out_img_dir / "overlay_answer_circles.png")
    assignments = ans_dbg.get("assignments", []) if isinstance(ans_dbg, dict) else []
    save_answer_overlay_labeled(answer_patch, assignments, key, out_img_dir / "overlay_answer_labeled.png")
    save_decoded_answers_table(out_img_dir, answers, key, assignments)

    pd.DataFrame(name_scores).to_csv(out_img_dir / "name_scores.csv", index=False)
    pd.DataFrame(name_dbg.get("per_col", [])).to_csv(out_img_dir / "name_columns.csv", index=False)
    name_hyp = [{"row_offset": off, "decoded": decode_name_with_offset(name_scores, off)} for off in range(-8, 9)]
    with (out_img_dir / "name_hypotheses.json").open("w", encoding="utf-8") as f:
        json.dump(name_hyp, f, indent=2)

    dump = {
        "image": img_path.name,
        "shape": {"w": w, "h": h},
        "name_roi": name_roi,
        "answer_roi": answer_roi,
        "name": {"decoded": decoded_name, "debug": name_dbg},
        "answers": {
            "decoded_count": len(answers),
            "first_20": {str(q): answers.get(q) for q in range(1, 21)},
            "debug": ans_dbg,
        },
        "score": {"value": score, "breakdown": score_dbg},
    }

    with (out_img_dir / "debug_dump.json").open("w", encoding="utf-8") as f:
        json.dump(dump, f, indent=2)

    return {
        "image": img_path.name,
        "name": decoded_name,
        "score": score,
        "correct": score_dbg["correct"],
        "wrong": score_dbg["wrong"],
        "blank": score_dbg["blank"],
        "decoded_answers": len(answers),
    }


def main() -> None:
    ensure_dir(OUT_DIR)
    ensure_dir(DEBUG_DIR)

    bmps = pick_bmps(BMP_DIR)
    xlsx = pick_first_xlsx(KEY_DIR)
    key = load_answer_key_xlsx(xlsx)
    save_answer_key_debug(key, xlsx)

    print(f"[KEY] {xlsx.name} | entries={len(key)}")
    print(f"[KEY] wrote {OUT_DIR / 'answer_key_parsed.csv'}")
    print(f"[KEY] wrote {OUT_DIR / 'answer_key_parsed.json'}")
    print(f"[OMR] processing {len(bmps)} files from {BMP_DIR}")

    rows: List[Dict] = []
    for bmp in bmps:
        row = process_one_image(bmp, key)
        rows.append(row)
        print(
            f"  - {bmp.name}: name='{row['name']}', score={row['score']:.2f}, "
            f"correct={row['correct']}, wrong={row['wrong']}, blank={row['blank']}"
        )

    df = pd.DataFrame(rows)
    csv_path = OUT_DIR / "results.csv"
    json_path = OUT_DIR / "results.json"

    df.to_csv(csv_path, index=False)
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2)

    print(f"[OUT] wrote {csv_path}")
    print(f"[OUT] wrote {json_path}")


if __name__ == "__main__":
    main()
