# -*- coding: utf-8 -*-
"""
河南专升本数据爬虫（数据来源：河南专升本网 lemedu.com 的公开查询接口）

功能：
  - 自动拉取全部“本科专业”列表
  - 按专业抓取：各院校历年录取分数、招生计划、省控线
  - 结果缓存为 JSON（可断点续跑），并可直接导出 Excel

用法示例：
  python scraper.py --all                 # 抓取全部本科专业
  python scraper.py --list                # 只打印全部专业，不抓取
  python scraper.py --majors "计算机科学与技术,法学"
  python scraper.py --file majors.txt     # 每行一个专业名
  python scraper.py --all --refresh        # 忽略缓存，重新抓取
  python scraper.py --all --workers 8      # 并发数（默认 8）

导出 Excel：
  python build_excel.py                    # 读取 cache/ 生成 Excel
"""
import argparse
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

# ---------------- 接口配置（如网站改版，只需改这里） ----------------
SERVER96 = "http://81.70.180.96:8080/"
SERVER91 = "http://120.77.55.191:8080/"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Referer": "http://www.lemedu.com/lqfs.html",
}
DEFAULT_CACHE = "cache"

_session = requests.Session()
_session.headers.update(HEADERS)


def _get_json(url, retries=3, timeout=30):
    last = None
    for i in range(retries):
        try:
            r = _session.get(url, timeout=timeout)
            r.encoding = "utf-8"
            return json.loads(r.text)
        except Exception as e:
            last = e
            time.sleep(1.2 * (i + 1))
    print("  ! 抓取失败:", url, last, file=sys.stderr)
    return None


def quote(s):
    return requests.utils.quote(s, safe="")


def fetch_majors():
    """返回全部本科专业名称列表。"""
    data = _get_json(SERVER96 + "search/bar?level=" + quote("本科") + "&name=")
    return data if isinstance(data, list) else []


def fetch_major_detail(major):
    """返回单个专业的全部原始数据。"""
    return {
        "major": major,
        "schoolScore": _get_json(SERVER91 + "search/schoolScore?name=" + quote(major)),
        "plan": _get_json(SERVER96 + "search/plan?year=2019&name=" + quote(major)),
        "provinceScore": _get_json(
            SERVER96 + "provinceScore/queryProvinceScore?benkeName=" + quote(major)
        ),
    }


# ---------------- 缓存 ----------------
def safe_name(major):
    return "".join(c if c not in '\\/:*?"<>|' else "_" for c in major)


def cache_file(cache_dir, major):
    return os.path.join(cache_dir, safe_name(major) + ".json")


def load_cache(cache_dir, major):
    p = cache_file(cache_dir, major)
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def save_cache(cache_dir, major, data):
    os.makedirs(cache_dir, exist_ok=True)
    with open(cache_file(cache_dir, major), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


# ---------------- 主流程 ----------------
def scrape(majors, cache_dir, workers=8, refresh=False):
    os.makedirs(cache_dir, exist_ok=True)

    to_fetch = [] if refresh else []
    for m in majors:
        if not refresh and load_cache(cache_dir, m) is not None:
            continue
        to_fetch.append(m)

    cached = len(majors) - len(to_fetch)
    print("专业总数: {}，已缓存: {}，待抓取: {}".format(len(majors), cached, len(to_fetch)))
    if not to_fetch:
        print("全部已缓存，无需抓取。")
        return

    done = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(fetch_major_detail, m): m for m in to_fetch}
        for fut in as_completed(futs):
            m = futs[fut]
            try:
                data = fut.result()
            except Exception as e:
                print("  ! {} 抓取异常: {}".format(m, e), file=sys.stderr)
                continue
            save_cache(cache_dir, m, data)
            done += 1
            print("[{}/{}] {} 完成".format(done, len(to_fetch), m))


def read_major_file(path):
    with open(path, "r", encoding="utf-8") as f:
        return [ln.strip() for ln in f if ln.strip()]


def main():
    ap = argparse.ArgumentParser(description="河南专升本数据爬虫")
    ap.add_argument("--all", action="store_true", help="抓取全部本科专业")
    ap.add_argument("--list", action="store_true", help="只列出全部本科专业")
    ap.add_argument("--majors", type=str, default="", help="指定专业，逗号分隔")
    ap.add_argument("--file", type=str, default="", help="专业清单文件，每行一个")
    ap.add_argument("--cache", type=str, default=DEFAULT_CACHE, help="缓存目录")
    ap.add_argument("--workers", type=int, default=8, help="并发线程数")
    ap.add_argument("--refresh", action="store_true", help="忽略缓存重新抓取")
    args = ap.parse_args()

    all_majors = fetch_majors()
    if args.list:
        for m in all_majors:
            print(m)
        print("共 {} 个专业".format(len(all_majors)), file=sys.stderr)
        return

    majors = []
    if args.all:
        majors = all_majors
    if args.majors:
        majors += [m.strip() for m in args.majors.split(",") if m.strip()]
    if args.file:
        majors += read_major_file(args.file)

    if not majors:
        ap.print_help()
        print("\n未指定专业，如需全部请加 --all", file=sys.stderr)
        return

    seen = set()
    majors = [m for m in majors if not (m in seen or seen.add(m))]
    scrape(majors, args.cache, workers=args.workers, refresh=args.refresh)


if __name__ == "__main__":
    main()
