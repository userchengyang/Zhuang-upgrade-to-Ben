# 河南专升本数据爬虫

爬取河南专升本网（lemedu.com）公开查询接口中，各本科专业对应院校的历年录取分数、招生计划与省控线，并整理成 Excel。

## 功能

- 自动获取全部本科专业列表
- 按专业抓取：各院校历年录取分数、招生计划（学制/学费/住宿费）、省控线
- 抓取结果缓存为 JSON，支持断点续跑
- 一键导出 Excel（三个工作表：录取分数线 / 招生计划 / 省控线）

## 环境

- Python 3.7+
- 依赖见 `requirements.txt`

```bash
pip install -r requirements.txt
```

## 使用

### 抓取数据

```bash
# 抓取全部本科专业
python scraper.py --all

# 只列出全部本科专业
python scraper.py --list

# 抓取指定专业（逗号分隔）
python scraper.py --majors "计算机科学与技术,法学,临床医学"

# 从文件读取专业清单（每行一个）
python scraper.py --file majors.txt

# 忽略缓存，重新抓取；指定并发数
python scraper.py --all --refresh --workers 8
```

抓到的原始数据保存在 `cache/` 目录，每个专业一个 JSON 文件。重复运行会自动跳过已抓取的专业。

### 导出 Excel

```bash
python build_excel.py
# 或指定缓存目录与输出文件
python build_excel.py --cache cache --out 河南专升本数据.xlsx
```

## 输出说明

`河南专升本数据.xlsx` 包含三张表：

| 工作表 | 内容 |
| --- | --- |
| 录取分数线 | 每个「专业 + 院校」的历年录取分数，`×` 表示当年无招生计划，`/` 表示分数未公布 |
| 招生计划 | 每个「专业 + 院校」的学制、学费、住宿费、招生计划、地址 |
| 省控线 | 每个专业历年的省控线及对应招生人数 |

## 参数

`scraper.py`

| 参数 | 说明 |
| --- | --- |
| `--all` | 抓取全部本科专业 |
| `--list` | 仅列出全部本科专业 |
| `--majors "A,B"` | 指定专业，逗号分隔 |
| `--file PATH` | 专业清单文件，每行一个 |
| `--cache DIR` | 缓存目录（默认 `cache`） |
| `--workers N` | 并发线程数（默认 8） |
| `--refresh` | 忽略缓存重新抓取 |

`build_excel.py`

| 参数 | 说明 |
| --- | --- |
| `--cache DIR` | 缓存目录（默认 `cache`） |
| `--out FILE` | 输出 Excel 文件名 |

## 免责声明

数据来源于河南专升本网及各院校官网公开信息，仅供学习参考，请以院校官方公布的数据为准。
