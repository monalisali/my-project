# -*- coding: utf-8 -*-
"""Merge per-batch digest markdown files into one deliverable markdown."""
import re
from datetime import date

DIGESTS = [
    r"C:\D\AIWork\my-project\.work\digests\digest_A.md",
    r"C:\D\AIWork\my-project\.work\digests\digest_B.md",
    r"C:\D\AIWork\my-project\.work\digests\digest_C.md",
    r"C:\D\AIWork\my-project\.work\digests\digest_D.md",
    r"C:\D\AIWork\my-project\.work\digests\digest_E.md",
]
OUT = r"C:\D\AIWork\my-project\translation\金钱心理学_中文章节导读.md"

HEADER = """# 金钱心理学
## The Psychology of Money —— 中文章节导读

> **原书作者**：Morgan Housel（摩根·豪泽尔）
> **原书出版**：Harriman House, 2020
> **文档性质**：原创中文解读导读（非译本），生成日期 {date}

---

## 导读说明

- 本文档为《The Psychology of Money》（中译名《金钱心理学》）的**原创中文解读导读**，供学习与研究参考。
- 本文档**并非原书的中文译本**：内容为对各章思想与论点的概括提炼，以解读者的原创语言撰写；如需完整内容，请阅读正版原书或其官方授权中译本。
- 原书中的插图、图表等图片元素未收录于本文档，各章仅解读文字内容。
- 财经专业名词与关键术语在首次出现时标注英文原文，例如：复利（Compounding）、安全边际（Margin of Safety）。
- 每章依次包含：**本章主旨、核心观点、关键故事与案例、术语与概念、实践启示**五个部分。
- 全书结构：引言 + 20 章 + 后记。
"""

def main():
    parts = [HEADER.format(date=date.today().isoformat())]
    n_ch = 0
    for path in DIGESTS:
        with open(path, encoding="utf-8") as f:
            text = f.read().strip()
        n_ch += len(re.findall(r"(?m)^# ", text))
        parts.append(text)
    merged = "\n\n---\n\n".join(parts) + "\n"
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(merged)
    print("saved:", OUT, "| chapters:", n_ch, "| chars:", len(merged))

if __name__ == "__main__":
    main()
