#!/usr/bin/env python3
"""
docx-official-format 子技能完整性自检。

检查必需文件是否齐全、有无误提交的缓存文件。
在 docx-official-format/ 目录下运行：python3 scripts/validate_skill.py
"""
from pathlib import Path
import sys

root = Path(__file__).resolve().parents[1]

# 子技能必需的文件清单
required = [
    root / 'SKILL.md',
    root / 'agents' / 'openai.yaml',
    root / 'references' / 'official-format.md',
    root / 'scripts' / 'format_docx.py',
    root / 'tests' / 'test_signature_detection.py',
]

missing = [str(p.relative_to(root)) for p in required if not p.exists()]
if missing:
    print('缺少必需文件：')
    for item in missing:
        print(f'  - {item}')
    sys.exit(1)

# 不应该出现在仓库里的缓存文件
bad = []
for pattern in ('__pycache__', '*.pyc', '.DS_Store'):
    bad.extend(str(p.relative_to(root)) for p in root.rglob(pattern))
if bad:
    print('发现不该提交的缓存文件：')
    for item in bad:
        print(f'  - {item}')
    sys.exit(1)

print('docx-official-format 自检通过 ✓')
