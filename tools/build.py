"""Проверка исходников и сборка статьи в Word через Pandoc."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parents[1]


def read_project() -> tuple[dict, list[Path]]:
    config = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
    missing = [key for key in ("title", "author", "group")
               if not isinstance(config.get(key), str) or not config[key].strip()]
    if missing:
        raise ValueError("Заполните project.json: " + ", ".join(missing))
    if config.get("article_url"):
        url = urlparse(config["article_url"])
        if url.scheme not in ("http", "https") or not url.netloc:
            raise ValueError("article_url должна содержать ссылку http:// или https://.")
    if not re.fullmatch(r"[A-Za-z0-9_-]+", config.get("output_name", "")):
        raise ValueError("output_name: используйте латинские буквы, цифры, - или _.")
    paths = []
    for line in (ROOT / "contents").read_text(encoding="utf-8").splitlines():
        name = line.strip()
        if not name or name.startswith("#"):
            continue
        path = (ROOT / name).resolve()
        if not path.is_relative_to(ROOT) or not path.is_file():
            raise ValueError(f"Не найден исходник внутри проекта: {name}")
        text = re.sub(r"<!--.*?-->", "", path.read_text(encoding="utf-8"), flags=re.S)
        if not text.strip():
            raise ValueError(f"Добавьте текст статьи: {name}")
        paths.append(path)
    if not paths:
        raise ValueError("В файле contents не перечислены исходники.")
    for name in ("makefile", "defaults.yaml", "bibliography.bib", "numeric-ru.csl", "template.docx"):
        if not (ROOT / name).is_file():
            raise ValueError(f"Не найден файл проекта: {name}")
    source = "\n".join(p.read_text(encoding="utf-8") for p in paths)
    cited = set(re.findall(r"@([A-Za-z][A-Za-z0-9_-]+)", source))
    references = set(re.findall(r"@\w+\s*\{\s*([^,]+),", (ROOT / "bibliography.bib").read_text(encoding="utf-8")))
    if cited - references:
        raise ValueError("Не найдены источники: " + ", ".join(sorted(cited - references)))
    if len(paths) < 2:
        raise ValueError("Документация должна состоять из нескольких Markdown-файлов.")
    if len(re.findall(r"\$\$.+?\$\$", source, flags=re.S)) < 2:
        raise ValueError("Добавьте формулы LaTeX в блоках $$ ... $$.")
    figures = re.findall(r"!\[[^\]]*\]\(([^\s)]+)\)", source)
    if not figures:
        raise ValueError("Добавьте рисунок в документацию Markdown.")
    for name in figures:
        picture = (ROOT / name).resolve()
        if not picture.is_relative_to(ROOT) or not picture.is_file():
            raise ValueError(f"Не найден рисунок внутри проекта: {name}")
    if not re.search(r"^\s*\|(?:\s*:?-{3,}:?\s*\|){2,}\s*$", source, flags=re.M):
        raise ValueError("Добавьте таблицу в документацию Markdown.")
    return config, paths


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("target", choices=("check", "docx"), nargs="?", default="docx")
    args = parser.parse_args()
    try:
        if args.target == "docx":
            from draw_search_pipeline import generate
            generate()
        config, paths = read_project()
        if args.target == "check":
            print(f"Исходники и метаданные проверены. Файлов: {len(paths)}.")
            return 0
        pandoc = os.environ.get("PANDOC") or shutil.which("pandoc")
        if not pandoc:
            portable = sorted((ROOT / ".tools").glob("pandoc-*/pandoc.exe"))
            if portable:
                pandoc = str(portable[-1])
        if not pandoc:
            raise ValueError("Установите Pandoc и добавьте его в PATH: https://pandoc.org/installing.html")
        output = ROOT / "build" / f"{config['output_name']}.docx"
        output.parent.mkdir(exist_ok=True)
        command = [pandoc, "--defaults", str(ROOT / "defaults.yaml"),
                   *map(str, paths), "--output", str(output)]
        template = ROOT / "template.docx"
        if template.is_file():
            command.extend(["--reference-doc", str(template)])
        subprocess.run(command, cwd=ROOT, check=True)
        from format_docx import format_document
        format_document(output, config)
        print(f"Создан документ: {output}")
        return 0
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        print(f"Ошибка сборки: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
