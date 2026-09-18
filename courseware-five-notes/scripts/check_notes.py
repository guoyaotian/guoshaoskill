"""Validate five numbered Markdown notes using a conservative character count."""

import argparse
import re
from pathlib import Path


HEADING = re.compile(r"^#\s+(?:📝\s*)?笔记\s*([0-9]+)\s*[：:｜|].+$", re.MULTILINE)


def validate(text):
    headings = list(HEADING.finditer(text))
    errors = []
    counts = []
    numbers = [int(item.group(1)) for item in headings]
    if numbers != [1, 2, 3, 4, 5]:
        errors.append(f"Expected notes 1-5 exactly once in order; found {numbers}.")
    if headings and text[:headings[0].start()].strip():
        errors.append("Place only the five notes in the draft; remove the preamble.")
    if len(re.findall(r"^#\s+", text, re.MULTILINE)) != len(headings):
        errors.append("Unrecognized level-1 heading; use '# 📝 笔记1｜Title'.")
    for index, heading in enumerate(headings):
        end = headings[index + 1].start() if index + 1 < len(headings) else len(text)
        section = text[heading.start():end]
        count = sum(not char.isspace() for char in section)
        number = int(heading.group(1))
        counts.append((number, count))
        if not text[heading.end():end].strip():
            errors.append(f"Note {number} has no body.")
        if count > 600:
            errors.append(f"Note {number} exceeds 600 characters by {count - 600}.")
    return counts, errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("draft", type=Path, help="UTF-8 Markdown containing only five notes")
    args = parser.parse_args()
    counts, errors = validate(args.draft.read_text(encoding="utf-8-sig"))
    for number, count in counts:
        print(f"Note {number}: {count}/600 non-whitespace characters")
    for error in errors:
        print(f"ERROR: {error}")
    if not errors:
        print("PASS: five ordered, nonempty notes within the character limit.")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
