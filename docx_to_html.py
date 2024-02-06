import argparse
import glob
import html
import json
import os
import posixpath
import shutil
from posixpath import join
import re
from typing import Optional, List, Tuple

import mammoth
from mammoth.cli import ImageWriter
from tqdm import tqdm


def main(path: str, html_path: str, image_path: str, image_path_static: str):
    os.makedirs(html_path, exist_ok=True)
    os.makedirs(image_path, exist_ok=True)

    for _path in [html_path, image_path]:
        for filepath in glob.glob(posixpath.join(_path, "*")):
            if ".gitignore" not in filepath:
                try:
                    shutil.rmtree(filepath)
                except NotADirectoryError:
                    os.remove(filepath)

    filenames = sorted(os.listdir(path))
    pbar = tqdm(filenames)
    for filename in pbar:
        if filename.endswith(".docx"):
            pbar.set_postfix(filename=filename)
            new_filename = filename.replace(".docx", "")
            _image_path = join(image_path, new_filename)
            if not os.path.exists(_image_path):
                os.makedirs(_image_path)
            image_converter = mammoth.images.img_element(ImageWriter(_image_path))
            with open(join(path, filename), "rb") as docx_file:
                result = mammoth.convert(
                    docx_file,
                    convert_image=image_converter,
                    output_format="html",
                )
            if len(os.listdir(_image_path)) == 0:
                os.rmdir(_image_path)
            text = result.value
            text = clean(text)
            text = html.unescape(text)
            metadata, text = extract_metadata(text)

            lines = Concatenator().concatenate(text)
            lines = tag_html(lines)
            lines = convert_footnotes(lines)
            lines = fix_images(lines, image_path=join(image_path_static, new_filename))

            filepath_json = join(html_path, new_filename + ".json")
            filepath_html = join(html_path, new_filename + ".html")
            with open(filepath_json, "w") as f:
                json.dump(metadata, f)
                f.write("\n")
            with open(filepath_html, "w", encoding="utf-8") as f:
                for line in lines:
                    f.write(line)
                    f.write("\n")


def clean(text: str) -> str:
    text = text.replace("<strong> </strong>", " ")
    text = re.sub(r"\s+", " ", text)
    return text


re_ends_with_hyphen = re.compile(r"(?<=\w)-\s?$")
re_ordered_list = re.compile(r"^(?:(\d+|[IVX]+|[a-z])[).]|\[(\d+)\])\s")


def extract_metadata(text: str) -> tuple[dict[str, str], str]:
    metadata: dict[str, str] = {}
    for field, regex in [
        ("titel", r"^<p>Titel: ([^<]+)</p>"),
        ("jaar", r"^<p>Jaar: ([^<]+)</p>"),
        ("omschrijving", r"^<p>Omschrijving: ([^<]+)</p>"),
        ("categorie", r"^<p>Categorie: ([^<]+)</p>"),
        ("afkomstig uit", r"^<p>Afkomstig uit: ([^<]+)</p>"),
    ]:
        match = re.search(regex, text, flags=re.IGNORECASE)
        if match is not None:
            metadata[field] = match.group(1)
        else:
            metadata[field] = ""
        text = re.sub(regex, "", text, flags=re.IGNORECASE)
    text = re.sub(r"<p>Tekst:\s?</p>", "", text, flags=re.IGNORECASE)
    return metadata, text


class Concatenator:
    """Make sure lines are concatenated properly."""

    def __init__(self):
        self.out: List[str] = []
        self.line: List[str] = []
        self.image_buffer: Optional[str] = None
        self.line_tags: Optional[Tuple[str, str]] = None

    def concatenate(self, text: str) -> List[str]:
        parts = text.split("</p><p>")
        for i, part in enumerate(parts):
            part = part.strip()
            part_is_empty = len(part) == 0
            if part_is_empty:
                self.flush()
                continue

            if part.startswith("<em>") and part.endswith("</em>"):
                part = part[4:-5].strip()
                self.line_tags = ("<em>", "</em>")
            previous_part = parts[i - 1] if i > 0 else None
            if previous_part is not None:
                previous_part = previous_part.replace("<em>", "").replace("</em>", "").strip()
            next_part = parts[i + 1] if i < len(parts) - 1 else None
            if next_part is not None:
                next_part = next_part.replace("<em>", "").replace("</em>", "").strip()

            is_image = part.startswith("<img ") and part.endswith(">")
            if is_image:
                self.image_buffer = part
                if not self.line:
                    self.flush()
                continue

            if i == 0 and part.startswith("<p>"):
                part = part[3:]
            elif i == len(parts) - 1 and part.endswith("</p>"):
                part = part[:-4]

            is_numbered_list = re.search(re_ordered_list, part) is not None
            is_title = i == 0
            is_heading = (
                re.match(r"^<strong>.+<\/strong>[.\s]*$", part) is not None or part.isupper()
            ) and i != len(parts) - 1

            if is_numbered_list or is_title or is_heading:
                self.flush()
            if is_title or is_heading:
                part = part.replace("<strong>", "").replace("</strong>", "")
                part = part.title()
            if is_title:
                part = "<h1>" + part + "</h1>"
            elif is_heading:
                part = "<h2>" + part + "</h2>"

            previous_ends_with_hyphen = (
                previous_part is not None
                and re.search(re_ends_with_hyphen, previous_part) is not None
            )
            if previous_ends_with_hyphen:
                self.line[-1] = re.sub(re_ends_with_hyphen, part, self.line[-1])
            else:
                self.line.append(part)

            part_ends_with_punctiation = re.search(r"[.!?]\s?$", part) is not None
            next_part_starts_with_capital_letter = next_part and next_part[0].isupper()
            next_part_is_image = next_part is not None and next_part.startswith("<img ")
            if (
                is_title
                or is_heading
                or is_image
                or (
                    (part_ends_with_punctiation or i < 3)
                    and (next_part_starts_with_capital_letter or next_part_is_image)
                )
            ):
                self.flush()
        self.flush()
        return self.out

    def flush(self):
        if self.line:
            line = " ".join(self.line).strip()
            if self.line_tags is not None:
                line = self.line_tags[0] + line + self.line_tags[1]
                self.line_tags = None
            self.out.append(line)
            self.line = []
        if self.image_buffer is not None:
            self.out.append(self.image_buffer)
            self.image_buffer = None


def tag_html(lines: list[str]) -> list[str]:
    out: List[str] = []
    numbered_list_types: list[str] = []  # keep track of lists we are currently in

    for i, part in enumerate(lines):
        assert len(part) != 0

        if numbered_list_types and part.startswith("<img "):
            # an image can't be part of a list, so first close the open list
            out.append("</ol>")
            numbered_list_types.pop()

        if part.startswith(("<img ", "<ul", "<ol", "<li", "<h")):
            out.append(part)
            continue

        match_numbered_list = re.search(re_ordered_list, part)
        is_numbered_list = match_numbered_list is not None
        if is_numbered_list:
            list_index_str = match_numbered_list.group(1) or match_numbered_list.group(2)  # type: ignore
            if list_index_str.isnumeric():
                list_type = "decimal"
                list_index = int(list_index_str)
            elif is_uppercase_roman_numeral(list_index_str):
                list_type = "upper-roman"
                list_index = roman_numeral_to_integer(list_index_str)
            elif list_index_str.isalpha():
                list_type = "lower-alpha"
                list_index = letter_to_integer(list_index_str)
            else:
                raise ValueError(f"Unknown list type: {list_index_str}")
            if numbered_list_types:
                if list_type not in numbered_list_types:
                    # start a deeper level
                    out.append("<ol>")
                    numbered_list_types.append(list_type)
                elif list_type == numbered_list_types[-1]:
                    # continue on the same level
                    pass
                else:
                    # end a level
                    out.append("</ol>")
                    numbered_list_types.pop()
            else:
                numbered_list_types.append(list_type)
                out.append("<ol>")
            part = re.sub(re_ordered_list, "", part)
            part = f'<li value="{list_index}" style="list-style-type:{list_type}">{part}</li>'
            out.append(part)
            continue
        elif not is_numbered_list and numbered_list_types:
            for _ in range(len(numbered_list_types)):
                out.append("</ol>")
            numbered_list_types = []

        out.append("<p>" + part + "</p>")

    return out


def is_uppercase_roman_numeral(value: str) -> bool:
    return not bool(set(value) - set("IVX"))


def roman_numeral_to_integer(value: str) -> int:
    return {
        "I": 1,
        "II": 2,
        "III": 3,
        "IV": 4,
        "V": 5,
        "VI": 6,
        "VII": 7,
        "VIII": 8,
        "IX": 9,
        "X": 10,
        "XI": 11,
        "XII": 12,
        "XIII": 13,
        "XIV": 14,
        "XV": 15,
    }[value.upper()]


def letter_to_integer(value: str) -> int:
    return ord(value.lower()) - ord("a") + 1


def fix_images(lines: list[str], image_path: str) -> list[str]:
    out = []
    for line in lines:
        if not line.startswith("<img "):
            out.append(line)
            continue
        image_url = posixpath.join(
            image_path,
            re.search(r"src=\"([^\"]+)\"", line).group(1),  # type: ignore
        )
        new_img = f'<img src="{image_url}" style="max-width:100%; height:auto;" />'
        new_line = f'<a href="{image_url}" target="_blank">{new_img}</a>'
        out.append(new_line)
    return out


def convert_footnotes(lines: list[str]) -> list[str]:
    out_inverted: List[str] = []
    footnotes: dict[int, str] = {}
    finding_footnotes = True
    for line in lines[::-1]:
        match = re.search(r'^<li value="(\d+)"[^>]+>(.+)</li>$', line)
        if match is None:
            finding_footnotes = False
        elif finding_footnotes:
            footnote_index = int(match.group(1))
            footnote_text = match.group(2)
            line = line.replace("<li ", f'<li id="footnote-{footnote_index}" ')
            if footnote_index in footnotes:
                raise ValueError(f"Duplicate footnote: {footnote_index}")
            footnotes[footnote_index] = footnote_text
        out_inverted.append(line)

    def func(_match: re.Match[str]) -> str:
        idx = int(_match.group(1) or _match.group(3))
        suffix = _match.group(2) or ""
        for i in range(1, 4):
            if not footnotes or idx == min(footnotes.keys()):
                break
            if idx == min(footnotes.keys()) + i:
                # we are skipping some footnote
                for _ in range(i):
                    # print(f"Skipping footnote {min(footnotes.keys())}")
                    footnotes.pop(min(footnotes.keys()))
        if not footnotes or idx != min(footnotes.keys()) or idx not in footnotes:
            return str(_match.group(0))
        text = footnotes.pop(idx)
        return f' <a href="#footnote-{idx}" title="{text}">[{idx}]</a>{suffix}'

    out = []
    for line in out_inverted[::-1]:
        # Replace references to footnotes with links
        line = re.sub(r" (\d{1,3})\)([\s.,])|\[(\d+)\]", func, line)
        out.append(line)

    return out


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--path", help="Folder with docx files with publications.")
    parser.add_argument("--html-path", help="html output path")
    parser.add_argument("--image-path", help="path to store images")
    parser.add_argument("--image-path-static", help="static path for images used in html")
    options = parser.parse_args()

    main(
        path=options.path,
        html_path=options.html_path,
        image_path=options.image_path,
        image_path_static=options.image_path_static,
    )
