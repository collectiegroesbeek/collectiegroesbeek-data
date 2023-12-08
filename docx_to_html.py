import argparse
import os
import re
from typing import Optional, List

import mammoth
from mammoth.cli import ImageWriter
from tqdm import tqdm


def main(path: str, html_path: str, image_path: str):
    filenames = sorted(os.listdir(path))
    pbar = tqdm(filenames)
    for filename in pbar:
        if filename.endswith(".docx"):
            pbar.set_postfix(filename=filename)
            new_filename = filename.lower().replace(".docx", "").replace(" ", "_")
            _image_path = os.path.join(image_path, new_filename)
            if not os.path.exists(_image_path):
                os.makedirs(_image_path)
            image_converter = mammoth.images.img_element(ImageWriter(_image_path))
            with open(os.path.join(path, filename), "rb") as docx_file:
                result = mammoth.convert(
                    docx_file,
                    convert_image=image_converter,
                    output_format="html",
                )
            if len(os.listdir(_image_path)) == 0:
                os.rmdir(_image_path)
            text = result.value
            text = replace_characters(text)
            text = clean(text)
            lines = Concatenator().concatenate(text)
            text = tag_html(lines)
            text = fix_img_path(text, image_path=_image_path)
            filepath_html = os.path.join(html_path, new_filename + ".html")
            with open(filepath_html, "w", encoding="utf-8") as f:
                f.write(text)


def replace_characters(text: str) -> str:
    text = text.replace("„", "&bdquo;")
    text = text.replace("”", "&rdquo;")
    text = text.replace("“", "&ldquo;")
    text = text.replace("±", "&#177;")
    text = text.replace("é", "&eacute;")
    text = text.replace("†", "&dagger;")
    text = text.replace("ó", "&oacute;")
    text = text.replace("½", "&frac12;")
    text = text.replace("⅓", "&frac13;")
    text = text.replace("ë", "&euml;")
    return text


def clean(text: str) -> str:
    text = text.replace("<strong> </strong>", " ")
    return text


re_ends_with_hyphen = re.compile(r"(?<=\w)-\s?$")
re_ordered_list = re.compile(r"^(\d+|[IVX]+|[a-z])[).]\s")


class Concatenator:
    """Make sure lines are concatenated properly."""

    def __init__(self):
        self.out: List[str] = []
        self.line: List[str] = []
        self.image_buffer: Optional[str] = None

    def concatenate(self, text: str) -> List[str]:
        parts = text.split("</p><p>")
        for i, part in enumerate(parts):
            part_is_empty = len(part) == 0
            if part_is_empty:
                continue

            is_image = part.startswith("<img ") and part.endswith(">")
            if is_image:
                part = part.replace("/>", " style='max-width:100%; height:auto;' />")
                part = re.sub(r'alt="[\w\s,]+"', "", part)
                self.image_buffer = part
                continue

            if i == 0 and part.startswith("<p>"):
                part = part[3:]
            elif i == len(parts) - 1 and part.endswith("</p>"):
                part = part[:-4]

            is_numbered_list = re.search(re_ordered_list, part) is not None
            is_title = i == 0
            is_heading = part.startswith("<strong>") and part.endswith("</strong>")
            if is_numbered_list or is_title or is_heading:
                self.flush()

            previous_ends_with_hyphen = (
                i > 0 and re.search(re_ends_with_hyphen, parts[i - 1]) is not None
            )
            if previous_ends_with_hyphen:
                self.line[-1] = re.sub(
                    re_ends_with_hyphen, part, self.line[-1]
                )
            else:
                self.line.append(part)

            part_ends_with_punctiation = re.search(r"[.!?]\s?$", part) is not None
            next_part_starts_with_capital_letter = (
                i < len(parts) - 1 and parts[i + 1][0].isupper()
            )
            if (
                is_title
                or is_heading
                or is_image
                or (part_ends_with_punctiation and next_part_starts_with_capital_letter)
            ):
                self.flush()
        self.flush()
        return self.out

    def flush(self):
        if self.line:
            self.out.append("".join(self.line).strip())
            self.line = []
        if self.image_buffer is not None:
            self.out.append(self.image_buffer)
            self.image_buffer = None


def tag_html(lines: List[str]) -> str:
    out: List[str] = []
    numbered_list_types = []

    for i, part in enumerate(lines):
        assert len(part) != 0

        if part.startswith(("<img ", "<ul", "<ol", "<li")):
            out.append(part)
            continue

        match_numbered_list = re.search(re_ordered_list, part)
        is_numbered_list = match_numbered_list is not None
        if is_numbered_list:
            list_index_str = match_numbered_list.group(1)
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

    result = "\n".join(out)
    return result


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
    }[value.upper()]


def letter_to_integer(value: str) -> int:
    return ord(value.lower()) - ord("a") + 1


def fix_img_path(text: str, image_path: str) -> str:
    text = text.replace("<img src=\"", f"<img src=\"{image_path}/")
    return text


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--path', help='Folder with docx files with publications.')
    parser.add_argument("--html-path", help="html output path")
    parser.add_argument("--image-path", help="path to store images")
    options = parser.parse_args()

    main(path=options.path, html_path=options.html_path, image_path=options.image_path)
