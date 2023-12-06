import os
import re
from typing import Optional

import mammoth
from mammoth.cli import ImageWriter


def main():
    path = r"D:\Dropbox\Anemashare\2023-11-29 publicaties Coll Groesbeek"
    image_converter = mammoth.images.img_element(ImageWriter("."))
    for filename in ["Tedingh van Cranenburg.docx"]:  # sorted(os.listdir(path)):
        if filename.endswith(".docx"):
            with open(os.path.join(path, filename), "rb") as docx_file:
                result = mammoth.convert(
                    docx_file,
                    convert_image=image_converter,
                    output_format="html",
                )
                text = result.value
                text = replace_characters(text)
                text = Concatenator().concatenate(text)
                with open("out.html", "w", encoding="utf-8") as f:
                    f.write(text)
                a = 1
                break


def replace_characters(text: str) -> str:
    text = text.replace("„", "&bdquo;")
    text = text.replace("”", "&rdquo;")
    text = text.replace("±", "&#177;")
    text = text.replace("é", "&eacute;")
    text = text.replace("†", "&dagger;")
    text = text.replace("ó", "&oacute;")
    text = text.replace("½", "&frac12;")
    text = text.replace("⅓", "&frac13;")
    text = text.replace("ë", "&euml;")
    return text


re_ends_with_hyphen = re.compile(r"(?<=\w)-\s?$")


class Concatenator:
    def __init__(self):
        self.out = []
        self.paragraph = []
        self.image_buffer: Optional[str] = None

    def concatenate(self, text: str) -> str:
        parts = text.split("</p><p>")
        for i, part in enumerate(parts):
            part_is_empty = len(part) == 0
            if part_is_empty:
                continue

            is_image = part.startswith("<img ") and part.endswith(">")
            if is_image:
                part = part.replace(">", " style='max-width:100%; height:auto;'>")
                self.image_buffer = part
                continue

            is_numbered_list = re.search(r"^\d+\)\s", part) is not None
            is_title = i == 0
            is_heading = part.startswith("<strong>") and part.endswith("</strong>")
            if is_numbered_list or is_title or is_heading:
                self.flush()

            previous_ends_with_hyphen = (
                i > 0 and re.search(re_ends_with_hyphen, parts[i - 1]) is not None
            )
            if previous_ends_with_hyphen:
                self.paragraph[-1] = re.sub(
                    re_ends_with_hyphen, part, self.paragraph[-1]
                )
            else:
                self.paragraph.append(part)

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
            if "<img " in part:
                a = 1
        self.flush()
        result = "</p><p>".join(self.out)
        return result

    def flush(self):
        self.out.append("".join(self.paragraph).strip())
        self.paragraph = []
        if self.image_buffer is not None:
            self.out.append(self.image_buffer)
            self.image_buffer = None


if __name__ == "__main__":
    main()
