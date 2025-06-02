import os
import tempfile
import time
from docx import Document
import re


class DocxReplacer:
    def __init__(self, input_file, output_file, replacements):
        self.input_file = input_file
        self.output_file = output_file
        self.replacements = replacements
        self.doc = Document(input_file)

    def _apply_replacement(self, text, replacements):
        for key, val in replacements.items():
            if key in text:
                text = text.replace(key, val)
        return text

    def _replace_in_paragraphs(self, final_replacements):
        for paragraph in self.doc.paragraphs:
            full_text = "".join(run.text for run in paragraph.runs)
            replaced_text = self._apply_replacement(full_text, final_replacements)
            if replaced_text != full_text:
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = replaced_text
                else:
                    paragraph.add_run(replaced_text)

    def _replace_in_tables(self, final_replacements):
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        full_text = "".join(run.text for run in paragraph.runs)
                        replaced_text = self._apply_replacement(full_text, final_replacements)
                        if replaced_text != full_text:
                            for run in paragraph.runs:
                                run.text = ""
                            if paragraph.runs:
                                paragraph.runs[0].text = replaced_text
                            else:
                                paragraph.add_run(replaced_text)

    def _replace_in_textboxes(self, final_replacements):
        for shape in self.doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent'):
            for text_elem in shape.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                if text_elem.text:
                    text_elem.text = self._apply_replacement(text_elem.text, final_replacements)

    def _generate_final_replacements(self):
        final_replacements = {}
        for item in self.replacements:
            key = item.get("key")
            value = item.get("value")
            type_ = item.get("type")

            if type_ == "string":
                final_replacements[key] = value
            elif type_ == "full_name":
                parts = value.strip().split(maxsplit=1)
                first_name = parts[0]
                surname = parts[1] if len(parts) > 1 else ''
                final_replacements["<first_name>"] = first_name
                final_replacements["<surname>"] = surname
            elif type_ == "id":
                if len(value) == 9:
                    for i, digit in enumerate(value):
                        final_replacements[f"<id_{i+1}>"] = digit
        return final_replacements

    def _clear_unreplaced_placeholders(self):
        placeholder_pattern = r"<[^>]+>"

        for paragraph in self.doc.paragraphs:
            full_text = "".join(run.text for run in paragraph.runs)
            cleaned_text = re.sub(placeholder_pattern, "", full_text)
            if cleaned_text != full_text:
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = cleaned_text
                else:
                    paragraph.add_run(cleaned_text)

        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        full_text = "".join(run.text for run in paragraph.runs)
                        cleaned_text = re.sub(placeholder_pattern, "", full_text)
                        if cleaned_text != full_text:
                            for run in paragraph.runs:
                                run.text = ""
                            if paragraph.runs:
                                paragraph.runs[0].text = cleaned_text
                            else:
                                paragraph.add_run(cleaned_text)

        for shape in self.doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent'):
            for text_elem in shape.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                if text_elem.text:
                    text_elem.text = re.sub(placeholder_pattern, "", text_elem.text)

    def run(self):
        final_replacements = self._generate_final_replacements()
        self._replace_in_paragraphs(final_replacements)
        self._replace_in_tables(final_replacements)
        self._replace_in_textboxes(final_replacements)
        self._clear_unreplaced_placeholders()
        self.doc.save(self.output_file)


# Example usage
if __name__ == "__main__":
    input_file = r"C:\word\bekoretKatsen.docx"
    output_file = rf"C:\word\bekoretKatsen_{time.time()}.docx"
    replacements = [
        {"key": "<cusName>", "value": "פיראס סרור", "type": "string"},
        {"key": "<cusPhone>", "value": "0543341222", "type": "string"},
    ]

    replacer = DocxReplacer(input_file, output_file, replacements)
    replacer.run()