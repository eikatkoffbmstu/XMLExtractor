import os
import re
import zipfile
import threading
from pathlib import Path
from tkinter import Tk, Button, Label, Entry, StringVar, filedialog, messagebox, ttk

import xml.etree.ElementTree as ET



DROP_TAG_SUFFIXES = {
    "created", "modified", "lastPrinted", "revision", "totalTime",
    "creator", "lastModifiedBy", "keywords", "description", "subject", "title",
    "application", "appVersion", "company", "manager",
}


DROP_ATTR_NAMES = {
    "id", "Id", "ID",
    "rsid", "rsidR", "rsidRDefault", "rsidP", "rsidRPr",
    "paraId", "textId",
}


DROP_ATTR_PATTERNS = [
    re.compile(r".*rsid.*", re.IGNORECASE),
    re.compile(r".*paraId.*", re.IGNORECASE),
    re.compile(r".*textId.*", re.IGNORECASE),
]


def _local_name(tag: str) -> str:
    """{namespace}local -> local"""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _should_drop_tag(tag: str) -> bool:
    return _local_name(tag) in DROP_TAG_SUFFIXES


def _clean_element(elem: ET.Element):

    to_del = []
    for k in elem.attrib.keys():
        lk = _local_name(k)
        if lk in DROP_ATTR_NAMES:
            to_del.append(k)
            continue
        for pat in DROP_ATTR_PATTERNS:
            if pat.match(lk):
                to_del.append(k)
                break
    for k in to_del:
        elem.attrib.pop(k, None)


    children = list(elem)
    for ch in children:
        if _should_drop_tag(ch.tag):
            elem.remove(ch)
        else:
            _clean_element(ch)


    if elem.text is not None:
        elem.text = elem.text.replace("\r\n", "\n")
    if elem.tail is not None:
        elem.tail = elem.tail.replace("\r\n", "\n")


def _sort_attribs(elem: ET.Element):

    if elem.attrib:
        items = sorted(elem.attrib.items(), key=lambda kv: kv[0])
        elem.attrib.clear()
        elem.attrib.update(items)
    for ch in list(elem):
        _sort_attribs(ch)


def normalize_xml_bytes(xml_bytes: bytes) -> bytes:
    """
    Превращает XML в более "стабильный":
    - удаляет шумовые узлы/атрибуты
    - сортирует атрибуты
    - записывает без лишних пробелов (ElementTree сам по себе не pretty-print)
    """

    xml_bytes = xml_bytes.lstrip()

    root = ET.fromstring(xml_bytes)
    _clean_element(root)
    _sort_attribs(root)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)



def ensure_outdir(base_outdir: Path, input_file: Path) -> Path:
    base_outdir.mkdir(parents=True, exist_ok=True)
    sub = base_outdir / (input_file.stem + "_xml")
    sub.mkdir(parents=True, exist_ok=True)
    return sub


def write_file(path: Path, data: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def extract_from_zip_office(input_file: Path, outdir: Path):
    """
    Для DOCX/XLSX/ODT/ODS:
    - достаём "контентные" XML
    - исключаем метаданные
    - нормализуем и сохраняем
    """
    suffix = input_file.suffix.lower()
    with zipfile.ZipFile(input_file, "r") as z:
        names = z.namelist()

        want = []

        if suffix == ".docx":

            want += [
                "word/document.xml",
                "word/styles.xml",
                "word/numbering.xml",
                "word/footnotes.xml",
                "word/endnotes.xml",
                "word/settings.xml",
            ]
            want += [
                "word/comments.xml",
                "word/header1.xml", "word/header2.xml", "word/header3.xml",
                "word/footer1.xml", "word/footer2.xml", "word/footer3.xml",
            ]

        elif suffix == ".xlsx":

            want += ["xl/workbook.xml", "xl/sharedStrings.xml", "xl/styles.xml"]
            # Все листы
            want += [n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")]
            want += [n for n in names if n.startswith("xl/tables/") and n.endswith(".xml")]
        elif suffix in (".odt", ".ods"):

            want += ["content.xml", "styles.xml", "settings.xml"]

        else:
            raise ValueError(f"Неподдерживаемый zip-офисный формат: {suffix}")
        seen = set()
        want_existing = []
        for n in want:
            if n in names and n not in seen:
                want_existing.append(n)
                seen.add(n)

        if not want_existing:
            raise ValueError("Не удалось найти значимые XML внутри файла (структура неожиданная).")

        saved = []
        for member in want_existing:
            raw = z.read(member)
            try:
                norm = normalize_xml_bytes(raw)
            except Exception as e:
                norm = raw
                member_out = outdir / (member.replace("/", "__") + ".RAW.xml")
                write_file(member_out, norm)
                saved.append(member_out.name)
                continue

            member_out = outdir / (member.replace("/", "__"))
            write_file(member_out, norm)
            saved.append(member_out.name)

        return saved


def extract_from_pdf(input_file: Path, outdir: Path):
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTTextContainer, LTTextLine
    except Exception:
        raise RuntimeError("Для PDF нужен пакет pdfminer.six: pip install pdfminer.six")

    root = ET.Element("pdf")
    root.set("source", input_file.name)

    page_index = 0
    for page_layout in extract_pages(str(input_file)):
        page_el = ET.SubElement(root, "page", index=str(page_index))
        lines = []

        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for obj in element:
                    if isinstance(obj, LTTextLine):
                        txt = obj.get_text()
                        txt = txt.replace("\r\n", "\n").strip("\n")
                        if txt:
                            for one_line in txt.split("\n"):
                                cleaned = re.sub(r"[ \t]+", " ", one_line).strip()
                                if cleaned:
                                    lines.append(cleaned)

        for ln in lines:
            ln_el = ET.SubElement(page_el, "line")
            ln_el.text = ln

        page_index += 1

    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    out = outdir / "pdf__extracted_text.xml"
    write_file(out, xml_bytes)
    return [out.name]


def extract_xmls(input_file: Path, base_outdir: Path):
    outdir = ensure_outdir(base_outdir, input_file)
    suffix = input_file.suffix.lower()

    if suffix in (".docx", ".xlsx", ".odt", ".ods"):
        saved = extract_from_zip_office(input_file, outdir)
    elif suffix == ".pdf":
        saved = extract_from_pdf(input_file, outdir)
    elif suffix in (".csv", ".txt"):
        raise ValueError("Для CSV/TXT экстракция XML не предусмотрена (как вы и говорили).")
    else:
        raise ValueError(f"Формат {suffix} пока не поддерживается.")

    return outdir, saved



class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("XML Report Extractor (для сравнения)")

        self.file_var = StringVar()
        self.out_var = StringVar(value=str(Path.cwd() / "extracted"))

        Label(root, text="Файл отчёта:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        Entry(root, textvariable=self.file_var, width=60).grid(row=0, column=1, padx=8, pady=6)
        Button(root, text="Выбрать...", command=self.pick_file).grid(row=0, column=2, padx=8, pady=6)

        Label(root, text="Папка вывода:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        Entry(root, textvariable=self.out_var, width=60).grid(row=1, column=1, padx=8, pady=6)
        Button(root, text="Выбрать...", command=self.pick_outdir).grid(row=1, column=2, padx=8, pady=6)

        self.run_btn = Button(root, text="Извлечь XML", command=self.run)
        self.run_btn.grid(row=2, column=1, pady=10)

        self.status = StringVar(value="Готово.")
        Label(root, textvariable=self.status).grid(row=3, column=0, columnspan=3, sticky="w", padx=8, pady=6)

        self.listbox = ttk.Treeview(root, columns=("name",), show="headings", height=10)
        self.listbox.heading("name", text="Сохранённые XML")
        self.listbox.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)

        root.grid_rowconfigure(4, weight=1)
        root.grid_columnconfigure(1, weight=1)

    def pick_file(self):
        path = filedialog.askopenfilename(
            title="Выберите отчёт",
            filetypes=[
                ("Reports", "*.docx *.pdf *.odt *.ods *.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.file_var.set(path)

    def pick_outdir(self):
        path = filedialog.askdirectory(title="Выберите папку вывода")
        if path:
            self.out_var.set(path)

    def run(self):
        file_path = self.file_var.get().strip()
        out_path = self.out_var.get().strip()

        if not file_path:
            messagebox.showerror("Ошибка", "Выберите файл отчёта.")
            return
        if not out_path:
            messagebox.showerror("Ошибка", "Выберите папку вывода.")
            return

        input_file = Path(file_path)
        base_outdir = Path(out_path)

        if not input_file.exists():
            messagebox.showerror("Ошибка", "Файл не найден.")
            return

        self.run_btn.config(state="disabled")
        self.status.set("Извлекаю...")
        t = threading.Thread(target=self._run_worker, args=(input_file, base_outdir), daemon=True)
        t.start()

    def _run_worker(self, input_file: Path, base_outdir: Path):
        try:
            outdir, saved = extract_xmls(input_file, base_outdir)
        except Exception as e:
            self.root.after(0, lambda: self._done_error(str(e)))
            return

        self.root.after(0, lambda: self._done_ok(outdir, saved))

    def _done_ok(self, outdir: Path, saved):
        for item in self.listbox.get_children():
            self.listbox.delete(item)
        for name in saved:
            self.listbox.insert("", "end", values=(name,))
        self.status.set(f"Готово. Папка: {outdir}")
        self.run_btn.config(state="normal")
        messagebox.showinfo("Готово", f"Сохранено XML: {len(saved)}\n\n{outdir}")

    def _done_error(self, msg: str):
        self.status.set("Ошибка.")
        self.run_btn.config(state="normal")
        messagebox.showerror("Ошибка", msg)


def main():
    root = Tk()
    app = App(root)
    root.minsize(760, 420)
    root.mainloop()

if __name__ == "__main__":
    main()
