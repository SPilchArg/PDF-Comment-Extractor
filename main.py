import re
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment
from langdetect import detect, LangDetectException
from collections import Counter

ARGOS = "argos"

SKIP_DESCRIPTIONS = {
    "sticky note", "underline", "highlight", "strikeout", "squiggly",
    "ink", "line", "square", "circle", "polygon", "polyline", "stamp",
    "caret", "free text", "freetext", "popup", "file attachment", "sound",
    "movie", "widget", "screen", "printer mark", "trap net", "watermark",
    "3d", "redact",
}

def extract_comments_from_pdf(pdf_path: str,
                               progress_callback=None,
                               status_callback=None) -> list[dict]:
    comments    = []
    comment_id  = 1
    file_name   = Path(pdf_path).name
    skip_lower  = {s.lower() for s in SKIP_DESCRIPTIONS}
    doc         = fitz.open(pdf_path)
    total_pages = len(doc)

    for page_num in range(total_pages):
        page           = doc[page_num]
        full_page_text = page.get_text()

        if status_callback:
            status_callback(f"Scanning page {page_num + 1} of {total_pages} …")
        if progress_callback:
            progress_callback((page_num + 1) / total_pages * 100)

        for annot in page.annots():
            raw_text = _get_annotation_text(annot)
            if not raw_text:
                continue
            if raw_text.strip().lower() in skip_lower:
                continue

            parsed      = parse_annotation_text(raw_text)
            description = format_description(parsed)

            if _is_junk_annotation(description, skip_lower):
                continue

            keywords       = build_keywords(parsed)
            annotated_text = _get_annotated_text(annot, page, full_page_text)
            language       = _detect_language(annotated_text)

            comments.append({
                "no":            comment_id,
                "file_name":     file_name,
                "description":   description,
                "found_on_page": page_num + 1,
                "language":      language,
                "keywords":      keywords,
            })
            comment_id += 1

    doc.close()
    return comments


def _get_annotation_text(annot) -> str:
    content = (annot.info.get("content") or "").strip()
    if not content:
        try:
            content = (annot.get_text() or "").strip()
        except Exception:
            pass
    return content


def _get_annotated_text(annot, page, full_page_text: str) -> str:
    rect = annot.rect
    for pad in (200, 400, 800):
        try:
            clip = fitz.Rect(
                max(rect.x0 - pad, 0), max(rect.y0 - pad, 0),
                rect.x1 + pad,         rect.y1 + pad,
            )
            text = page.get_text("text", clip=clip).strip()
            if len(text) >= 80:
                return text
        except Exception:
            pass
    return full_page_text


def _is_junk_annotation(description: str, skip_lower: set) -> bool:
    return description.strip().lower() in skip_lower


def parse_annotation_text(text: str) -> dict:
    fields = {
        "author": "", "status": "", "description": "",
        "reason": "", "severity": "", "freetext": [], "extra": [],
    }
    known_keys = {
        "status":      r"^status\s*[:\-]\s*",
        "description": r"^description\s*[:\-]\s*",
        "reason":      r"^reason\s*[:\-]\s*",
        "severity":    r"^severity\s*[:\-]\s*",
    }

    lines          = [l.strip() for l in text.splitlines() if l.strip()]
    argos_stripped = False
    has_structured = False

    for line in lines:
        if not argos_stripped and line.lower() == ARGOS:
            fields["author"] = line
            argos_stripped   = True
            continue

        matched = False
        for field, pattern in known_keys.items():
            if re.match(pattern, line, re.IGNORECASE):
                fields[field] = re.sub(
                    pattern, "", line, flags=re.IGNORECASE).strip()
                matched        = True
                has_structured = True
                break

        if matched:
            continue

        if has_structured:
            fields["extra"].append(line)
        else:
            fields["freetext"].append(line)

    fields["_has_structured"] = has_structured
    return fields


def format_description(parsed: dict) -> str:
    has_structured = parsed.get("_has_structured", False)
    freetext       = parsed.get("freetext", [])
    extra          = parsed.get("extra",    [])

    if (not has_structured and not freetext
            and not parsed["status"] and not parsed["description"]
            and not parsed["reason"] and not parsed["severity"]):
        return "Strikethrough Text"

    if not has_structured and freetext:
        return "\n".join(freetext)

    lines = []
    if parsed["status"]:      lines.append(f"Status: {parsed['status']}")
    if parsed["description"]: lines.append(f"Description: {parsed['description']}")
    if parsed["reason"]:      lines.append(f"Reason: {parsed['reason']}")
    if parsed["severity"]:    lines.append(f"Severity: {parsed['severity']}")
    for ft in freetext:       lines.append(ft)
    for ex in extra:          lines.append(ex)

    return "\n".join(lines) if lines else "Strikethrough Text"


def build_keywords(parsed: dict) -> str:
    has_structured = parsed.get("_has_structured", False)
    freetext       = parsed.get("freetext", [])

    if (not has_structured and not freetext
            and not parsed["status"] and not parsed["description"]
            and not parsed["reason"] and not parsed["severity"]):
        return "N/A"

    keywords = []
    for word in parsed.get("severity", "").split():
        clean = re.sub(r'[^\w\-]', '', word)
        if clean and clean not in keywords:
            keywords.append(clean)

    all_text = " ".join([
        parsed.get("status", ""),      parsed.get("description", ""),
        parsed.get("reason", ""),      parsed.get("severity", ""),
        " ".join(parsed.get("freetext", [])),
    ])
    if re.search(r'\bDTP\b', all_text, re.IGNORECASE):
        if "DTP" not in keywords:
            keywords.insert(0, "DTP")

    return ", ".join(keywords) if keywords else "N/A"


def _detect_language(text: str) -> str:
    text = text.strip()
    if len(text) < 20:
        return "Unknown"
    try:
        return convert_to_bcp47(detect(text))
    except LangDetectException:
        return "Unknown"


def convert_to_bcp47(lang_code: str) -> str:
    lang_map = {
        "pl": "pl-PL", "en": "en-US", "de": "de-DE", "fr": "fr-FR",
        "es": "es-ES", "it": "it-IT", "pt": "pt-PT", "ru": "ru-RU",
        "nl": "nl-NL", "cs": "cs-CZ", "sk": "sk-SK", "hu": "hu-HU",
        "ro": "ro-RO", "bg": "bg-BG", "hr": "hr-HR", "sr": "sr-RS",
        "uk": "uk-UA", "sv": "sv-SE", "da": "da-DK", "fi": "fi-FI",
        "nb": "nb-NO", "tr": "tr-TR", "ar": "ar-SA",
        "zh-cn": "zh-CN", "zh-tw": "zh-TW",
        "ja": "ja-JP", "ko": "ko-KR",
    }
    return lang_map.get(lang_code.lower(), f"{lang_code}-{lang_code.upper()}")


def create_excel_report(comments: list[dict],
                         output_path: str,
                         status_callback=None) -> None:
    if status_callback:
        status_callback("Building Excel workbook …")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comments"
    wrap = Alignment(horizontal='left', vertical='top', wrap_text=True)

    headers = [
        ("No.",                  8),
        ("File Name",           30),
        ("Description/Comment", 60),
        ("Found on Page",       15),
        ("Language",            15),
        ("Keywords Appearing",  30),
    ]

    for col_idx, (label, width) in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=label)
        ws.column_dimensions[chr(64 + col_idx)].width = width

    ws.freeze_panes = 'A2'

    if not comments:
        ws.cell(row=2, column=1, value="No comments found in the PDF file.")
    else:
        for row_idx, comment in enumerate(comments, start=2):
            row_data = [
                comment["no"], comment["file_name"], comment["description"],
                comment["found_on_page"], comment["language"], comment["keywords"],
            ]
            for col_idx, value in enumerate(row_data, start=1):
                cell           = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.alignment = wrap
            line_count = comment["description"].count('\n') + 1
            ws.row_dimensions[row_idx].height = max(15, line_count * 15)

    if status_callback:
        status_callback("Building Summary sheet …")

    ws_sum = wb.create_sheet("Summary")
    create_summary_sheet(ws_sum, comments)

    if status_callback:
        status_callback("Saving file …")

    wb.save(output_path)


def create_summary_sheet(ws, comments: list[dict]) -> None:
    ws.cell(row=1, column=1, value="COMMENTS SUMMARY")
    ws.cell(row=3, column=1, value="Metric")
    ws.cell(row=3, column=2, value="Value")

    metrics = [
        ("Total Comments Found", len(comments)),
        ("Pages with Comments",
         len(set(c["found_on_page"] for c in comments)) if comments else 0),
        ("Languages Detected",
         len(set(c["language"]     for c in comments)) if comments else 0),
    ]
    for idx, (metric, value) in enumerate(metrics, start=4):
        ws.cell(row=idx, column=1, value=metric)
        ws.cell(row=idx, column=2, value=value)

    if not comments:
        return

    lang_ctr = Counter(c["language"] for c in comments)
    kw_ctr   = Counter(
        kw.strip()
        for c in comments
        for kw in c["keywords"].split(", ")
        if kw.strip() and kw.strip() != "N/A"
    )

    lang_start = 4 + len(metrics) + 1
    ws.cell(row=lang_start,     column=1, value="Language Breakdown")
    ws.cell(row=lang_start + 1, column=1, value="Language")
    ws.cell(row=lang_start + 1, column=2, value="Count")
    for idx, (lang, count) in enumerate(lang_ctr.most_common(),
                                        start=lang_start + 2):
        ws.cell(row=idx, column=1, value=lang)
        ws.cell(row=idx, column=2, value=count)

    kw_start = lang_start + 2 + len(lang_ctr) + 1
    ws.cell(row=kw_start,     column=1, value="Keywords Breakdown")
    ws.cell(row=kw_start + 1, column=1, value="Keyword")
    ws.cell(row=kw_start + 1, column=2, value="Occurrences")
    for idx, (kw, count) in enumerate(kw_ctr.most_common(),
                                      start=kw_start + 2):
        ws.cell(row=idx, column=1, value=kw)
        ws.cell(row=idx, column=2, value=count)



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Comment Extractor")
        self.resizable(False, False)
        self.configure(padx=24, pady=20, bg="#f5f5f5")

        self._pdf_path    = tk.StringVar()
        self._output_dir  = tk.StringVar()

        self._pdf_path.trace_add("write", self._on_pdf_changed)

        self._build_ui()


    def _build_ui(self):
        BG     = "#f5f5f5"
        ACCENT = "#2563eb"
        FONT   = ("Segoe UI", 10)
        FONT_B = ("Segoe UI", 10, "bold")
        FONT_H = ("Segoe UI", 14, "bold")

        tk.Label(self, text="PDF Comment Extractor",
                 font=FONT_H, bg=BG, fg="#1e293b").grid(
            row=0, column=0, columnspan=3, pady=(0, 18), sticky="w")

        tk.Label(self, text="Input PDF", font=FONT_B,
                 bg=BG, fg="#334155").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 4))

        self._pdf_name_var = tk.StringVar(value="No file selected")
        tk.Label(self, textvariable=self._pdf_name_var,
                 font=("Segoe UI", 9), bg="#e2e8f0", fg="#1e293b",
                 anchor="w", relief="solid", bd=1, padx=6, pady=4,
                 width=46).grid(
            row=2, column=0, columnspan=2, sticky="ew")

        tk.Button(self, text="Browse …", font=FONT,
                  bg="#e2e8f0", fg="#1e293b", relief="flat",
                  activebackground="#cbd5e1", cursor="hand2",
                  command=self._browse_pdf).grid(
            row=2, column=2, padx=(8, 0), ipady=4, ipadx=6)

        tk.Label(self, text="Output Folder", font=FONT_B,
                 bg=BG, fg="#334155").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(14, 4))

        self._dir_name_var = tk.StringVar(value="Same folder as input PDF")
        tk.Label(self, textvariable=self._dir_name_var,
                 font=("Segoe UI", 9), bg="#e2e8f0", fg="#1e293b",
                 anchor="w", relief="solid", bd=1, padx=6, pady=4,
                 width=46).grid(
            row=4, column=0, columnspan=2, sticky="ew")

        tk.Button(self, text="Browse …", font=FONT,
                  bg="#e2e8f0", fg="#1e293b", relief="flat",
                  activebackground="#cbd5e1", cursor="hand2",
                  command=self._browse_output_dir).grid(
            row=4, column=2, padx=(8, 0), ipady=4, ipadx=6)

        self._filename_hint_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._filename_hint_var,
                 font=("Segoe UI", 8), bg=BG, fg="#94a3b8",
                 anchor="w").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(2, 0))

        tk.Label(self, text="Progress", font=FONT_B,
                 bg=BG, fg="#334155").grid(
            row=6, column=0, sticky="w", pady=(18, 4))

        self._progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(self, variable=self._progress_var,
                        maximum=100, length=460,
                        mode="determinate").grid(
            row=7, column=0, columnspan=3, sticky="ew")

        self._status_var = tk.StringVar(value="Ready.")
        tk.Label(self, textvariable=self._status_var,
                 font=("Segoe UI", 9), bg=BG, fg="#64748b",
                 anchor="w", width=60).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

        tk.Label(self, text="Log", font=FONT_B,
                 bg=BG, fg="#334155").grid(
            row=9, column=0, sticky="w", pady=(14, 4))

        log_frame = tk.Frame(self, bg=BG)
        log_frame.grid(row=10, column=0, columnspan=3, sticky="nsew")

        self._log = tk.Text(log_frame, height=10, width=62,
                            font=("Consolas", 9),
                            bg="#1e293b", fg="#e2e8f0",
                            relief="flat", state="disabled",
                            wrap="word", padx=8, pady=6)
        scrollbar = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=scrollbar.set)
        self._log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._run_btn = tk.Button(
            self, text="▶  Run Extraction",
            font=("Segoe UI", 11, "bold"),
            bg=ACCENT, fg="white",
            activebackground="#1d4ed8", activeforeground="white",
            relief="flat", cursor="hand2",
            command=self._start_extraction,
            pady=8)
        self._run_btn.grid(
            row=11, column=0, columnspan=3,
            sticky="ew", pady=(18, 0))


    def _on_pdf_changed(self, *_):
        pdf = self._pdf_path.get()
        if pdf:
            self._pdf_name_var.set(Path(pdf).name)
            self._update_filename_hint()
        else:
            self._pdf_name_var.set("No file selected")
            self._filename_hint_var.set("")

    def _update_filename_hint(self):
        pdf = self._pdf_path.get()
        if not pdf:
            self._filename_hint_var.set("")
            return
        out_name = Path(pdf).stem + ".xlsx"
        folder   = self._output_dir.get()
        if folder:
            folder_label = Path(folder).name or folder
        else:
            folder_label = Path(pdf).parent.name or "same folder"
        self._filename_hint_var.set(f"→  {out_name}  (in '{folder_label}')")

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="Select input PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self._pdf_path.set(path)
            # Default output dir = same folder as the PDF (only if not set yet)
            if not self._output_dir.get():
                self._dir_name_var.set("Same folder as input PDF")

    def _browse_output_dir(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self._output_dir.set(folder)
            self._dir_name_var.set(Path(folder).name or folder)
            self._update_filename_hint()

    def _log_line(self, message: str):
        self._log.configure(state="normal")
        self._log.insert("end", message + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _set_status(self, message: str):
        self._status_var.set(message)
        self._log_line(message)
        self.update_idletasks()

    def _set_progress(self, value: float):
        self._progress_var.set(value)
        self.update_idletasks()

    def _resolve_output_path(self) -> str:
        pdf        = self._pdf_path.get().strip()
        out_name   = Path(pdf).stem + ".xlsx"
        folder     = self._output_dir.get().strip()
        out_folder = Path(folder) if folder else Path(pdf).parent
        return str(out_folder / out_name)

    def _start_extraction(self):
        pdf_path = self._pdf_path.get().strip()

        if not pdf_path:
            messagebox.showwarning("Missing input",
                                   "Please select an input PDF file.")
            return
        if not Path(pdf_path).exists():
            messagebox.showerror("File not found",
                                 f"Cannot find:\n{pdf_path}")
            return

        output_path = self._resolve_output_path()

        self._progress_var.set(0)
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        self._run_btn.configure(state="disabled", text="Running …")
        self._set_status("Starting …")

        thread = threading.Thread(
            target=self._run_extraction,
            args=(pdf_path, output_path),
            daemon=True)
        thread.start()

    def _run_extraction(self, pdf_path: str, output_path: str):
        try:
            self._set_status(f"Opening: {Path(pdf_path).name}")

            comments = extract_comments_from_pdf(
                pdf_path,
                progress_callback=self._set_progress,
                status_callback=self._set_status,
            )

            self._set_status(
                f"Extraction complete — {len(comments)} comment(s) found.")
            self._set_progress(100)

            create_excel_report(
                comments,
                output_path,
                status_callback=self._set_status,
            )

            self._set_progress(100)
            self._set_status(f"Saved: {Path(output_path).name}")
            self._log_line("─" * 50)
            self._log_line(f"✓ Done!  {len(comments)} comment(s) written.")

            messagebox.showinfo(
                "Complete",
                f"{len(comments)} comment(s) extracted.\n\n"
                f"File:  {Path(output_path).name}\n"
                f"Folder:  {Path(output_path).parent}")

        except Exception as exc:
            self._set_status(f"ERROR: {exc}")
            messagebox.showerror("Error", str(exc))

        finally:
            self.after(0, self._reset_run_button)

    def _reset_run_button(self):
        self._run_btn.configure(state="normal", text="▶  Run Extraction")

if __name__ == "__main__":
    app = App()
    app.mainloop()