#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui_app.py
==========
โปรแกรมตรวจอักษรไทย-อังกฤษสำหรับไฟล์ Word (.docx) — หน้าจอสำหรับพนักงานทั่วไป

ใช้งานง่าย 3 ขั้นตอน: เลือกไฟล์ -> กด "เริ่มตรวจสอบ" -> เปิดผลลัพธ์
ทำงานอยู่เบื้องหลังด้วย spellcheck_engine.py (เอนจินตรวจคำตัวเดียวกับที่ใช้ผ่าน
บรรทัดคำสั่งได้ด้วย) รันแบบออฟไลน์ล้วน ไม่ต้องต่ออินเทอร์เน็ต
"""

import os
import queue
import sys
import threading
import traceback
import webbrowser
from tkinter import (
    Tk, Toplevel, StringVar, BooleanVar, IntVar, PhotoImage,
    filedialog, messagebox, END, DISABLED, NORMAL, WORD,
)
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

import spellcheck_engine as engine

APP_TITLE = "โปรแกรมตรวจอักษรไทย-อังกฤษ (Word)"
APP_FONT_FAMILY = "Tahoma"  # ฟอนต์นี้มีมากับ Windows ทุกเครื่อง แสดงผลภาษาไทยได้ดี

LIMITATIONS_TEXT = """ข้อจำกัดของโปรแกรมที่ควรทราบ

- โปรแกรมนี้เป็น "ตัวช่วยกรอง" คำที่มีแนวโน้มสะกดผิด ไม่ใช่เครื่องมือที่แม่นยำ
  100% ยังคงต้องมีคนอ่านทวนเอกสารสำคัญอีกครั้งก่อนส่งเสมอ

- การตัดคำภาษาไทยมีความคลุมเครือโดยธรรมชาติ (ภาษาไทยไม่มีช่องว่างระหว่างคำ)
  ถ้าคำที่พิมพ์ผิดบังเอิญถูกตัดเป็นคำย่อยที่ถูกต้องพอดีทั้งคู่ เช่น พิมพ์
  "อนุมัด" ผิด แล้วระบบตัดเป็น "อนุ" + "มัด" ซึ่งทั้งสองคำนี้มีอยู่จริงใน
  พจนานุกรม โปรแกรมจะมองไม่เห็นคำผิดนั้น

- ไม่รองรับข้อความในกล่องข้อความ (text box), SmartArt หรือวัตถุฝังอื่น ๆ

- เชิงอรรถ/อ้างอิงท้ายเรื่อง จะถูกตรวจและแสดงในรายงาน Excel เท่านั้น
  (ไม่สามารถไฮไลต์ย้อนกลับเข้าไฟล์ Word ได้)

- คำในหัวกระดาษ/ท้ายกระดาษ จะถูกไฮไลต์สีให้ แต่จะไม่มีคอมเมนต์คำแนะนำกำกับ
  (ข้อจำกัดของไฟล์ Word) ให้ดูคำแนะนำในรายงาน Excel แทน

- ชื่อเฉพาะ/ชื่อบริษัท/ศัพท์เทคนิคที่ไม่อยู่ในพจนานุกรม จะถูกตีธงว่าน่าสงสัย
  ได้เสมอ ถ้าเอกสารของหน่วยงานมีคำเหล่านี้ซ้ำ ๆ ให้เพิ่มลงในพจนานุกรมเสริม
  (ปุ่ม "พจนานุกรมเสริม (ถ้ามี)" ในหน้าหลัก) จะได้ไม่ถูกแจ้งซ้ำอีก
"""


class QueueWriter:
    """เปลี่ยนข้อความสถานะจาก engine ให้ไหลเข้าคิว เพื่อให้ thread เบื้องหลังส่ง
    ข้อความมาอัปเดตหน้าจอหลัก (tkinter) ได้อย่างปลอดภัย"""

    def __init__(self, q: "queue.Queue"):
        self.q = q

    def __call__(self, msg: str):
        self.q.put(("log", msg))


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("760x780")
        root.minsize(700, 640)

        self.default_font = (APP_FONT_FAMILY, 10)
        self.bold_font = (APP_FONT_FAMILY, 11, "bold")
        self.header_font = (APP_FONT_FAMILY, 15, "bold")

        style = ttk.Style()
        try:
            style.theme_use("vista")  # ให้หน้าตาดูใกล้เคียงโปรแกรม Windows ทั่วไป
        except Exception:
            pass
        style.configure("TButton", font=self.default_font, padding=6)
        style.configure("Big.TButton", font=self.bold_font, padding=10)
        style.configure("TLabel", font=self.default_font)
        style.configure("TCheckbutton", font=self.default_font)

        self.input_path = StringVar()
        self.extra_dict_path = StringVar()
        self.output_dir = StringVar()
        self.same_folder_as_input = BooleanVar(value=True)
        self.check_thai = BooleanVar(value=True)
        self.check_english = BooleanVar(value=True)
        self.make_docx = BooleanVar(value=True)

        self.result_docx_path = None
        self.result_excel_path = None
        self.result_out_dir = None

        self.msg_queue: "queue.Queue" = queue.Queue()
        self.worker_thread = None

        self._build_ui()
        self.root.after(120, self._poll_queue)

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        pad = {"padx": 14, "pady": 6}

        header = ttk.Frame(self.root)
        header.pack(fill="x", **pad)
        ttk.Label(header, text=APP_TITLE, font=self.header_font).pack(side="left")
        ttk.Button(header, text="ℹ️ ข้อจำกัด/วิธีใช้", command=self._show_about).pack(side="right")

        # ---------- เลือกไฟล์ ----------
        file_frame = ttk.LabelFrame(self.root, text="1) เลือกไฟล์ Word ที่ต้องการตรวจ")
        file_frame.pack(fill="x", **pad)
        row = ttk.Frame(file_frame)
        row.pack(fill="x", padx=10, pady=10)
        entry = ttk.Entry(row, textvariable=self.input_path, font=self.default_font, state="readonly")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="เลือกไฟล์ .docx ...", command=self._choose_input).pack(side="left")

        # ---------- ตัวเลือก ----------
        opt_frame = ttk.LabelFrame(self.root, text="2) ตัวเลือก (ปกติไม่ต้องแก้อะไร)")
        opt_frame.pack(fill="x", **pad)

        lang_row = ttk.Frame(opt_frame)
        lang_row.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(lang_row, text="ตรวจภาษา:").pack(side="left", padx=(0, 10))
        ttk.Checkbutton(lang_row, text="ไทย", variable=self.check_thai).pack(side="left", padx=6)
        ttk.Checkbutton(lang_row, text="อังกฤษ", variable=self.check_english).pack(side="left", padx=6)
        ttk.Checkbutton(
            lang_row, text="สร้างไฟล์ Word ไฮไลต์ (ไม่ติ๊ก = เอาแค่รายงาน Excel)",
            variable=self.make_docx,
        ).pack(side="left", padx=18)

        dict_row = ttk.Frame(opt_frame)
        dict_row.pack(fill="x", padx=10, pady=4)
        ttk.Label(dict_row, text="พจนานุกรมเสริม (ถ้ามี):").pack(side="left", padx=(0, 10))
        ttk.Entry(dict_row, textvariable=self.extra_dict_path, font=self.default_font, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(dict_row, text="เลือกไฟล์ ...", command=self._choose_extra_dict).pack(side="left")
        ttk.Button(dict_row, text="ล้าง", command=lambda: self.extra_dict_path.set("")).pack(side="left", padx=(6, 0))

        out_row = ttk.Frame(opt_frame)
        out_row.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Checkbutton(
            out_row, text="บันทึกผลลัพธ์ไว้โฟลเดอร์เดียวกับไฟล์ต้นฉบับ",
            variable=self.same_folder_as_input, command=self._toggle_output_dir,
        ).pack(side="left")

        self.out_dir_row = ttk.Frame(opt_frame)
        self.out_dir_row.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Label(self.out_dir_row, text="โฟลเดอร์ผลลัพธ์:").pack(side="left", padx=(0, 10))
        self.out_dir_entry = ttk.Entry(
            self.out_dir_row, textvariable=self.output_dir, font=self.default_font, state="readonly"
        )
        self.out_dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.out_dir_btn = ttk.Button(self.out_dir_row, text="เลือกโฟลเดอร์ ...", command=self._choose_output_dir)
        self.out_dir_btn.pack(side="left")
        self._toggle_output_dir()

        # ---------- ปุ่มเริ่ม ----------
        action_frame = ttk.Frame(self.root)
        action_frame.pack(fill="x", **pad)
        self.start_btn = ttk.Button(
            action_frame, text="▶  เริ่มตรวจสอบ", style="Big.TButton", command=self._start_check
        )
        self.start_btn.pack(fill="x")

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=14, pady=(0, 6))

        # ---------- log ----------
        log_frame = ttk.LabelFrame(self.root, text="สถานะการทำงาน")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log_box = ScrolledText(log_frame, height=6, font=(APP_FONT_FAMILY, 9), wrap=WORD, state=DISABLED)
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)

        # ---------- ผลลัพธ์ ----------
        self.result_frame = ttk.LabelFrame(self.root, text="3) ผลลัพธ์")
        self.result_frame.pack(fill="x", **pad)
        self.result_label = ttk.Label(self.result_frame, text="ยังไม่ได้ตรวจสอบ", font=self.bold_font)
        self.result_label.pack(anchor="w", padx=10, pady=(10, 6))

        btn_row = ttk.Frame(self.result_frame)
        btn_row.pack(fill="x", padx=10, pady=(0, 10))
        self.open_docx_btn = ttk.Button(
            btn_row, text="📄 เปิดไฟล์ Word ที่ตรวจแล้ว", command=self._open_result_docx, state=DISABLED
        )
        self.open_docx_btn.pack(side="left", padx=(0, 8))
        self.open_excel_btn = ttk.Button(
            btn_row, text="📊 เปิดรายงาน Excel", command=self._open_result_excel, state=DISABLED
        )
        self.open_excel_btn.pack(side="left", padx=(0, 8))
        self.open_folder_btn = ttk.Button(
            btn_row, text="📁 เปิดโฟลเดอร์ผลลัพธ์", command=self._open_result_folder, state=DISABLED
        )
        self.open_folder_btn.pack(side="left")

    # ------------------------------------------------------------ helpers

    def _toggle_output_dir(self):
        state = DISABLED if self.same_folder_as_input.get() else NORMAL
        for w in (self.out_dir_entry, self.out_dir_btn):
            w.configure(state=state if w is self.out_dir_btn else "readonly")
        if self.same_folder_as_input.get():
            self.output_dir.set("")

    def _choose_input(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ Word",
            filetypes=[("Word Document", "*.docx"), ("ไฟล์ทั้งหมด", "*.*")],
        )
        if path:
            self.input_path.set(path)
            self.result_label.configure(text="ยังไม่ได้ตรวจสอบ")
            for b in (self.open_docx_btn, self.open_excel_btn, self.open_folder_btn):
                b.configure(state=DISABLED)

    def _choose_extra_dict(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์พจนานุกรมเสริม (.txt)",
            filetypes=[("Text file", "*.txt"), ("ไฟล์ทั้งหมด", "*.*")],
        )
        if path:
            self.extra_dict_path.set(path)

    def _choose_output_dir(self):
        path = filedialog.askdirectory(title="เลือกโฟลเดอร์สำหรับบันทึกผลลัพธ์")
        if path:
            self.output_dir.set(path)

    def _show_about(self):
        win = Toplevel(self.root)
        win.title("ข้อจำกัด / วิธีใช้งาน")
        win.geometry("560x520")
        box = ScrolledText(win, font=(APP_FONT_FAMILY, 10), wrap=WORD)
        box.pack(fill="both", expand=True, padx=10, pady=10)
        box.insert("1.0", LIMITATIONS_TEXT)
        box.configure(state=DISABLED)
        ttk.Button(win, text="ปิด", command=win.destroy).pack(pady=(0, 10))

    def _log(self, msg: str):
        self.log_box.configure(state=NORMAL)
        self.log_box.insert(END, msg + "\n")
        self.log_box.see(END)
        self.log_box.configure(state=DISABLED)

    def _set_busy(self, busy: bool):
        self.start_btn.configure(state=DISABLED if busy else NORMAL)
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    # ------------------------------------------------------------ actions

    def _start_check(self):
        input_path = self.input_path.get().strip()
        if not input_path:
            messagebox.showwarning(APP_TITLE, "กรุณาเลือกไฟล์ Word (.docx) ก่อน")
            return
        if not os.path.exists(input_path):
            messagebox.showerror(APP_TITLE, f"ไม่พบไฟล์:\n{input_path}")
            return
        if not (self.check_thai.get() or self.check_english.get()):
            messagebox.showwarning(APP_TITLE, "กรุณาเลือกอย่างน้อย 1 ภาษาที่จะตรวจ (ไทย หรือ อังกฤษ)")
            return

        extra_dict = self.extra_dict_path.get().strip() or None
        if self.same_folder_as_input.get():
            out_dir = os.path.dirname(os.path.abspath(input_path))
        else:
            out_dir = self.output_dir.get().strip()
            if not out_dir:
                messagebox.showwarning(APP_TITLE, "กรุณาเลือกโฟลเดอร์ผลลัพธ์ หรือติ๊กช่อง \"บันทึกในโฟลเดอร์เดียวกับไฟล์ต้นฉบับ\"")
                return

        self.log_box.configure(state=NORMAL)
        self.log_box.delete("1.0", END)
        self.log_box.configure(state=DISABLED)
        self.result_label.configure(text="กำลังตรวจสอบ กรุณารอสักครู่ ...")
        for b in (self.open_docx_btn, self.open_excel_btn, self.open_folder_btn):
            b.configure(state=DISABLED)
        self._set_busy(True)

        self.worker_thread = threading.Thread(
            target=self._run_worker,
            args=(input_path, extra_dict, out_dir, self.check_thai.get(), self.check_english.get(), self.make_docx.get()),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_worker(self, input_path, extra_dict, out_dir, check_thai, check_english, make_docx):
        reporter = QueueWriter(self.msg_queue)
        try:
            dicts = engine.load_dictionaries(extra_dict, verbose=False, on_progress=reporter)
            doc, findings = engine.analyze_document(
                input_path, dicts, check_thai=check_thai, check_english=check_english,
                verbose=False, on_progress=reporter,
            )

            base = os.path.splitext(os.path.basename(input_path))[0]
            os.makedirs(out_dir, exist_ok=True)

            excel_path = os.path.join(out_dir, f"{base}_รายงานตรวจคำ.xlsx")
            reporter("กำลังสร้างรายงาน Excel ...")
            engine.build_excel_report(findings, excel_path, os.path.basename(input_path))

            docx_out_path = None
            if make_docx:
                reporter("กำลังสร้างไฟล์ Word ไฮไลต์ ...")
                applied, skipped, hf_skipped = engine.apply_highlights_and_comments(doc, findings)
                docx_out_path = os.path.join(out_dir, f"{base}_ตรวจแล้ว.docx")
                doc.save(docx_out_path)
                reporter(f"ไฮไลต์แล้ว {applied} จุด (ข้าม {skipped} จุดที่ซ้อนทับ/ไฮไลต์ไม่ได้)")

            high = sum(1 for f in findings if f.confidence == engine.CONFIDENCE_HIGH)
            med = sum(1 for f in findings if f.confidence == engine.CONFIDENCE_MED)
            low = sum(1 for f in findings if f.confidence == engine.CONFIDENCE_LOW)

            self.msg_queue.put((
                "done",
                {
                    "total": len(findings), "high": high, "med": med, "low": low,
                    "docx": docx_out_path, "excel": excel_path, "out_dir": out_dir,
                },
            ))
        except Exception as e:
            tb = traceback.format_exc()
            self.msg_queue.put(("error", f"{e}\n\n{tb}"))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "done":
                    self._on_done(payload)
                elif kind == "error":
                    self._on_error(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    def _on_done(self, info):
        self._set_busy(False)
        self.result_docx_path = info["docx"]
        self.result_excel_path = info["excel"]
        self.result_out_dir = info["out_dir"]

        self.result_label.configure(
            text=(
                f"พบทั้งหมด {info['total']} จุด  —  "
                f"ความเชื่อมั่นสูง {info['high']} / กลาง {info['med']} / ต่ำ {info['low']}"
            )
        )
        if self.result_docx_path:
            self.open_docx_btn.configure(state=NORMAL)
        self.open_excel_btn.configure(state=NORMAL)
        self.open_folder_btn.configure(state=NORMAL)
        self._log("เสร็จสิ้น ✓")
        messagebox.showinfo(APP_TITLE, "ตรวจสอบเสร็จแล้ว ดูผลลัพธ์และเปิดไฟล์ได้จากปุ่มด้านล่าง")

    def _on_error(self, err_text):
        self._set_busy(False)
        self.result_label.configure(text="เกิดข้อผิดพลาด")
        self._log("⚠ เกิดข้อผิดพลาด:\n" + err_text)
        messagebox.showerror(
            APP_TITLE,
            "เกิดข้อผิดพลาดระหว่างตรวจสอบไฟล์\n\n"
            "กรุณาตรวจสอบว่าไฟล์เป็น .docx ที่ไม่เสียหาย (ไฟล์ .doc แบบเก่าต้องแปลง\n"
            "เป็น .docx ก่อน) ถ้ายังไม่ได้ ให้ดูรายละเอียดในช่อง \"สถานะการทำงาน\"",
        )

    def _open_result_docx(self):
        self._open_path(self.result_docx_path)

    def _open_result_excel(self):
        self._open_path(self.result_excel_path)

    def _open_result_folder(self):
        self._open_path(self.result_out_dir)

    def _open_path(self, path):
        if not path or not os.path.exists(path):
            messagebox.showwarning(APP_TITLE, "ไม่พบไฟล์/โฟลเดอร์นี้แล้ว")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"เปิดไม่สำเร็จ: {e}")


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
