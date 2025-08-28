import os
import sys
# Mitigate native lib crashes on macOS/Python 3.13 by limiting threads and allowing duplicate OpenMP
os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont
import ttkbootstrap as tb
from ttkbootstrap.constants import SUCCESS, PRIMARY, SECONDARY, LIGHT, DARK
import requests
from dotenv import load_dotenv
from PIL import Image, ImageEnhance, ImageFilter, ImageTk
import pytesseract
from pdf2image import convert_from_path, pdfinfo_from_path
import threading
import re
import numpy as np
import time
from datetime import datetime
import gc
try:
    from docx import Document
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False
# Load environment variables from .env if present
load_dotenv()

def main():
    print("កំពុងបើកកម្មវិធី...")
    # Create main window with minimal clean theme
    app = tb.Window(themename="flatly")
    app.title("អានអេអាយ")
    app.geometry("1280x900")
    app.minsize(900, 700)
    app.configure(bg='#ffffff')
    # Center the window and bring it to front briefly (helps on macOS)
    try:
        app.place_window_center()
    except Exception:
        pass
    try:
        app.attributes('-topmost', True)
        app.after(500, lambda: app.attributes('-topmost', False))
    except Exception:
        pass
    # Ensure window is visible and focused
    try:
        app.update_idletasks()
        app.deiconify()
        app.lift()
        app.focus_force()
    except Exception:
        pass

    # Simple About dialog showing app info and developer
    def show_about():
        try:
            message = (
                "AanAI\n"
                "Version: 1.0.0\n"
                "Developer: PHAL PHEAKDEY\n\n"
                "Khmer-English OCR with bundled Tesseract & Poppler."
            )
            messagebox.showinfo("About AanAI", message)
        except Exception:
            pass

    # Add a menubar with Help > About
    try:
        menubar = tk.Menu(app)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        app.config(menu=menubar)
    except Exception:
        pass

    # --- Helpers ---
    def resource_path(relative_path: str) -> str:
        """Get absolute path to resource, works for dev and PyInstaller bundle."""
        try:
            base_path = sys._MEIPASS  # type: ignore[attr-defined]
        except Exception:
            base_path = os.path.dirname(__file__)
        return os.path.join(base_path, relative_path)

    def app_base_dir() -> str:
        """Directory of the running app (dist folder when frozen)."""
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(__file__)
    def get_modern_font():
        # Modern clean fonts for minimal design
        candidates = [
            "SF Pro Display", "Inter", "Segoe UI", "Roboto", 
            "Arial", "Helvetica Neue", "system-ui"
        ]
        available = set(tkfont.families())
        for name in candidates:
            if name in available:
                return name
        return tkfont.nametofont("TkDefaultFont").actual("family")

    def pick_khmer_capable_font():
        # Try common Khmer-capable fonts; fall back to system default
        candidates = [
            "Noto Sans Khmer", "Khmer OS System", "Khmer Sangam MN", "Khmer MN",
            "Noto Serif Khmer", "Arial Unicode MS", "Segoe UI"
        ]
        available = set(tkfont.families())
        for name in candidates:
            if name in available:
                return name
        return tkfont.nametofont("TkDefaultFont").actual("family")

    # Try to register bundled Noto Sans Khmer font on Windows so Tk can use it
    def try_register_noto_sans_khmer():
        try:
            ttf_path = resource_path(os.path.join("assets", "fonts", "NotoSansKhmer-Regular.ttf"))
            if os.path.exists(ttf_path) and os.name == 'nt':
                import ctypes
                FR_PRIVATE = 0x10
                added = ctypes.windll.gdi32.AddFontResourceExW(ttf_path, FR_PRIVATE, 0)
                # Notify running apps fonts changed
                ctypes.windll.user32.SendMessageW(0xFFFF, 0x001D, 0, 0)
                return added > 0
        except Exception:
            pass
        return False

    # Try to register Noto Sans Khmer (if bundled) before we query families
    try_register_noto_sans_khmer()

    # Initialize PaddleOCR and stats tracking
    paddle_ocr = None
    processing_stats = {
        'start_time': None,
        'file_size': 0,
        'pages_processed': 0,
        'total_pages': 0,
        'characters_extracted': 0,
        'processing_speed': 0,
        'ocr_engine': 'Tesseract'
    }

    # គ្មានប្រើ Modal Progress ទៀត
    
    def map_lang_to_tess(lang: str | None) -> str:
        """Map our simple lang hint to Tesseract language codes."""
        if not lang or lang == "mixed":
            return "khm+eng"
        l = lang.lower()
        if "kh" in l or "km" in l:
            return "khm"
        return "eng"

    def extract_text_from_results(results, conf_threshold: float = 0.3):
        """[Legacy] Parser for Paddle results (kept if needed for future)."""
        lines = []
        try:
            if not results:
                return lines
            # Some versions return [list_of_lines], others return list_of_lines directly
            candidate = results
            if isinstance(results, list) and len(results) == 1 and isinstance(results[0], list):
                candidate = results[0]
            for item in candidate:
                # Expected: [box, (text, conf)]
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    text_conf = item[1]
                    if isinstance(text_conf, (list, tuple)) and len(text_conf) >= 2:
                        text, conf = text_conf[0], float(text_conf[1])
                        if conf >= conf_threshold and isinstance(text, str):
                            lines.append(text)
        except Exception:
            pass
        return lines

    def run_tesseract_with_timeout(image: Image.Image, tess_lang: str, timeout_seconds: float = 60.0) -> str:
        """Run pytesseract with a timeout to avoid indefinite stalls."""
        result_holder = {}
        error_holder = {}
        def _target():
            try:
                result_holder['r'] = pytesseract.image_to_string(image, lang=tess_lang)
            except Exception as e:
                error_holder['e'] = e
        t = threading.Thread(target=_target, daemon=True)
        t.start()
        t.join(timeout_seconds)
        if t.is_alive():
            raise TimeoutError(f"OCR timed out after {timeout_seconds:.0f}s")
        if 'e' in error_holder:
            raise error_holder['e']
        return result_holder.get('r', "")
    
    def update_stats(file_path=None, page_num=None, total_pages=None, text_length=None):
        """Update processing statistics"""
        if file_path:
            processing_stats['start_time'] = time.time()
            processing_stats['file_size'] = os.path.getsize(file_path) if os.path.exists(file_path) else 0
            processing_stats['pages_processed'] = 0
            processing_stats['characters_extracted'] = 0
        
        if total_pages:
            processing_stats['total_pages'] = total_pages
            
        if page_num:
            processing_stats['pages_processed'] = page_num
            
        if text_length:
            processing_stats['characters_extracted'] += text_length
            
        # Calculate processing speed
        if processing_stats['start_time']:
            elapsed = time.time() - processing_stats['start_time']
            if elapsed > 0:
                processing_stats['processing_speed'] = processing_stats['pages_processed'] / elapsed
    
    def format_file_size(size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def preprocess_image_for_ocr(image, lang="eng"):
        """Preprocess with stability in mind: cap size and keep RGB"""
        try:
            # Convert to RGB if needed
            if image.mode != 'RGB':
                image = image.convert('RGB')

            # Downscale very large images to reduce memory/CPU
            max_side = 2000  # cap the longest side
            w, h = image.size
            if max(w, h) > max_side:
                scale = max_side / float(max(w, h))
                new_size = (int(w * scale), int(h * scale))
                image = image.resize(new_size, Image.BILINEAR)

            return image

        except Exception:
            # If preprocessing fails, return original
            return image
    def detect_language(image):
        """Fast language detection for Khmer and English"""
        try:
            # Skip full OCR for detection - assume mixed content for speed
            # This avoids double processing and speeds up the workflow
            return "mixed"
                    
        except Exception:
            pass
        
        return "mixed"
    def find_tesseract_binary() -> str | None:
        """Try bundled vendor path first (resource_path and app dir), then common system paths."""
        # 0) Prefer bundled vendor path via resource_path
        for path_builder in (
            lambda: resource_path(os.path.join("vendor", "tesseract", "tesseract.exe")),
            lambda: os.path.join(app_base_dir(), "vendor", "tesseract", "tesseract.exe"),
        ):
            try:
                p = path_builder()
                if os.path.exists(p):
                    vendor_dir = os.path.dirname(p)
                    tessdata_dir = os.path.join(vendor_dir, "tessdata")
                    if os.path.isdir(tessdata_dir):
                        os.environ["TESSDATA_PREFIX"] = vendor_dir
                    # Put vendor dir on PATH for any DLL side-loading on Windows
                    if os.name == 'nt' and vendor_dir not in os.environ.get('PATH', ''):
                        os.environ['PATH'] = vendor_dir + os.pathsep + os.environ.get('PATH', '')
                    return p
            except Exception:
                pass
        # 1) Common system paths
        candidates = [
            "/opt/homebrew/bin/tesseract",   # macOS ARM Homebrew
            "/usr/local/bin/tesseract",      # macOS Intel Homebrew
            "/usr/bin/tesseract",            # Linux
            "C:/Program Files/Tesseract-OCR/tesseract.exe",  # Windows default
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    def preferred_tessdata_dir(tess_cmd: str | None) -> str | None:
        """Choose tessdata dir. Prefer bundled vendor; else next to the detected tesseract.exe."""
        # Prefer bundled vendor path (resource_path then app dir)
        for builder in (
            lambda: resource_path(os.path.join("vendor", "tesseract", "tessdata")),
            lambda: os.path.join(app_base_dir(), "vendor", "tesseract", "tessdata"),
        ):
            try:
                p = builder()
                parent = os.path.dirname(p)
                if os.path.isdir(p) or os.path.isdir(parent):
                    os.makedirs(p, exist_ok=True)
                    return p
            except Exception:
                pass
        # Fallback to sibling tessdata of the tesseract binary
        if tess_cmd:
            try:
                cand = os.path.join(os.path.dirname(tess_cmd), "tessdata")
                if os.path.isdir(os.path.dirname(cand)):
                    os.makedirs(cand, exist_ok=True)
                    return cand
            except Exception:
                pass
        return None

    def ensure_traineddata(lang_codes: list[str], tess_cmd: str | None) -> tuple[bool, str, str | None]:
        """Ensure <lang>.traineddata files exist. Downloads missing ones.
        Returns (ok, message, tessdata_dir).
        """
        td_dir = preferred_tessdata_dir(tess_cmd)
        if not td_dir:
            return False, "មិនអាចកំណត់ទីតាំង tessdata បានទេ", None
        base = os.path.dirname(td_dir)
        # Set TESSDATA_PREFIX to the directory containing 'tessdata'
        try:
            os.environ["TESSDATA_PREFIX"] = base
        except Exception:
            pass
        missing = []
        for l in lang_codes:
            dest = os.path.join(td_dir, f"{l}.traineddata")
            if os.path.exists(dest):
                continue
            missing.append((l, dest))
        if not missing:
            return True, f"tessdata រួចរាល់នៅ: {td_dir}", td_dir
        # Download missing from official mirrors
        for l, dest in missing:
            ok = False
            for url in (
                f"https://github.com/tesseract-ocr/tessdata/raw/main/{l}.traineddata",
                f"https://github.com/tesseract-ocr/tessdata_best/raw/main/{l}.traineddata",
            ):
                try:
                    r = requests.get(url, timeout=60)
                    if r.status_code == 200 and r.content:
                        with open(dest, "wb") as f:
                            f.write(r.content)
                        ok = True
                        break
                except Exception:
                    continue
            if not ok:
                return False, f"ខកខានទាញយក {l}.traineddata ទៅ {td_dir}", td_dir
        return True, f"បានតម្លើង traineddata ទៅ: {td_dir}", td_dir

    def ensure_lang_available(lang: str) -> tuple[bool, str]:
        """Ensure Tesseract binary and tessdata are available (prefer bundled vendor path)."""
        cmd = find_tesseract_binary()
        if cmd:
            pytesseract.pytesseract.tesseract_cmd = cmd
        # Determine required language files
        tess_lang = map_lang_to_tess(lang)
        required = [p.strip() for p in tess_lang.split('+') if p.strip()]
        ok_td, msg_td, _td_dir = ensure_traineddata(required, cmd)
        if not ok_td:
            return False, msg_td
        # Verify tesseract works
        try:
            ver = pytesseract.get_tesseract_version()
        except Exception as e:
            hint_paths = [
                resource_path(os.path.join("vendor", "tesseract", "tesseract.exe")),
                os.path.join(app_base_dir(), "vendor", "tesseract", "tesseract.exe"),
            ]
            return False, (
                "Tesseract not available. Checked: "
                + "; ".join(hint_paths)
                + f". Error: {e}"
            )
        return True, f"Using Tesseract {ver} (lang '{tess_lang}') at '{pytesseract.pytesseract.tesseract_cmd}'"
    def guess_poppler_path():
        # 0) Prefer bundled Windows vendor path
        try:
            vendor_poppler = resource_path(os.path.join("vendor", "poppler", "bin"))
            if os.path.isdir(vendor_poppler):
                # On Windows, ensure poppler bin is in PATH for runtime DLLs
                if os.name == 'nt' and vendor_poppler not in os.environ.get('PATH', ''):
                    os.environ['PATH'] = vendor_poppler + os.pathsep + os.environ.get('PATH', '')
                return vendor_poppler
        except Exception:
            pass
        # 1) Environment variable override
        env_path = os.environ.get("POPPLER_PATH")
        if env_path and os.path.isdir(env_path):
            return env_path
        # 2) Common macOS Homebrew locations
        candidates = [
            "/usr/local/opt/poppler/bin",               # Intel Homebrew (older)
            "/opt/homebrew/opt/poppler/bin",            # Apple Silicon Homebrew
            "/usr/local/bin", "/opt/homebrew/bin"
        ]
        # 3) Common Windows locations
        win_candidates = [
            r"C:\\Program Files\\poppler\\bin",
            r"C:\\Program Files (x86)\\poppler\\bin",
            r"C:\\poppler\\bin",
        ]
        if os.name == 'nt':
            candidates.extend(win_candidates)
        for p in candidates:
            if os.path.isdir(p):
                return p
        return None

    def ocr_image(path, lang: str = None):
        try:
            img = Image.open(path)
            if lang is None:
                lang = detect_language(img)
                app.after(0, lambda: progress_var.set(f"រកឃើញភាសា: {lang}"))

            # Preprocess image
            processed = preprocess_image_for_ocr(img, lang)

            # Run Tesseract with spinner
            tess_lang = map_lang_to_tess(lang)
            app.after(0, lambda: [progress.configure(mode="indeterminate"), progress.start(10), progress_var.set("កំពុងអានអក្សរ...")])
            try:
                text = run_tesseract_with_timeout(processed, tess_lang, timeout_seconds=60)
            finally:
                app.after(0, lambda: [progress.stop(), progress.configure(mode="determinate")])
            # បញ្ចប់ការកែច្នៃរូបភាព

            return text, lang
        except Exception as e:
            raise RuntimeError(f"ការអានអក្សរពីរូបភាពបរាជ័យ: {e}")

    def ocr_pdf(path, lang: str = None):
        poppler_path = guess_poppler_path()
        # Get page count first to stream pages one-by-one
        try:
            info = pdfinfo_from_path(path, poppler_path=poppler_path) if poppler_path else pdfinfo_from_path(path)
            total_pages = int(info.get("Pages", 1))
        except Exception as e:
            raise RuntimeError(f"មិនអាចអានព័ត៌មាន PDF បានទេ: {e}")

        texts = []
        detected_lang = lang

        for i in range(1, total_pages + 1):
            # Process one page at a time with proper error handling and cleanup
            page = None
            img_array = None
            try:
                # Render only one page at a time at lower DPI to reduce memory usage
                try:
                    page_imgs = convert_from_path(
                        path,
                        dpi=200,
                        first_page=i,
                        last_page=i,
                        poppler_path=poppler_path,
                    )
                except TypeError:
                    # Fallback for environments without poppler_path support
                    page_imgs = convert_from_path(
                        path,
                        dpi=200,
                        first_page=i,
                        last_page=i,
                    )
                if not page_imgs:
                    continue
                page = page_imgs[0]

                if detected_lang is None and i == 1:
                    # កំណត់ភាសាដោយស្វ័យប្រវត្តិពីទំព័រទី១
                    detected_lang = detect_language(page)
                    app.after(0, lambda: progress_var.set(f"រកឃើញភាសា: {detected_lang}"))
                    update_stats(total_pages=total_pages)

                # Preprocess page for better OCR
                processed_page = preprocess_image_for_ocr(page, detected_lang)

                # Run Tesseract per page
                tess_lang = map_lang_to_tess(detected_lang)
                app.after(0, lambda: [progress.configure(mode="indeterminate"), progress.start(10), progress_var.set(f"កំពុងអានទំព័រ {i}...")])
                try:
                    page_text = run_tesseract_with_timeout(processed_page, tess_lang, timeout_seconds=90)
                finally:
                    app.after(0, lambda: [progress.stop(), progress.configure(mode="determinate")])

                texts.append(page_text)
                update_stats(page_num=i, text_length=len(page_text))
                progress_percent = (i / total_pages) * 100
                app.after(0, lambda: [
                    progress_var.set(f"កំពុងដំណើរការ ទំព័រ {i}/{total_pages}"),
                    progress.configure(mode="determinate"),
                    progress.configure(value=progress_percent),
                    stats_var.set(f"អក្សរដែលបានស្រង់ចេញ {processing_stats['characters_extracted']} • ល្បឿន {processing_stats['processing_speed']:.1f} ទំព័រ/វិនាទី")
                ])
            except Exception as e:
                raise RuntimeError(f"ការអានអក្សរបរាជ័យលើទំព័រ {i}: {e}")
            finally:
                # Free memory explicitly
                if page is not None:
                    del page
                if img_array is not None:
                    del img_array
                gc.collect()

        return "\n\n".join(texts), detected_lang

    def choose_file_and_ocr():
        filetypes = [
            ("គាំទ្រ", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
            ("ឯកសារ PDF", "*.pdf"),
            ("រូបភាព", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
            ("ឯកសារទាំងអស់", "*.*"),
        ]
        path = filedialog.askopenfilename(title="ជ្រើសរើស PDF ឬ រូបភាព", filetypes=filetypes)
        if not path:
            return
        output.delete("1.0", tk.END)
        output.insert(tk.END, f"កំពុងដំណើរការ: {os.path.basename(path)}\n")
        app.update_idletasks()
        
        # Update file info
        file_size = os.path.getsize(path) if os.path.exists(path) else 0
        file_size_var.set(f"{os.path.basename(path)} • {format_file_size(file_size)}")
        
        # Initialize stats
        update_stats(file_path=path)
        progress_var.set("កំពុងចាប់ផ្ដើម OCR...")
        progress.configure(mode="indeterminate")
        progress.start(10)

        def worker():
            try:
                # ពិនិត្យមើល Tesseract
                ok, hint = ensure_lang_available("mixed")
                if not ok:
                    raise RuntimeError(hint)
                
                ext = os.path.splitext(path)[1].lower()
                if ext == ".pdf":
                    text, detected_lang = ocr_pdf(path)
                else:
                    text, detected_lang = ocr_image(path)
                
                def finish_ok():
                    output.delete("1.0", tk.END)
                    extracted_text = text.strip() or "<រកមិនឃើញអត្ថបទ>"
                    output.insert(tk.END, extracted_text)
                    
                    # Update final stats
                    char_count = len(extracted_text)
                    elapsed_time = time.time() - processing_stats['start_time']
                    
                    char_count_var.set(f"ចំនួនអក្សរ {char_count:,}")
                    progress_var.set(f"បានបញ្ចប់ក្នុង {elapsed_time:.1f} វិនាទី")
                    stats_var.set(f"បានស្រង់អក្សរ {char_count:,}")
                    
                    progress.stop()
                    progress.configure(value=100)
                app.after(0, finish_ok)
            except Exception as e:
                def finish_err(err):
                    progress.stop()
                    status_var.set("មានបញ្ហា")
                    messagebox.showerror("បញ្ហា OCR", str(err))
                app.after(0, finish_err, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def save_as_txt():
        text = output.get("1.0", tk.END)
        if not text:
            messagebox.showinfo("រក្សាទុក", "មិនមានអត្ថបទត្រូវរក្សាទុកទេ។")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("អត្ថបទ", "*.txt")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            messagebox.showinfo("បានរក្សាទុក", f"បានរក្សាទុកទៅ {path}")
        except Exception as e:
            messagebox.showerror("បញ្ហាក្នុងការរក្សាទុក", str(e))

    def save_as_docx():
        if not DOCX_AVAILABLE:
            messagebox.showerror("បាត់កញ្ចប់", "python-docx មិនទាន់ដំឡើង។ សូមដំឡើងដើម្បីរក្សាទុកជា .docx។")
            return
        text = output.get("1.0", tk.END)
        if not text:
            messagebox.showinfo("រក្សាទុក", "មិនមានអត្ថបទត្រូវរក្សាទុកទេ។")
            return
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("ឯកសារ Word", "*.docx")])
        if not path:
            return
        try:
            doc = Document()
            for para in text.split("\n\n"):
                doc.add_paragraph(para)
            doc.save(path)
            messagebox.showinfo("បានរក្សាទុក", f"បានរក្សាទុកទៅ {path}")
        except Exception as e:
            messagebox.showerror("បញ្ហាក្នុងការរក្សាទុក", str(e))

    def copy_text():
        # Copy selection if available; otherwise copy all
        try:
            start = output.index(tk.SEL_FIRST)
            end = output.index(tk.SEL_LAST)
        except tk.TclError:
            start, end = "1.0", tk.END
        text = output.get(start, end)
        app.clipboard_clear()
        app.clipboard_append(text)
        status_var.set("បានចម្លង")

    # --- Rich text helpers ---
    def get_selection_range():
        try:
            return output.index(tk.SEL_FIRST), output.index(tk.SEL_LAST)
        except tk.TclError:
            # No selection: apply to current word
            insert_idx = output.index(tk.INSERT)
            word_start = output.search(r"\m", insert_idx, backwards=True, regexp=True) or insert_idx
            word_end = output.search(r"\M", insert_idx, forwards=True, regexp=True) or insert_idx
            return word_start, word_end

    def toggle_tag(tag):
        start, end = get_selection_range()
        if output.tag_ranges(tag):
            output.tag_remove(tag, start, end)
        else:
            output.tag_add(tag, start, end)

    def set_bold():
        toggle_tag('bold')

    def set_italic():
        toggle_tag('italic')

    def set_underline():
        toggle_tag('underline')

    def align_left():
        start, end = get_selection_range()
        output.tag_add('left', start, end)
        output.tag_remove('center', start, end)
        output.tag_remove('right', start, end)

    def align_center():
        start, end = get_selection_range()
        output.tag_add('center', start, end)
        output.tag_remove('left', start, end)
        output.tag_remove('right', start, end)

    def align_right():
        start, end = get_selection_range()
        output.tag_add('right', start, end)
        output.tag_remove('left', start, end)
        output.tag_remove('center', start, end)

    def toggle_bullets():
        start, end = get_selection_range()
        start_line = int(float(start))
        end_line = int(float(end))
        # Decide to add or remove bullets based on first line
        line_start_idx = f"{start_line}.0"
        line_text = output.get(line_start_idx, f"{start_line}.end")
        add = not line_text.strip().startswith("• ")
        for ln in range(start_line, end_line + 1):
            li_start = f"{ln}.0"
            if add:
                output.insert(li_start, "• ")
                output.tag_add('bullet', li_start, f"{ln}.2")
            else:
                current = output.get(li_start, f"{ln}.2")
                if current == "• ":
                    output.delete(li_start, f"{ln}.2")

    def clear_formatting():
        start, end = get_selection_range()
        for tag in ('bold','italic','underline','h1','left','center','right','bullet'):
            output.tag_remove(tag, start, end)

    def increase_font():
        start, end = get_selection_range()
        output.tag_add('larger', start, end)

    def decrease_font():
        start, end = get_selection_range()
        output.tag_add('smaller', start, end)

    def copy_as_markdown():
        # Simple Markdown export (bold/italic/underline, headings, lists)
        try:
            start = output.index(tk.SEL_FIRST)
            end = output.index(tk.SEL_LAST)
        except tk.TclError:
            start, end = "1.0", tk.END
        # Normalize to Tk index strings (e.g., '1.0'), then get line numbers safely
        def _line_num(idx: str) -> int:
            norm = output.index(idx)
            return int(norm.split('.')[0])
        lines = []
        cur_line = _line_num(start)
        end_line = _line_num(end)
        for ln in range(cur_line, end_line + 1):
            lstart = f"{ln}.0"
            lend = f"{ln}.end"
            txt = output.get(lstart, lend)
            if not txt:
                lines.append("")
                continue
            # Determine formatting spans
            spans = []
            idx = lstart
            while output.compare(idx, '<', lend):
                next_idx = output.index(f"{idx} +1c")
                char = output.get(idx, next_idx)
                tags = output.tag_names(idx)
                spans.append((char, set(tags)))
                idx = next_idx
            # Build markdown line
            md = []
            bullet_prefix = "- " if 'bullet' in output.tag_names(lstart) or txt.strip().startswith('• ') else ""
            if txt.strip().startswith('• '):
                txt = txt.replace('• ', '', 1)
            active_bold = False; active_italic = False; active_underline = False
            for ch, tgs in spans:
                # Close tags if needed
                if active_bold and 'bold' not in tgs:
                    md.append('**'); active_bold = False
                if active_italic and 'italic' not in tgs:
                    md.append('*'); active_italic = False
                if active_underline and 'underline' not in tgs:
                    md.append('_'); active_underline = False
                # Open tags
                if 'bold' in tgs and not active_bold:
                    md.append('**'); active_bold = True
                if 'italic' in tgs and not active_italic:
                    md.append('*'); active_italic = True
                if 'underline' in tgs and not active_underline:
                    md.append('_'); active_underline = True
                md.append(ch)
            # Close any remaining
            if active_bold: md.append('**')
            if active_italic: md.append('*')
            if active_underline: md.append('_')
            line_md = ''.join(md)
            # Heading detection
            if 'h1' in output.tag_names(lstart):
                line_md = f"# {line_md}"
            lines.append(bullet_prefix + line_md)
        md_text = "\n".join(lines)
        app.clipboard_clear()
        app.clipboard_append(md_text)
        status_var.set("បានចម្លងជា Markdown")

    # --- AI Proofreading (Gemini / Google Generative Language API) ---
    def ai_proofread_text(text: str) -> str:
        """Send text to Google Gemini to correct OCR mistakes (Khmer + English).
        Requires GEMINI_API_KEY in environment.
        """
        api_key = "AIzaSyCsI699HjAERzJlZq6U2n_nfhK_CYO2hN8"
        if not api_key:
            raise RuntimeError("មិនឃើញ GEMINI_API_KEY នៅក្នុងបរិស្ថាន។ សូមកំណត់ GEMINI_API_KEY មុន។")
        try:
            url = (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                "gemini-2.0-flash:generateContent?key=" + api_key
            )
            system_prompt = (
                "You are a careful proofreader for OCR output containing Khmer and English. "
                "Fix recognition mistakes, spacing, punctuation, and obvious misspellings. "
                "Preserve the original meaning and formatting as much as possible. "
                "Return only the corrected text with line breaks."
            )
            payload = {
                "contents": [
                    {
                        "parts": [
                            {"text": system_prompt},
                            {"text": text},
                        ]
                    }
                ]
            }
            headers = {"Content-Type": "application/json"}
            resp = requests.post(url, headers=headers, json=payload, timeout=60)
            if resp.status_code != 200:
                # Try to surface error message from API
                try:
                    err = resp.json()
                except Exception:
                    err = resp.text
                raise RuntimeError(f"Gemini API បរាជ័យ: {resp.status_code} {err}")
            data = resp.json()
            # Typical structure: candidates[0].content.parts[0].text
            candidates = data.get("candidates") or []
            if not candidates:
                return text
            content = candidates[0].get("content") or {}
            parts = content.get("parts") or []
            # Concatenate all parts' text
            corrected = "".join(p.get("text", "") for p in parts)
            return corrected.strip() or text
        except Exception as e:
            raise RuntimeError(f"បរាជ័យក្នុងការតភ្ជាប់ទៅ Gemini: {e}")

    def start_ai_proofread():
        # Get current text
        try:
            cur_text = output.get("1.0", tk.END).strip()
        except Exception:
            cur_text = ""
        if not cur_text:
            messagebox.showinfo("ឆែកជាមួយអេអាយ", "មិនមានអត្ថបទសម្រាប់កែសម្រួលទេ។")
            return

        # Update UI state
        progress.configure(mode="indeterminate")
        progress.start(10)
        status_var.set("កំពុងពិនិត្យជាមួយ AI...")
        progress_var.set("កំពុងកែអក្ខរាវិរុទ្ធដោយ AI...")

        def worker():
            try:
                corrected = ai_proofread_text(cur_text)
                def finish_ok():
                    output.delete("1.0", tk.END)
                    output.insert(tk.END, corrected)
                    status_var.set("បានកែសម្រួលដោយ AI")
                    progress.stop()
                    progress.configure(mode="determinate", value=100)
                app.after(0, finish_ok)
            except Exception as e:
                def finish_err(msg=str(e)):
                    progress.stop()
                    progress.configure(mode="determinate")
                    status_var.set("មានបញ្ហា AI")
                    messagebox.showerror("ឆែកជាមួយអេអាយ", msg)
                app.after(0, finish_err)

        threading.Thread(target=worker, daemon=True).start()

    def clear_text():
        output.delete("1.0", tk.END)
        status_var.set("បានសម្អាត")

    # --- Modern UI Design ---
    # Get font family first
    app_font_family = pick_khmer_capable_font()
    
    # ផ្នែកក្បាលកម្មវិធី
    header_frame = tb.Frame(app, padding=0)
    header_frame.pack(fill=tk.X, padx=32, pady=(32, 24))
    
    # បង្ហាញរូបសញ្ញា (logo)
    try:
        _logo_img = Image.open(resource_path("logo.png"))
        # បន្ថយទំហំឲ្យសមរម្យ
        h = 48
        w = int(_logo_img.width * (h / _logo_img.height))
        _logo_img = _logo_img.resize((w, h), Image.BILINEAR)
        logo_photo = ImageTk.PhotoImage(_logo_img)
        logo_label = tb.Label(header_frame, image=logo_photo)
        logo_label.image = logo_photo  # guard from GC
        logo_label.pack(side=tk.RIGHT)
    except Exception:
        pass

    # ចំណងជើងកម្មវិធី
    title_font = tkfont.Font(family=get_modern_font(), size=24, weight="normal")
    title = tb.Label(header_frame, text="អានអេអាយ", 
                    font=title_font, foreground='#1a1a1a')
    title.pack(anchor='w')
    
    # ចំណងជើងរង
    subtitle_font = tkfont.Font(family=get_modern_font(), size=14, weight="normal")
    subtitle = tb.Label(header_frame, text="ស្រង់អត្ថបទពីរូបភាព និង PDF", 
                       font=subtitle_font, foreground='#6b7280')
    subtitle.pack(anchor='w', pady=(4, 0))
    
    # ផ្នែកសកម្មភាព
    action_frame = tb.Frame(app)
    action_frame.pack(fill=tk.X, padx=32, pady=(0, 24))
    
    # តំបន់ជ្រើសឯកសារ
    upload_frame = tb.Frame(action_frame, style='Card.TFrame', padding=24)
    upload_frame.pack(fill=tk.X, pady=(0, 16))
    
    # ប៊ូតុងជ្រើសរើសឯកសារ
    upload_btn = tb.Button(upload_frame, text="បើកឯកសារ", 
                          command=choose_file_and_ocr, 
                          bootstyle='outline-primary',
                          width=20)
    upload_btn.pack()
    
    # ព័ត៌មានណែនាំ
    hint_font = tkfont.Font(family=get_modern_font(), size=12)
    hint_label = tb.Label(upload_frame, text="ជ្រើសរើសឯកសារ PDF ឬ រូបភាព (PNG, JPG, JPEG, TIFF, BMP, WEBP)",
                         font=hint_font, foreground='#9ca3af')
    hint_label.pack(pady=(8, 0))
    
    # ប៊ូតុងសកម្មភាព
    action_buttons_frame = tb.Frame(app)
    action_buttons_frame.pack(fill=tk.X, padx=32, pady=(0, 32))
    
    # ប្រអប់ប៊ូតុង
    button_container = tb.Frame(action_buttons_frame)
    button_container.pack(anchor='e')
    
    # រក្សាទុកជា TXT
    save_txt_btn = tb.Button(button_container, text="រក្សាទុកជា TXT", 
                            command=save_as_txt, bootstyle='outline-secondary',
                            width=12)
    save_txt_btn.pack(side=tk.LEFT, padx=(0, 8))
    
    # រក្សាទុកជា DOCX
    if DOCX_AVAILABLE:
        save_docx_btn = tb.Button(button_container, text="រក្សាទុកជា DOCX", 
                                 command=save_as_docx, bootstyle='primary',
                                 width=12)
        save_docx_btn.pack(side=tk.LEFT, padx=(0, 8))

    # ផ្នែកស្ថិតិ និងស្ថានភាព
    stats_frame = tb.Frame(app)
    stats_frame.pack(fill=tk.X, padx=32, pady=(0, 16))
    
    # បង្ហាញស្ថិតិជា card
    stats_container = tb.Frame(stats_frame)
    stats_container.pack(fill=tk.X)
    
    # Card ព័ត៌មានឯកសារ
    file_info_frame = tb.Frame(stats_container, style='Card.TFrame', padding=16)
    file_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
    
    file_info_title = tb.Label(file_info_frame, text="ព័ត៌មានឯកសារ", 
                              font=tkfont.Font(family=get_modern_font(), size=12, weight='bold'),
                              foreground='#374151')
    file_info_title.pack(anchor='w')
    
    file_size_var = tk.StringVar(value="មិនទាន់ជ្រើសឯកសារ")
    file_size_label = tb.Label(file_info_frame, textvariable=file_size_var,
                              font=tkfont.Font(family=get_modern_font(), size=11),
                              foreground='#6b7280')
    file_size_label.pack(anchor='w', pady=(4, 0))
    
    # Card ការដំណើរការ
    progress_info_frame = tb.Frame(stats_container, style='Card.TFrame', padding=16)
    progress_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 8))
    
    progress_title = tb.Label(progress_info_frame, text="ការដំណើរការ", 
                             font=tkfont.Font(family=get_modern_font(), size=12, weight='bold'),
                             foreground='#374151')
    progress_title.pack(anchor='w')
    
    progress_var = tk.StringVar(value="រង់ចាំដំណើរការ")
    progress_label = tb.Label(progress_info_frame, textvariable=progress_var,
                             font=tkfont.Font(family=get_modern_font(), size=11),
                             foreground='#6b7280')
    progress_label.pack(anchor='w', pady=(4, 0))
    
    # Card ស្ថិតិ
    stats_info_frame = tb.Frame(stats_container, style='Card.TFrame', padding=16)
    stats_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
    
    stats_title = tb.Label(stats_info_frame, text="ស្ថិតិ", 
                          font=tkfont.Font(family=get_modern_font(), size=12, weight='bold'),
                          foreground='#374151')
    stats_title.pack(anchor='w')
    
    stats_var = tk.StringVar(value="បានស្រង់អក្សរ 0")
    stats_label = tb.Label(stats_info_frame, textvariable=stats_var,
                          font=tkfont.Font(family=get_modern_font(), size=11),
                          foreground='#6b7280')
    stats_label.pack(anchor='w', pady=(4, 0))
    
    char_count_var = tk.StringVar(value="ចំនួនអក្សរ 0")
    char_count_label = tb.Label(stats_info_frame, textvariable=char_count_var,
                          font=tkfont.Font(family=get_modern_font(), size=11),
                          foreground='#6b7280')
    char_count_label.pack(anchor='w', pady=(4, 0))
    
    # Modern progress bar
    progress_bar_frame = tb.Frame(app)
    progress_bar_frame.pack(fill=tk.X, padx=32, pady=(0, 16))
    
    progress = tb.Progressbar(progress_bar_frame, mode="determinate", 
                             bootstyle='primary')
    progress.pack(fill=tk.X)

    # តំបន់លទ្ធផលអត្ថបទ
    # ក្បាលផ្នែកលទ្ធផល
    output_header = tb.Frame(app, padding=0)
    output_header.pack(fill=tk.X, padx=32, pady=(20, 5))
    
    output_title = tb.Label(output_header, text="📝 លទ្ធផលអត្ថបទ", 
                           font=(app_font_family, 14, "bold"), 
                           bootstyle="inverse-light")
    output_title.pack(side=tk.LEFT)
    
    # Create main container with minimal padding
    main_frame = tb.Frame(app, padding=0)
    main_frame.pack(fill=tk.BOTH, expand=True)
    main_frame.configure(style='Card.TFrame')
    
    # Helper to expand/collapse non-output sections for larger reading area
    expanded = {'on': False, 'geom': None}
    def toggle_expand_output():
        if not expanded['on']:
            # Hide non-essential frames to maximize the text area
            try:
                header_frame.pack_forget()
                action_frame.pack_forget()
                action_buttons_frame.pack_forget()
                stats_frame.pack_forget()
                progress_bar_frame.pack_forget()
                output_header.pack_forget()
                status_frame.pack_forget()
            except Exception:
                pass
            try:
                # Save original geometry once
                if not expanded['geom']:
                    expanded['geom'] = app.geometry()
                app.geometry("1600x1000")
            except Exception:
                pass
            expanded['on'] = True
            try:
                expand_btn.configure(text="Restore")
            except Exception:
                pass
        else:
            # Restore the original layout
            try:
                # Ensure packing order is restored correctly: header -> action -> buttons -> stats -> progress -> output_header -> main_frame
                # Forget main_frame temporarily so output_header comes before it again
                try:
                    main_frame.pack_forget()
                except Exception:
                    pass
                header_frame.pack(fill=tk.X, padx=32, pady=(32, 24))
                action_frame.pack(fill=tk.X, padx=32, pady=(0, 24))
                action_buttons_frame.pack(fill=tk.X, padx=32, pady=(0, 32))
                stats_frame.pack(fill=tk.X, padx=32, pady=(0, 16))
                progress_bar_frame.pack(fill=tk.X, padx=32, pady=(0, 16))
                output_header.pack(fill=tk.X, padx=32, pady=(20, 5))
                main_frame.pack(fill=tk.BOTH, expand=True)
                status_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
            except Exception:
                pass
            try:
                if expanded['geom']:
                    app.geometry(expanded['geom'])
                else:
                    app.geometry("1280x900")
            except Exception:
                pass
            expanded['on'] = False
            try:
                expand_btn.configure(text="Expand")
            except Exception:
                pass

    # Formatting toolbar (like Google Docs lite)
    toolbar = tb.Frame(main_frame)
    toolbar.pack(fill=tk.X, padx=32, pady=(0, 6))
    tb.Button(toolbar, text="B", width=3, command=set_bold, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,4))
    tb.Button(toolbar, text="I", width=3, command=set_italic, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,4))
    tb.Button(toolbar, text="U", width=3, command=set_underline, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,8))
    tb.Button(toolbar, text="⟸", width=3, command=align_left, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,2))
    tb.Button(toolbar, text="≡", width=3, command=align_center, bootstyle='secondary').pack(side=tk.LEFT, padx=2)
    tb.Button(toolbar, text="⟹", width=3, command=align_right, bootstyle='secondary').pack(side=tk.LEFT, padx=(2,8))
    tb.Button(toolbar, text="• List", command=toggle_bullets, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,8))
    tb.Button(toolbar, text="A+", width=4, command=increase_font, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,4))
    tb.Button(toolbar, text="A-", width=4, command=decrease_font, bootstyle='secondary').pack(side=tk.LEFT, padx=(0,8))
    tb.Button(toolbar, text="Clear", command=clear_formatting, bootstyle='outline-danger').pack(side=tk.LEFT)
    tb.Button(toolbar, text="Copy MD", command=copy_as_markdown, bootstyle='outline-info').pack(side=tk.RIGHT)
    expand_btn = tb.Button(toolbar, text="Expand", command=toggle_expand_output, bootstyle='outline-primary')
    expand_btn.pack(side=tk.RIGHT, padx=(0,8))
    tb.Button(toolbar, text="ឆែកជាមួយអេអាយ", command=start_ai_proofread, bootstyle='success').pack(side=tk.RIGHT, padx=(0,8))

    # Text area with scrollbar
    text_frame = tb.Frame(main_frame)
    text_frame.pack(fill=tk.BOTH, expand=True, padx=32, pady=(0, 10))
    
    # បង្កើតតំបន់អត្ថបទ
    output_font = tkfont.Font(family=app_font_family, size=13)
    output = tk.Text(text_frame, wrap=tk.WORD, font=output_font,
                    bg='#ffffff', fg='#1a1a1a', insertbackground='#3498db',
                    selectbackground='#3498db', selectforeground='#ffffff',
                    relief='flat', borderwidth=0, padx=15, pady=15)
    
    scrollbar = tb.Scrollbar(text_frame, orient="vertical", command=output.yview, bootstyle="info-round")
    output.configure(yscrollcommand=scrollbar.set)
    
    # Configure rich-text tags
    output.tag_configure('bold', font=output_font.copy())
    output.tag_configure('italic', font=output_font.copy())
    output.tag_configure('underline', font=output_font.copy())
    output.tag_configure('h1', font=tkfont.Font(family=app_font_family, size=18, weight='bold'))
    output.tag_configure('left', justify='left')
    output.tag_configure('center', justify='center')
    output.tag_configure('right', justify='right')
    output.tag_configure('bullet', foreground='#111827')
    # Apply actual styles to bold/italic/underline
    output.tag_configure('bold', font=tkfont.Font(family=app_font_family, size=13, weight='bold'))
    output.tag_configure('italic', font=tkfont.Font(family=app_font_family, size=13, slant='italic'))
    output.tag_configure('underline', font=tkfont.Font(family=app_font_family, size=13, underline=1))
    output.tag_configure('larger', font=tkfont.Font(family=app_font_family, size=15))
    output.tag_configure('smaller', font=tkfont.Font(family=app_font_family, size=11))

    output.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # សារ​ស្វាគមន៍
    welcome_text = """🎯 សូមស្វាគមន៍មកកាន់កម្មវិធីអានអេអាយ!

📋 របៀបប្រើប្រាស់:
• ចុចប៊ូតុង "បើកឯកសារ" ដើម្បីជ្រើសរើសឯកសារ PDF ឬ រូបភាព
• កម្មវិធីនឹងកំណត់ភាសាដោយស្វ័យប្រវត្តិ (ខ្មែរ/អង់គ្លេស/ចម្រុះ)
• លទ្ធផលអត្ថបទនឹងបង្ហាញនៅទីនេះ
• អាចរក្សាទុកជាឯកសារ .txt ឬ .docx

🔍 ប្រភេទឯកសារដែលគាំទ្រ:
• រូបភាព: PNG, JPG, JPEG, TIF, TIFF, BMP, WEBP
• ឯកសារ: PDF

✨ លក្ខណៈពិសេស:
• កំណត់ភាសាដោយស្វ័យប្រវត្តិ
• ការកែលម្អរូបភាពសម្រាប់អានអក្សរខ្មែរ
• គាំទ្រអក្សរខ្មែរពេញលេញ"""
    
    output.insert(tk.END, welcome_text)

    # បាតបង្ហាញស្ថានភាព
    status_frame = tb.Frame(app, bootstyle="dark")
    status_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
    
    # រូបតំណាង និងអត្ថបទស្ថានភាព
    status_icon = tb.Label(status_frame, text="⚡", font=("Arial", 14))
    status_icon.pack(side=tk.LEFT, padx=(0, 8))
    
    status_var = tk.StringVar(value="រង់ចាំ...")
    status_label = tb.Label(status_frame, textvariable=status_var, 
                           font=(app_font_family, 11), 
                           bootstyle="inverse-info")
    status_label.pack(side=tk.LEFT)
    
    # Run the app
    try:
        app.mainloop()
    except Exception as e:
        import traceback
        print("Error in mainloop:", e)
        traceback.print_exc()

if __name__ == "__main__":
    main()

