"""
cert_mailer_fixed.py

FIXED VERSION: Full GUI certificate generator & sender (single-file).

Features:
 - Tkinter GUI with file pickers for template, participants (CSV/XLSX), fonts, output folder.
 - Preview first certificate.
 - Generate PDFs (no email) and Generate+Send via Gmail (use App Password).
 - Participant Name: Title Case, Poppins Bold (default 48pt), auto-shrink to minimum 40pt.
 - Security code: <eventcode>-<YY>-<xxx> placed bottom-left, Arial 10pt, BLACK (visible).
 - Logs progress in GUI.
 - Does NOT save passwords to disk.
 - Requirements: pip install pillow pandas openpyxl

FIXES:
 - Increased font sizes (48pt default, 40pt min)
 - Security code now black and larger (10pt)
 - Better positioning for landscape certificates
 - Improved underline detection
 - More generous text width calculation
"""

import os
import io
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import smtplib, ssl
from email.message import EmailMessage

# ------------------ Utility Functions ------------------

def safe_filename(s: str) -> str:
    """Make a safe filename from a name string."""
    return "".join(c for c in s if c.isalnum() or c in " _-").strip()

def title_case_name(s: str) -> str:
    """Return the name in Title Case (first letter capitalized each word)."""
    parts = [p for p in str(s).strip().split() if p]
    return " ".join(p.capitalize() for p in parts)

def load_participants(path: str) -> pd.DataFrame:
    """Load participants CSV/XLSX; require Name and Email columns (case-insensitive)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xls", ".xlsx"):
        df = pd.read_excel(path)
    else:
        raise ValueError("Participants file must be .csv, .xls or .xlsx")
    # Clean header names (strip whitespace)
    df = df.rename(columns={c: c.strip() for c in df.columns})
    cols_lower = {c.lower(): c for c in df.columns}
    if 'name' not in cols_lower or 'email' not in cols_lower:
        raise ValueError("Participants file must contain 'Name' and 'Email' columns (case-insensitive).")
    df = df[[cols_lower['name'], cols_lower['email']]]
    df.columns = ['Name', 'Email']
    df = df.dropna(subset=['Name', 'Email']).reset_index(drop=True)
    return df

def image_to_pdf_bytes(img: Image.Image) -> bytes:
    """Return PDF bytes for a PIL Image."""
    with io.BytesIO() as buf:
        img_rgb = img.convert('RGB')
        img_rgb.save(buf, format='PDF')
        return buf.getvalue()

# ------------------ Improved Underline Detection ------------------

def find_horizontal_underline_y(img: Image.Image, darkness_threshold=100, min_contig_fraction=0.4):
    """
    IMPROVED: Try to detect a long horizontal dark bar (underline) in the middle area of the template.
    Returns y coordinate (row) of the bar, or None if not found.
    Made more lenient for your certificate template.
    """
    gray = img.convert('L')
    w, h = gray.size
    px = gray.load()

    best_row = None
    best_run = 0

    # Search in middle area where underline should be
    top = h // 3  # Start searching from 1/3 down
    bottom = 2 * h // 3  # End at 2/3 down

    for y in range(top, bottom):
        curr = 0
        contig_max = 0
        curr_start = 0
        left_start_of_max = 0
        
        for x in range(w):
            if px[x, y] < darkness_threshold:
                if curr == 0:
                    curr_start = x
                curr += 1
                if curr > contig_max:
                    contig_max = curr
                    left_start_of_max = curr_start
            else:
                curr = 0
                
        if contig_max > best_run:
            best_run = contig_max
            best_row = (y, left_start_of_max, left_start_of_max + contig_max - 1)

    if best_row:
        y, left, right = best_row
        # More lenient criteria - accept if reasonable contiguous run
        if best_run >= w * min_contig_fraction or best_run >= w * 0.15:
            return y
    return None

# ------------------ Improved Drawing Routine ------------------

def draw_name_and_code_on_template(template_path: str,
                                   name_raw: str,
                                   poppins_path: str,
                                   arial_path: str,
                                   event_code: str,
                                   year: str,
                                   idx_1_based: int,
                                   font_size_default: int = 48,  # INCREASED from 34
                                   font_size_min: int = 40):     # INCREASED from 30
    """
    IMPROVED VERSION:
    - Loads template image.
    - Draws participant name in Title Case (Poppins Bold), centered above detected underline.
      Font starts at font_size_default and shrinks until fits or reaches font_size_min.
    - Draws security code bottom-left: '{event_code}-{year}-{idx:03d}' in Arial 10pt, BLACK.
    - Returns PIL.Image object (RGB).
    """
    name = title_case_name(name_raw)
    img = Image.open(template_path).convert('RGB')
    draw = ImageDraw.Draw(img)
    w, h = img.size

    # Try loading Poppins Bold at default size; if not available, try common fallbacks
    def try_load_font(path_list, size):
        for p in path_list:
            try:
                if p and os.path.isfile(p):
                    return ImageFont.truetype(p, size)
            except Exception:
                continue
        # try system fallbacks
        for sys_name in ["Poppins-Bold.ttf", "DejaVuSans-Bold.ttf", "Arial-Bold.ttf", "Arial.ttf"]:
            try:
                return ImageFont.truetype(sys_name, size)
            except Exception:
                continue
        # final fallback: load default PIL bitmap (not ideal)
        return ImageFont.load_default()

    # detect underline y
    underline_y = find_horizontal_underline_y(img)
    print(f"DEBUG: Detected underline at y={underline_y}")

    # IMPROVED: More generous width calculation
    if underline_y is not None:
        # Find the actual underline bounds
        gray = img.convert('L')
        px = gray.load()
        
        # Find longest contiguous dark run on underline_y
        best_left = None
        best_right = None
        curr_left = None
        curr_run = 0
        best_run = 0
        threshold = 100
        
        for x in range(w):
            if px[x, underline_y] < threshold:
                if curr_left is None:
                    curr_left = x
                    curr_run = 1
                else:
                    curr_run += 1
            else:
                if curr_left is not None:
                    if curr_run > best_run:
                        best_run = curr_run
                        best_left = curr_left
                        best_right = x - 1
                curr_left = None
                curr_run = 0
                
        # Handle case where underline goes to edge
        if curr_left is not None and curr_run > best_run:
            best_run = curr_run
            best_left = curr_left
            best_right = w - 1
            
        if best_left is not None and best_right is not None:
            # Be more generous with padding - allow text to be wider than underline
            padding = int((best_right - best_left) * 0.1)  # 10% padding
            avail_left = max(0, best_left - padding)
            avail_right = min(w, best_right + padding)
            max_text_width = avail_right - avail_left
            center_x_for_text = (avail_left + avail_right) // 2
            print(f"DEBUG: Underline bounds: {best_left}-{best_right}, text width: {max_text_width}")
        else:
            # fallback to generous width
            margin = int(w * 0.15)
            max_text_width = w - margin * 2
            center_x_for_text = w // 2
            print(f"DEBUG: No underline bounds found, using fallback width: {max_text_width}")
    else:
        # no underline detected: use generous margins
        margin = int(w * 0.15)
        max_text_width = w - margin * 2
        center_x_for_text = w // 2
        print(f"DEBUG: No underline detected, using full width: {max_text_width}")

    # Determine font size that fits (but be more generous)
    font_size = int(font_size_default)
    font = try_load_font([poppins_path], font_size)
    
    while True:
        # get text bbox
        bbox = draw.textbbox((0,0), name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        print(f"DEBUG: Font size {font_size}, text width: {text_w}, max allowed: {max_text_width}")
        
        if text_w <= max_text_width or font_size <= int(font_size_min):
            break
        font_size -= 2  # Reduce by 2 instead of 1 for faster convergence
        font = try_load_font([poppins_path], font_size)

    # final font
    name_font = font
    print(f"DEBUG: Final font size: {font_size}")

    # compute x,y for name
    text_bbox = draw.textbbox((0,0), name, font=name_font)
    text_w = text_bbox[2] - text_bbox[0]
    text_h = text_bbox[3] - text_bbox[1]
    x = center_x_for_text - text_w // 2

    if underline_y is not None:
        # place the bottom of text above underline with INCREASED spacing
        gap = max(25, int(text_h * 0.4))  # INCREASED gap - minimum 25px or 40% of text height
        bottom_target = underline_y - gap
        y = bottom_target - text_h
    else:
        # Center vertically in upper half if no underline
        y = (h // 3) - text_h // 2

    print(f"DEBUG: Text position: ({x}, {y}), size: {text_w}x{text_h}")

    # Draw name with subtle outline for better readability
    outline = max(1, font_size // 40)
    text_color = (0, 0, 0)  # Black text
    
    # Draw subtle outline
    for ox in range(-outline, outline+1):
        for oy in range(-outline, outline+1):
            if ox != 0 or oy != 0:
                draw.text((x+ox, y+oy), name, font=name_font, fill=(64, 64, 64))
    
    # Draw main text
    draw.text((x, y), name, font=name_font, fill=text_color)

    # FIXED: Draw security code bottom-left - BLACK and larger
    code_str = f"{event_code}-{year}-{idx_1_based:03d}"
    
    # Load Arial for security code - INCREASED size to 10pt
    try:
        if arial_path and os.path.isfile(arial_path):
            code_font = ImageFont.truetype(arial_path, 10)  # INCREASED from 8
        else:
            code_font = ImageFont.truetype("Arial.ttf", 10)  # INCREASED from 8
    except Exception:
        # fallback
        try:
            code_font = ImageFont.truetype("DejaVuSans.ttf", 10)
        except Exception:
            code_font = ImageFont.load_default()

    bbox_code = draw.textbbox((0,0), code_str, font=code_font)
    code_h = bbox_code[3] - bbox_code[1]
    pad_left = 12    # INCREASED padding
    pad_bottom = 12  # INCREASED padding
    cx = pad_left
    cy = h - pad_bottom - code_h
    
    # FIXED: Draw in BLACK instead of light gray
    draw.text((cx, cy), code_str, font=code_font, fill=(0, 0, 0))  # BLACK instead of (136,136,136)
    print(f"DEBUG: Security code '{code_str}' at ({cx}, {cy})")

    # Store debug info
    img.info['font_size_used'] = font_size
    img.info['name_text_width'] = text_w
    img.info['max_text_width'] = max_text_width
    img.info['underline_y'] = underline_y

    return img

# ------------------ Email helper ------------------

def send_email_smtp_pdf(smtp_server: str, smtp_port: int,
                        sender_email: str, app_password: str,
                        recipient_email: str, subject: str,
                        pdf_bytes: bytes, attachment_filename: str):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg.set_content("")  # empty body by request
    msg.add_attachment(pdf_bytes, maintype='application', subtype='pdf', filename=attachment_filename)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
        server.login(sender_email, app_password)
        server.send_message(msg)

# ------------------ GUI Application ------------------

class CertificateApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Certificate Generator & Sender - FIXED VERSION")
        self.geometry("900x750")  # Slightly larger window
        self.resizable(False, False)

        # state
        self.template_path = None
        self.participants_path = None
        self.participants_df = None
        self.output_folder = os.path.join(os.getcwd(), "certificates_output")
        self.poppins_path = os.path.join(os.getcwd(), "Poppins-Bold.ttf")
        self.arial_path = os.path.join(os.getcwd(), "Arial.ttf")

        # SMTP defaults (Gmail)
        self.smtp_server_var = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port_var = tk.IntVar(value=465)

        # UI build
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill='both', expand=True)

        # Section 1: Files
        s1 = ttk.LabelFrame(frm, text="1) Files", padding=8)
        s1.pack(fill='x', pady=6)
        ttk.Button(s1, text="Choose Template Image (PNG/JPG)", command=self.choose_template).grid(row=0, column=0, sticky='w')
        self.lbl_template = ttk.Label(s1, text="No template chosen", width=70)
        self.lbl_template.grid(row=0, column=1, sticky='w', padx=8)
        ttk.Button(s1, text="Choose Participants (CSV/XLSX)", command=self.choose_participants).grid(row=1, column=0, sticky='w', pady=6)
        self.lbl_part = ttk.Label(s1, text="No participants chosen", width=70)
        self.lbl_part.grid(row=1, column=1, sticky='w', padx=8)
        ttk.Button(s1, text="Choose Output Folder", command=self.choose_output_folder).grid(row=2, column=0, sticky='w', pady=6)
        self.lbl_out = ttk.Label(s1, text=self.output_folder, width=70)
        self.lbl_out.grid(row=2, column=1, sticky='w', padx=8)

        # Section 2: Fonts
        s2 = ttk.LabelFrame(frm, text="2) Fonts (optional)", padding=8)
        s2.pack(fill='x', pady=6)
        ttk.Button(s2, text="Choose Poppins-Bold.ttf", command=self.choose_poppins).grid(row=0, column=0, sticky='w')
        self.lbl_poppins = ttk.Label(s2, text=os.path.basename(self.poppins_path) if os.path.exists(self.poppins_path) else "(using system/default)", width=50)
        self.lbl_poppins.grid(row=0, column=1, sticky='w', padx=8)
        ttk.Button(s2, text="Choose Arial.ttf", command=self.choose_arial).grid(row=1, column=0, sticky='w', pady=6)
        self.lbl_arial = ttk.Label(s2, text=os.path.basename(self.arial_path) if os.path.exists(self.arial_path) else "(using system/default)", width=50)
        self.lbl_arial.grid(row=1, column=1, sticky='w', padx=8)

        # Section 3: Event details - UPDATED DEFAULTS
        s3 = ttk.LabelFrame(frm, text="3) Event details", padding=8)
        s3.pack(fill='x', pady=6)
        ttk.Label(s3, text="Event code (abcd):").grid(row=0, column=0, sticky='w')
        self.event_code_var = tk.StringVar(value="quiz")
        ttk.Entry(s3, textvariable=self.event_code_var, width=18).grid(row=0, column=1, sticky='w')
        ttk.Label(s3, text="Year (YY):").grid(row=0, column=2, sticky='w', padx=(12,0))
        self.year_var = tk.StringVar(value="25")
        ttk.Entry(s3, textvariable=self.year_var, width=8).grid(row=0, column=3, sticky='w')

        ttk.Label(s3, text="Name font size (default):").grid(row=1, column=0, sticky='w', pady=6)
        self.font_size_var = tk.IntVar(value=70)  # INCREASED to 70 as requested
        ttk.Entry(s3, textvariable=self.font_size_var, width=8).grid(row=1, column=1, sticky='w')
        ttk.Label(s3, text="Min font size:").grid(row=1, column=2, sticky='w', padx=(12,0))
        self.font_min_var = tk.IntVar(value=60)  # INCREASED to 60
        ttk.Entry(s3, textvariable=self.font_min_var, width=8).grid(row=1, column=3, sticky='w')

        # Section 4: Email / SMTP
        s4 = ttk.LabelFrame(frm, text="4) Sending (Gmail recommended - use App Password)", padding=8)
        s4.pack(fill='x', pady=6)

        # Sender Email
        ttk.Label(s4, text="Sender Email:").grid(row=0, column=0, sticky='w')
        self.sender_var = tk.StringVar()
        ttk.Entry(s4, textvariable=self.sender_var, width=36).grid(row=0, column=1, sticky='w')

        # Email Subject (new field)
        ttk.Label(s4, text="Email Subject:").grid(row=1, column=0, sticky='w', pady=6)
        self.subject_var = tk.StringVar(value="Here's your Certificate of Participation!")
        ttk.Entry(s4, textvariable=self.subject_var, width=36).grid(row=1, column=1, sticky='w')

        # App Password
        ttk.Label(s4, text="App Password:").grid(row=2, column=0, sticky='w', pady=6)
        self.password_var = tk.StringVar()
        ttk.Entry(s4, textvariable=self.password_var, show="*", width=36).grid(row=2, column=1, sticky='w')

        # SMTP Server
        ttk.Label(s4, text="SMTP Server:").grid(row=3, column=0, sticky='w')
        ttk.Entry(s4, textvariable=self.smtp_server_var, width=36).grid(row=3, column=1, sticky='w')

        # Port
        ttk.Label(s4, text="Port:").grid(row=4, column=0, sticky='w', pady=6)
        ttk.Entry(s4, textvariable=self.smtp_port_var, width=36).grid(row=4, column=1, sticky='w')


        # Buttons (Preview / Generate / Send)
        btn_frame = ttk.Frame(frm)
        btn_frame.pack(fill='x', pady=8)
        ttk.Button(btn_frame, text="Preview First Certificate", command=self.preview_first).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Generate PDFs (no email)", command=self._threaded(self.generate_pdfs)).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Send Certificates (generate + email)", command=self._threaded(self.send_certificates)).pack(side='left', padx=6)

        # Log area
        log_frame = ttk.LabelFrame(frm, text="Log / Output (Debug info included)", padding=8)
        log_frame.pack(fill='both', expand=True, pady=6)
        self.log_text = tk.Text(log_frame, wrap='word', state='disabled', height=18)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

    # ---------------- File pickers ----------------

    def choose_template(self):
        p = filedialog.askopenfilename(title="Choose template image", filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if p:
            self.template_path = p
            self.lbl_template.config(text=os.path.basename(p))
            self.log(f"Template selected: {p}")
            # Log template dimensions
            try:
                with Image.open(p) as img:
                    w, h = img.size
                    self.log(f"Template dimensions: {w} x {h} pixels")
            except Exception as e:
                self.log(f"Could not read template dimensions: {e}")

    def choose_participants(self):
        p = filedialog.askopenfilename(title="Choose participants file (CSV/XLSX)", filetypes=[("CSV/Excel", "*.csv *.xlsx *.xls")])
        if p:
            try:
                df = load_participants(p)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load participants: {e}")
                return
            self.participants_path = p
            self.participants_df = df
            self.lbl_part.config(text=f"{os.path.basename(p)} — {len(df)} rows")
            self.log(f"Participants loaded: {p} — {len(df)} rows")

    def choose_output_folder(self):
        p = filedialog.askdirectory(title="Choose output folder")
        if p:
            self.output_folder = p
            self.lbl_out.config(text=self.output_folder)
            self.log(f"Output folder: {p}")

    def choose_poppins(self):
        p = filedialog.askopenfilename(title="Choose Poppins-Bold.ttf", filetypes=[("Font file", "*.ttf")])
        if p:
            self.poppins_path = p
            self.lbl_poppins.config(text=os.path.basename(p))
            self.log(f"Poppins font selected: {p}")

    def choose_arial(self):
        p = filedialog.askopenfilename(title="Choose Arial.ttf (optional)", filetypes=[("Font file", "*.ttf")])
        if p:
            self.arial_path = p
            self.lbl_arial.config(text=os.path.basename(p))
            self.log(f"Arial font selected: {p}")

    # ---------------- Logging ----------------

    def log(self, msg: str):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + "\n")
        self.log_text.see('end')
        self.log_text.configure(state='disabled')
        # Also print to console for debugging
        print(msg)

    # ---------------- Preview ----------------

    def preview_first(self):
        if not self.template_path:
            messagebox.showwarning("Missing template", "Choose a template image first.")
            return
        if self.participants_df is None or len(self.participants_df) == 0:
            messagebox.showwarning("Missing participants", "Choose a participants file with at least one row.")
            return
        first = self.participants_df.iloc[0]
        name = str(first['Name'])
        try:
            self.log(f"Generating preview for: {name}")
            img = draw_name_and_code_on_template(
                self.template_path,
                name,
                self.poppins_path,
                self.arial_path,
                self.event_code_var.get(),
                self.year_var.get(),
                1,
                font_size_default=self.font_size_var.get(),
                font_size_min=self.font_min_var.get()
            )
            
            # Show debug info
            info_parts = []
            if 'font_size_used' in img.info:
                info_parts.append(f"Font size used: {img.info['font_size_used']}pt")
            if 'name_text_width' in img.info:
                info_parts.append(f"Text width: {img.info['name_text_width']}px")
            if 'max_text_width' in img.info:
                info_parts.append(f"Max allowed: {img.info['max_text_width']}px")
            if 'underline_y' in img.info:
                info_parts.append(f"Underline Y: {img.info['underline_y']}")
            info_text = " | ".join(info_parts)
            
            self.log(f"Preview info: {info_text}")
            
            # Display in Toplevel with scroll if needed
            preview_win = tk.Toplevel(self)
            preview_win.title(f"Preview: {name}")
            
            # Scale image if too large for screen
            screen_width = preview_win.winfo_screenwidth()
            screen_height = preview_win.winfo_screenheight()
            img_w, img_h = img.size
            
            max_preview_w = int(screen_width * 0.8)
            max_preview_h = int(screen_height * 0.7)
            
            if img_w > max_preview_w or img_h > max_preview_h:
                # Calculate scale to fit
                scale = min(max_preview_w / img_w, max_preview_h / img_h)
                new_w = int(img_w * scale)
                new_h = int(img_h * scale)
                img_preview = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                self.log(f"Preview scaled to {new_w}x{new_h} (scale: {scale:.2f})")
            else:
                img_preview = img
            
            # Convert to PhotoImage
            b = io.BytesIO()
            img_preview.save(b, format='PNG')
            b.seek(0)
            photo = tk.PhotoImage(data=b.read())
            
            frame = ttk.Frame(preview_win)
            frame.pack(fill='both', expand=True, padx=10, pady=10)
            
            lbl = ttk.Label(frame, image=photo)
            lbl.image = photo  # Keep reference
            lbl.pack()
            
            info_label = ttk.Label(frame, text=info_text, font=('Arial', 8))
            info_label.pack(pady=5)
            
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Preview error", str(e))
            self.log(f"Preview failed: {e}")

    # ---------------- Generate PDFs only ----------------

    def generate_pdfs(self):
        try:
            self._validate(need_email=False)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        os.makedirs(self.output_folder, exist_ok=True)
        df = self.participants_df
        total = len(df)
        self.log(f"Generating {total} certificates (PDF only)...")
        count = 0
        for idx, row in df.iterrows():
            name = str(row['Name'])
            try:
                img = draw_name_and_code_on_template(
                    self.template_path,
                    name,
                    self.poppins_path,
                    self.arial_path,
                    self.event_code_var.get(),
                    self.year_var.get(),
                    idx + 1,
                    font_size_default=self.font_size_var.get(),
                    font_size_min=self.font_min_var.get()
                )
                pdf_bytes = image_to_pdf_bytes(img)
                safe = safe_filename(name)
                filename = f"{safe}_{idx+1:03d}.pdf"
                with open(os.path.join(self.output_folder, filename), 'wb') as f:
                    f.write(pdf_bytes)
                self.log(f"Saved: {filename} (font: {img.info.get('font_size_used', '?')}pt)")
                count += 1
            except Exception as e:
                self.log(f"Failed to generate for {name}: {e}")
                traceback.print_exc()
        self.log(f"Generation complete: {count}/{total} certificates saved")

    # ---------------- Generate + Send ----------------

    def send_certificates(self):
        try:
            self._validate(need_email=True)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        os.makedirs(self.output_folder, exist_ok=True)
        df = self.participants_df
        total = len(df)
        sender = self.sender_var.get().strip()
        password = self.password_var.get().strip()
        smtp_server = self.smtp_server_var.get().strip()
        smtp_port = int(self.smtp_port_var.get())
        subject = self.subject_var.get().strip()
        if not subject:
            subject = "Your Certificate is here"


        self.log(f"Starting certificate generation and sending to {total} participants...")
        success = 0
        failed = 0

        for idx, row in df.iterrows():
            name = str(row['Name'])
            recipient = str(row['Email']).strip()
            safe = safe_filename(name)
            filename = f"{safe}_{idx+1:03d}.pdf"
            try:
                # Generate certificate
                img = draw_name_and_code_on_template(
                    self.template_path,
                    name,
                    self.poppins_path,
                    self.arial_path,
                    self.event_code_var.get(),
                    self.year_var.get(),
                    idx + 1,
                    font_size_default=self.font_size_var.get(),
                    font_size_min=self.font_min_var.get()
                )
                pdf_bytes = image_to_pdf_bytes(img)
                
                # Save copy locally
                with open(os.path.join(self.output_folder, filename), 'wb') as f:
                    f.write(pdf_bytes)

                # Send email (empty body as requested)
                send_email_smtp_pdf(smtp_server, smtp_port, sender, password, recipient, subject, pdf_bytes, filename)
                self.log(f"✓ Sent to {name} <{recipient}> (font: {img.info.get('font_size_used', '?')}pt)")
                success += 1
                
                # Small delay to avoid rate limiting
                import time
                time.sleep(0.5)
                
            except Exception as e:
                self.log(f"✗ Failed for {name} <{recipient}>: {e}")
                traceback.print_exc()
                failed += 1

        self.log(f"Sending finished. Success: {success}, Failed: {failed}/{total}")
        if failed == 0:
            self.log("🎉 All certificates sent successfully!")
        else:
            self.log(f"⚠️ {failed} certificates failed to send. Check the logs above for details.")

    # ---------------- Validation & helpers ----------------

    def _validate(self, need_email=False):
        if not self.template_path:
            raise ValueError("Template image not selected.")
        if not os.path.exists(self.template_path):
            raise ValueError("Template image file does not exist.")
        if self.participants_df is None or len(self.participants_df) == 0:
            raise ValueError("Participants file not selected or empty.")
        try:
            font_default = int(self.font_size_var.get())
            font_min = int(self.font_min_var.get())
            if font_default < 10 or font_min < 10:
                raise ValueError("Font sizes must be at least 10pt.")
            if font_min > font_default:
                raise ValueError("Minimum font size cannot be larger than default font size.")
        except ValueError as ve:
            if "invalid literal" in str(ve):
                raise ValueError("Font sizes must be valid integers.")
            else:
                raise ve
        
        if need_email:
            sender = self.sender_var.get().strip()
            password = self.password_var.get().strip()
            if not sender:
                raise ValueError("Sender email is required for sending certificates.")
            if "@" not in sender:
                raise ValueError("Sender email appears to be invalid.")
            if not password:
                raise ValueError("App Password is required for sending certificates.")
            
            try:
                smtp_port = int(self.smtp_port_var.get())
                if smtp_port < 1 or smtp_port > 65535:
                    raise ValueError("SMTP port must be between 1 and 65535.")
            except ValueError:
                raise ValueError("SMTP port must be a valid integer.")
            
            smtp_server = self.smtp_server_var.get().strip()
            if not smtp_server:
                raise ValueError("SMTP server is required.")

    def _threaded(self, func):
        def wrapper():
            try:
                t = threading.Thread(target=func, daemon=True)
                t.start()
            except Exception as e:
                self.log(f"Failed to start background task: {e}")
                messagebox.showerror("Error", f"Failed to start task: {e}")
        return wrapper

# ------------------ Run App ------------------

if __name__ == "__main__":
    print("Starting Certificate Generator & Sender - FIXED VERSION")
    print("Fixes applied:")
    print("- Security code now BLACK and larger (10pt)")
    print("- Increased font sizes (48pt default, 40pt minimum)")
    print("- Better underline detection and positioning")
    print("- More generous text width calculation")
    print("- Added debug information and better error handling")
    print("- Added rate limiting for email sending")
    print("-" * 60)
    
    app = CertificateApp()
    app.mainloop()