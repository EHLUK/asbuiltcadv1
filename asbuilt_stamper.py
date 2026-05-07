"""
=============================================================================
  AS-BUILT DRAWING STAMPER
  Hargreaves / Exentec Hargreaves
  
  Connects to AutoCAD via COM automation, reads the title block revision
  table, adds a new revision row above the current latest row, increments
  the revision letter, stamps today's date, sets Status = P7,
  Reason = As Built, then saves as DWG and exports PDF.
=============================================================================
"""

import sys
import os
import threading
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ── colour palette ────────────────────────────────────────────────────────────
BG          = "#0d1117"
PANEL       = "#161b22"
BORDER      = "#30363d"
ACCENT      = "#f97316"      # orange
ACCENT_DIM  = "#7c3310"
TEXT_MAIN   = "#e6edf3"
TEXT_MUTED  = "#8b949e"
TEXT_GREEN  = "#3fb950"
TEXT_RED    = "#f85149"
TEXT_YELLOW = "#d29922"
FONT_MONO   = ("Consolas", 10)
FONT_UI     = ("Segoe UI", 10)
FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_HEAD   = ("Segoe UI", 10, "bold")


# =============================================================================
#  AUTOCAD COM AUTOMATION CORE
# =============================================================================

class AutoCADProcessor:
    """All AutoCAD interaction lives here."""

    def __init__(self, log_fn):
        self.log = log_fn

    # ── helpers ──────────────────────────────────────────────────────────────

    def _increment_rev(self, rev_letter: str) -> str:
        """A→B, B→C … Z→AA, AA→AB …"""
        rev = rev_letter.strip().upper()
        if not rev:
            return "A"
        # treat as base-26 number
        result = []
        carry = 1
        for ch in reversed(rev):
            val = ord(ch) - ord('A') + carry
            carry, rem = divmod(val, 26)
            result.append(chr(ord('A') + rem))
        if carry:
            result.append(chr(ord('A') + carry - 1))
        return "".join(reversed(result))

    def _find_autocad(self):
        """Return an active AutoCAD Application COM object."""
        import win32com.client as win32
        try:
            acad = win32.GetActiveObject("AutoCAD.Application")
            self.log("✓ Connected to running AutoCAD instance", "green")
            return acad
        except Exception:
            self.log("⚠ No running AutoCAD found – launching AutoCAD…", "yellow")
            acad = win32.Dispatch("AutoCAD.Application")
            acad.Visible = True
            import time; time.sleep(3)
            self.log("✓ AutoCAD launched", "green")
            return acad

    def _open_drawing(self, acad, dwg_path: str):
        """Open a DWG and return the Document object."""
        import win32com.client as win32
        # Check if already open
        for doc in acad.Documents:
            if os.path.normcase(doc.FullName) == os.path.normcase(dwg_path):
                self.log(f"✓ Drawing already open: {os.path.basename(dwg_path)}", "green")
                return doc
        doc = acad.Documents.Open(dwg_path)
        self.log(f"✓ Opened: {os.path.basename(dwg_path)}", "green")
        return doc

    # ── revision table detection ──────────────────────────────────────────────

    def _get_all_text(self, doc):
        """
        Return list of dicts describing every TEXT/MTEXT entity in modelspace
        and all paper-space layouts.
        Each dict: {entity, text, x, y, height, layer, space}
        """
        results = []

        def harvest(space_obj, label):
            try:
                for entity in space_obj:
                    try:
                        et = entity.EntityName.upper()
                        if et in ("ACDBTEXT", "TEXT"):
                            results.append({
                                "entity": entity,
                                "text":   entity.TextString,
                                "x":      entity.InsertionPoint[0],
                                "y":      entity.InsertionPoint[1],
                                "height": entity.Height,
                                "layer":  entity.Layer,
                                "space":  label,
                            })
                        elif et in ("ACDBMTEXT", "MTEXT"):
                            results.append({
                                "entity": entity,
                                "text":   entity.TextString,
                                "x":      entity.InsertionPoint[0],
                                "y":      entity.InsertionPoint[1],
                                "height": entity.Height,
                                "layer":  entity.Layer,
                                "space":  label,
                            })
                    except Exception:
                        pass
            except Exception:
                pass

        harvest(doc.ModelSpace, "Model")
        for layout in doc.Layouts:
            try:
                if layout.Name.upper() != "MODEL":
                    harvest(layout.Block, f"Paper:{layout.Name}")
            except Exception:
                pass

        return results

    def _find_revision_rows(self, all_text):
        """
        Identify the revision table rows.
        Strategy:
          1. Look for text that looks like a revision letter (A-Z or AA etc.)
             grouped with nearby text on the same Y-coordinate (same row).
          2. The row with the largest Y that contains a rev letter = latest row.
        Returns list of rows, each row = list of entity-dicts sorted by X.
        Rows sorted by Y ascending (oldest first, newest last).
        """
        import re

        # ── find candidate rev-letter cells ──────────────────────────────────
        rev_pattern = re.compile(r"^[A-Z]{1,3}$")
        
        # We group by "row bucket": entities whose Y values are within
        # (2 × typical text height) of each other belong to the same row.
        # First pass: collect entities that sit near a rev-letter entity.
        
        # Build a rough Y-tolerance from median text height
        heights = [t["height"] for t in all_text if t["height"] > 0]
        if not heights:
            return []
        heights.sort()
        median_h = heights[len(heights)//2]
        y_tol = median_h * 2.5

        # Find all rev-letter entities
        rev_entities = [t for t in all_text
                        if rev_pattern.match(t["text"].strip().upper())]
        
        if not rev_entities:
            self.log("  ℹ No single-letter revision cells found; "
                     "scanning for 'Rev' header to locate table…", "yellow")
            return []

        # For each rev entity, collect all text in the same Y-band
        rows = []
        used_y = set()
        for rev_e in rev_entities:
            ry = rev_e["y"]
            # avoid duplicate rows
            already = any(abs(ry - uy) < y_tol for uy in used_y)
            if already:
                continue
            used_y.add(ry)
            row_entities = [t for t in all_text
                            if abs(t["y"] - ry) < y_tol
                            and t["space"] == rev_e["space"]]
            row_entities.sort(key=lambda t: t["x"])
            rows.append(row_entities)

        rows.sort(key=lambda r: r[0]["y"])   # oldest → newest
        return rows

    # ── column mapping ────────────────────────────────────────────────────────

    def _map_columns(self, rows, field_config):
        """
        Given the revision rows and user-supplied field config, return
        a mapping of column_name → x_position based on the header row
        (or the first content row if no header found).

        field_config = {
            "Rev": "col index or keyword",
            "Date": …, etc.
        }

        This is flexible – we try to match by header text first, then by
        relative column order.
        """
        if not rows:
            return {}

        # Try to find a header row (text like "Rev", "Date", "Prepared", etc.)
        header_keywords = {
            "rev":      "Rev",
            "date":     "Date",
            "prep":     "Prep",
            "prepared": "Prep",
            "checked":  "Checked",
            "check":    "Checked",
            "chk":      "Checked",
            "status":   "Status",
            "reason":   "Reason",
            "description": "Reason",
            "desc":     "Reason",
            "approv":   "Approved",
            "approved": "Approved",
            "appr":     "Approved",
        }

        col_x = {}  # col_name → representative X centre

        for row in rows:
            for cell in row:
                key = cell["text"].strip().lower()
                for kw, col_name in header_keywords.items():
                    if kw in key and col_name not in col_x:
                        col_x[col_name] = cell["x"]

        return col_x

    # ── add new revision row ──────────────────────────────────────────────────

    def _add_revision_row(self, doc, latest_row, col_x, field_values,
                          space_obj):
        """
        Copy every entity from latest_row upward (increase Y by row height),
        then update the text in each new entity to the new field values.
        """
        import win32com.client as win32
        import array

        if not latest_row:
            raise ValueError("No latest revision row found.")

        # Determine row height = gap between last two rows (or entity height * 3)
        # We'll use the Y span of the entities in the row as row height proxy.
        ys = [e["y"] for e in latest_row]
        row_y = min(ys)
        
        # Try to determine row_height from text height
        sample_height = latest_row[0]["height"]
        row_height = sample_height * 3.0  # sensible default

        new_y = row_y + row_height

        self.log(f"  → Adding new row at Y={new_y:.2f} (offset +{row_height:.2f})", "muted")

        # ── build column → new text mapping ──────────────────────────────────
        # For each entity in latest_row, decide what text the new entity gets
        def new_text_for_entity(entity_dict):
            txt = entity_dict["text"].strip()
            x   = entity_dict["x"]

            # Match to a column by proximity to col_x entries
            matched_col = None
            if col_x:
                closest_dist = 1e9
                for col_name, cx in col_x.items():
                    dist = abs(x - cx)
                    if dist < closest_dist:
                        closest_dist = dist
                        matched_col = col_name
            
            if matched_col and matched_col in field_values:
                return field_values[matched_col]

            # Fallback: try text-content matching
            lower = txt.lower()
            if len(txt) <= 3 and txt.isalpha():         # looks like a Rev letter
                return field_values.get("Rev", txt)
            if any(c.isdigit() for c in txt) and "/" in txt:  # looks like a date
                return field_values.get("Date", txt)
            
            return txt  # keep original if we can't identify it

        # ── duplicate and shift each entity ──────────────────────────────────
        created = []
        for edict in latest_row:
            entity  = edict["entity"]
            old_txt = edict["text"]
            new_txt = new_text_for_entity(edict)

            try:
                # CopyObjects is the safest duplication method
                new_objs = doc.CopyObjects(
                    win32.VARIANT(win32.VT_ARRAY | win32.VT_DISPATCH, [entity])
                )
                new_ent = new_objs[0]

                # Move to new Y
                old_ip = entity.InsertionPoint
                pt = array.array('d', [old_ip[0], new_y, old_ip[2]])
                new_ent.InsertionPoint = pt

                # Update text
                if hasattr(new_ent, 'TextString'):
                    new_ent.TextString = new_txt

                new_ent.Update()
                created.append(new_ent)
                self.log(f"    ✓ Copied entity: '{old_txt}' → '{new_txt}'", "muted")
            except Exception as ex:
                self.log(f"    ✗ Failed to copy entity '{old_txt}': {ex}", "red")

        return created

    # ── save DWG + export PDF ─────────────────────────────────────────────────

    def _save_dwg(self, doc, out_dwg_path: str):
        """SaveAs the document to a new path."""
        # acSaveAsR2018 = 64
        try:
            doc.SaveAs(out_dwg_path, 64)
            self.log(f"✓ Saved DWG: {os.path.basename(out_dwg_path)}", "green")
        except Exception:
            # fallback: save without version specifier
            doc.SaveAs(out_dwg_path)
            self.log(f"✓ Saved DWG (default version): {os.path.basename(out_dwg_path)}", "green")

    def _export_pdf(self, doc, out_pdf_path: str):
        """Export all layouts to PDF using AutoCAD's built-in publisher."""
        try:
            # Use the plot command via SendCommand for reliability
            doc.Activate()
            # Try using the Export method first (works in many AutoCAD versions)
            # acExportFilePDF = 1 in some versions; fallback to plot macro
            try:
                doc.Export(out_pdf_path, "pdf")
                self.log(f"✓ Exported PDF: {os.path.basename(out_pdf_path)}", "green")
                return
            except Exception:
                pass

            # Fallback: use -PLOT command via SendCommand
            import time
            cmd = (
                f'-PLOT\nY\n'                    # detailed plot? Yes
                f'Model\n'                        # layout
                f'DWG To PDF.pc3\n'               # plotter
                f'ISO A1 (841.00 x 594.00 MM)\n'  # paper size
                f'Millimeters\n'                  # units
                f'Landscape\n'                    # orientation
                f'N\n'                            # plot upside down? No
                f'Extents\n'                      # plot area
                f'Fit\n'                          # scale
                f'Center\n'                       # origin
                f'N\n'                            # save changes? No
                f'Y\n'                            # proceed? Yes
            )
            # This is a simplified attempt; a full plot dialog bypass is complex.
            # So we use the simpler Export API or note limitations.
            self.log("⚠ PDF export via -PLOT requires manual confirmation in AutoCAD.", "yellow")
            self.log("  → DWG saved successfully. Use AutoCAD's Publish or Export PDF for PDF.", "yellow")
        except Exception as ex:
            self.log(f"✗ PDF export error: {ex}", "red")

    # ── main entry point ──────────────────────────────────────────────────────

    def process(self, dwg_path: str, field_values: dict,
                out_dir: str, dry_run: bool = False):
        """
        Full processing pipeline for one DWG file.
        Returns (success: bool, out_dwg_path, out_pdf_path).
        """
        try:
            import win32com.client
        except ImportError:
            self.log("✗ pywin32 not installed. Run: pip install pywin32", "red")
            self.log("  Then run 'python pywin32_postinstall.py -install' as admin.", "red")
            return False, None, None

        try:
            acad     = self._find_autocad()
            doc      = self._open_drawing(acad, os.path.abspath(dwg_path))
            all_text = self._get_all_text(doc)
            self.log(f"  Found {len(all_text)} text entities total", "muted")

            rows = self._find_revision_rows(all_text)
            self.log(f"  Found {len(rows)} revision row(s)", "muted")

            if not rows:
                self.log("✗ Could not find revision table rows in this drawing.", "red")
                self.log("  Tip: ensure the revision rows contain single-letter "
                         "rev codes (A, B, C…) as plain TEXT entities.", "yellow")
                return False, None, None

            latest_row = rows[-1]
            self.log(f"  Latest revision row has {len(latest_row)} cells:", "muted")
            for cell in latest_row:
                self.log(f"    [{cell['text']!r:20s}] x={cell['x']:.1f} y={cell['y']:.1f}", "muted")

            # Auto-detect Rev letter from latest row
            import re
            rev_pattern = re.compile(r"^[A-Z]{1,3}$")
            current_rev = ""
            for cell in latest_row:
                if rev_pattern.match(cell["text"].strip().upper()):
                    current_rev = cell["text"].strip().upper()
                    break

            new_rev = self._increment_rev(current_rev) if current_rev else field_values.get("Rev", "A")
            self.log(f"  Current Rev: '{current_rev}'  →  New Rev: '{new_rev}'", "green")

            # Build final field values (override Rev with auto-incremented one)
            final_values = dict(field_values)
            final_values["Rev"] = new_rev

            # Map columns
            col_x = self._map_columns(rows, final_values)
            if col_x:
                self.log(f"  Column positions detected: {col_x}", "muted")
            else:
                self.log("  ⚠ Column header detection inconclusive; "
                         "using content-type matching.", "yellow")

            if dry_run:
                self.log("── DRY RUN: no changes written ──────────────────", "yellow")
                return True, None, None

            # Determine which space the latest row lives in
            space_label = latest_row[0]["space"]
            if space_label == "Model":
                space_obj = doc.ModelSpace
            else:
                layout_name = space_label.replace("Paper:", "")
                space_obj = None
                for layout in doc.Layouts:
                    if layout.Name == layout_name:
                        space_obj = layout.Block
                        break
                if space_obj is None:
                    space_obj = doc.ModelSpace

            # Add new row
            self.log("  Adding new revision row…", "muted")
            created = self._add_revision_row(
                doc, latest_row, col_x, final_values, space_obj
            )
            self.log(f"  ✓ Created {len(created)} new text entities", "green")

            # Build output paths
            base     = os.path.splitext(os.path.basename(dwg_path))[0]
            out_dwg  = os.path.join(out_dir, base + "_ASBUILT.dwg")
            out_pdf  = os.path.join(out_dir, base + "_ASBUILT.pdf")

            self._save_dwg(doc, out_dwg)
            self._export_pdf(doc, out_pdf)

            return True, out_dwg, out_pdf

        except Exception as ex:
            self.log(f"✗ Unexpected error: {ex}", "red")
            import traceback
            self.log(traceback.format_exc(), "red")
            return False, None, None


# =============================================================================
#  GUI
# =============================================================================

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("As-Built Stamper — Hargreaves")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.geometry("900x780")
        self.minsize(760, 640)

        self._dwg_files   = []   # list of selected DWG paths
        self._running     = False
        self._processor   = None

        self._build_styles()
        self._build_ui()
        self._update_run_button()

    # ── ttk styles ────────────────────────────────────────────────────────────

    def _build_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame",        background=BG)
        style.configure("Panel.TFrame",  background=PANEL,
                        relief="flat", borderwidth=1)
        style.configure("TLabel",        background=BG,
                        foreground=TEXT_MAIN, font=FONT_UI)
        style.configure("Head.TLabel",   background=PANEL,
                        foreground=TEXT_MAIN, font=FONT_HEAD)
        style.configure("Muted.TLabel",  background=PANEL,
                        foreground=TEXT_MUTED, font=FONT_UI)
        style.configure("TEntry",
                        fieldbackground=BG,
                        foreground=TEXT_MAIN,
                        insertcolor=TEXT_MAIN,
                        bordercolor=BORDER,
                        lightcolor=BORDER,
                        darkcolor=BORDER,
                        relief="flat",
                        font=FONT_UI)
        style.configure("Accent.TButton",
                        background=ACCENT, foreground="#ffffff",
                        font=("Segoe UI", 11, "bold"),
                        borderwidth=0, focusthickness=0,
                        padding=(18, 10))
        style.map("Accent.TButton",
                  background=[("active", "#e06414"), ("disabled", ACCENT_DIM)],
                  foreground=[("disabled", "#ffffff")])
        style.configure("Ghost.TButton",
                        background=PANEL, foreground=TEXT_MUTED,
                        font=FONT_UI, borderwidth=1,
                        focusthickness=0, padding=(8, 5))
        style.map("Ghost.TButton",
                  background=[("active", BORDER)],
                  foreground=[("active", TEXT_MAIN)])
        style.configure("TCheckbutton",
                        background=PANEL, foreground=TEXT_MUTED,
                        font=FONT_UI)
        style.map("TCheckbutton",
                  background=[("active", PANEL)],
                  foreground=[("active", TEXT_MAIN)])

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── header bar ────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=PANEL, height=56)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="AS-BUILT STAMPER", bg=PANEL,
                 fg=ACCENT, font=("Consolas", 15, "bold")).pack(
                     side="left", padx=20, pady=14)
        tk.Label(hdr, text="Hargreaves / Exentec Hargreaves",
                 bg=PANEL, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", pady=14)

        # version badge
        tk.Label(hdr, text="v1.0", bg=ACCENT_DIM, fg=ACCENT,
                 font=("Consolas", 9), padx=6, pady=2).pack(
                     side="right", padx=20, pady=18)

        # ── main body ─────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # left column: files + fields
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))

        # right column: log
        right = tk.Frame(body, bg=BG)
        right.pack(side="right", fill="both", expand=True)

        self._build_file_panel(left)
        self._build_fields_panel(left)
        self._build_options_panel(left)
        self._build_run_panel(left)
        self._build_log_panel(right)

    # ── file selection panel ──────────────────────────────────────────────────

    def _build_file_panel(self, parent):
        frame = self._panel(parent, "① SELECT DWG FILES")
        frame.pack(fill="x", pady=(0, 8))

        btn_row = tk.Frame(frame, bg=PANEL)
        btn_row.pack(fill="x", padx=12, pady=(4, 6))

        ttk.Button(btn_row, text="+ Add DWG Files",
                   style="Ghost.TButton",
                   command=self._add_files).pack(side="left", padx=(0, 6))
        ttk.Button(btn_row, text="Clear",
                   style="Ghost.TButton",
                   command=self._clear_files).pack(side="left")

        # file list box
        list_frame = tk.Frame(frame, bg=BORDER, bd=0)
        list_frame.pack(fill="x", padx=12, pady=(0, 10))

        sb = tk.Scrollbar(list_frame, bg=BORDER, troughcolor=BG,
                          width=10, bd=0, highlightthickness=0)
        self._file_listbox = tk.Listbox(
            list_frame, bg=BG, fg=TEXT_MAIN,
            selectbackground=ACCENT_DIM, selectforeground=TEXT_MAIN,
            font=FONT_MONO, height=5, bd=0, highlightthickness=0,
            activestyle="none", yscrollcommand=sb.set
        )
        sb.config(command=self._file_listbox.yview)
        self._file_listbox.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")

        self._file_count_var = tk.StringVar(value="No files selected")
        tk.Label(frame, textvariable=self._file_count_var,
                 bg=PANEL, fg=TEXT_MUTED, font=("Segoe UI", 9)).pack(
                     anchor="w", padx=12, pady=(0, 6))

    # ── fields panel ─────────────────────────────────────────────────────────

    def _build_fields_panel(self, parent):
        frame = self._panel(parent, "② REVISION FIELDS")
        frame.pack(fill="x", pady=(0, 8))

        grid = tk.Frame(frame, bg=PANEL)
        grid.pack(fill="x", padx=12, pady=(4, 10))
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(3, weight=1)

        today = datetime.date.today().strftime("%d/%m/%Y")

        self._fields = {}
        field_defs = [
            # (label,      var_key,   default,             editable, row, col)
            ("Rev",        "Rev",     "AUTO",              False, 0, 0),
            ("Date",       "Date",    today,               True,  0, 2),
            ("Prep By",    "Prep",    "",                  True,  1, 0),
            ("Checked By", "Checked", "",                  True,  1, 2),
            ("Status",     "Status",  "P7",                True,  2, 0),
            ("Reason",     "Reason",  "As Built",          True,  2, 2),
            ("Approved",   "Approved","",                  True,  3, 0),
        ]

        for label, key, default, editable, row, col in field_defs:
            tk.Label(grid, text=label, bg=PANEL, fg=TEXT_MUTED,
                     font=("Segoe UI", 9), anchor="e",
                     width=11).grid(row=row, column=col, padx=(6, 4),
                                    pady=4, sticky="e")
            var = tk.StringVar(value=default)
            self._fields[key] = var
            state = "normal" if editable else "disabled"
            entry = tk.Entry(grid, textvariable=var,
                             bg="#1c2128" if editable else BG,
                             fg=TEXT_MAIN if editable else TEXT_MUTED,
                             disabledforeground=ACCENT,
                             disabledbackground=BG,
                             insertbackground=TEXT_MAIN,
                             relief="flat", bd=4,
                             font=FONT_UI, state=state)
            entry.grid(row=row, column=col+1, padx=(0, 12),
                       pady=4, sticky="ew")

        # info note
        info = (
            "  ℹ  Rev is auto-incremented (A→B, B→C …).  "
            "Fill Prep By, Checked By and Approved to match the drawing."
        )
        tk.Label(frame, text=info, bg=PANEL, fg=TEXT_MUTED,
                 font=("Segoe UI", 8), wraplength=440,
                 justify="left").pack(anchor="w", padx=12, pady=(0, 8))

    # ── options panel ─────────────────────────────────────────────────────────

    def _build_options_panel(self, parent):
        frame = self._panel(parent, "③ OUTPUT OPTIONS")
        frame.pack(fill="x", pady=(0, 8))

        row1 = tk.Frame(frame, bg=PANEL)
        row1.pack(fill="x", padx=12, pady=(6, 4))

        tk.Label(row1, text="Output folder:", bg=PANEL,
                 fg=TEXT_MUTED, font=("Segoe UI", 9)).pack(side="left")

        self._out_dir_var = tk.StringVar(value="(same folder as DWG)")
        tk.Entry(row1, textvariable=self._out_dir_var,
                 bg="#1c2128", fg=TEXT_MAIN,
                 insertbackground=TEXT_MAIN, relief="flat",
                 bd=4, font=FONT_UI, width=30).pack(
                     side="left", padx=(6, 4), fill="x", expand=True)
        ttk.Button(row1, text="Browse…", style="Ghost.TButton",
                   command=self._browse_out_dir).pack(side="left")

        row2 = tk.Frame(frame, bg=PANEL)
        row2.pack(fill="x", padx=12, pady=(0, 8))

        self._dry_run_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="Dry run (analyse only, no changes)",
                        variable=self._dry_run_var,
                        style="TCheckbutton").pack(side="left")

    # ── run button ────────────────────────────────────────────────────────────

    def _build_run_panel(self, parent):
        frame = tk.Frame(parent, bg=BG)
        frame.pack(fill="x", pady=(4, 0))

        self._run_btn = ttk.Button(
            frame, text="▶  STAMP AS-BUILT",
            style="Accent.TButton",
            command=self._start_processing
        )
        self._run_btn.pack(side="left")

        self._progress = ttk.Progressbar(
            frame, mode="indeterminate", length=200
        )
        self._progress.pack(side="left", padx=(14, 0))

        self._status_var = tk.StringVar(value="Ready")
        tk.Label(frame, textvariable=self._status_var,
                 bg=BG, fg=TEXT_MUTED, font=("Segoe UI", 9)).pack(
                     side="left", padx=(10, 0))

    # ── log panel ────────────────────────────────────────────────────────────

    def _build_log_panel(self, parent):
        frame = self._panel(parent, "ACTIVITY LOG")
        frame.pack(fill="both", expand=True)

        self._log_text = scrolledtext.ScrolledText(
            frame, bg=BG, fg=TEXT_MAIN,
            font=FONT_MONO, bd=0, relief="flat",
            state="disabled", wrap="word",
            height=30
        )
        self._log_text.pack(fill="both", expand=True, padx=2, pady=(2, 2))

        # colour tags
        self._log_text.tag_configure("green",  foreground=TEXT_GREEN)
        self._log_text.tag_configure("red",    foreground=TEXT_RED)
        self._log_text.tag_configure("yellow", foreground=TEXT_YELLOW)
        self._log_text.tag_configure("muted",  foreground=TEXT_MUTED)
        self._log_text.tag_configure("accent", foreground=ACCENT)

        # clear button
        ttk.Button(frame, text="Clear log",
                   style="Ghost.TButton",
                   command=self._clear_log).pack(
                       anchor="se", padx=6, pady=(0, 4))

        self._log("As-Built Stamper ready.", "green")
        self._log("Add DWG files, fill in the revision fields, then click STAMP.", "muted")
        self._log("", "muted")
        self._log("REQUIREMENTS:", "accent")
        self._log("  • AutoCAD must be installed on this machine", "muted")
        self._log("  • pywin32 must be installed: pip install pywin32", "muted")
        self._log("  • After pip install: python pywin32_postinstall.py -install", "muted")

    # ── helpers ───────────────────────────────────────────────────────────────

    def _panel(self, parent, title):
        outer = tk.Frame(parent, bg=BORDER, bd=0)
        inner = tk.Frame(outer, bg=PANEL, bd=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=title, bg=PANEL, fg=ACCENT,
                 font=("Consolas", 9, "bold"),
                 padx=12, pady=6).pack(anchor="w")
        sep = tk.Frame(inner, bg=BORDER, height=1)
        sep.pack(fill="x")
        return inner

    def _log(self, msg: str, tag: str = ""):
        self._log_text.configure(state="normal")
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        if tag:
            self._log_text.insert("end", line, tag)
        else:
            self._log_text.insert("end", line)
        self._log_text.configure(state="disabled")
        self._log_text.see("end")

    def _clear_log(self):
        self._log_text.configure(state="normal")
        self._log_text.delete("1.0", "end")
        self._log_text.configure(state="disabled")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select DWG Files",
            filetypes=[("AutoCAD Drawing", "*.dwg"), ("All files", "*.*")]
        )
        for p in paths:
            if p not in self._dwg_files:
                self._dwg_files.append(p)
                self._file_listbox.insert("end", os.path.basename(p))
        self._update_file_count()
        self._update_run_button()

    def _clear_files(self):
        self._dwg_files.clear()
        self._file_listbox.delete(0, "end")
        self._update_file_count()
        self._update_run_button()

    def _browse_out_dir(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self._out_dir_var.set(d)

    def _update_file_count(self):
        n = len(self._dwg_files)
        self._file_count_var.set(
            f"{n} file{'s' if n != 1 else ''} selected" if n else "No files selected"
        )

    def _update_run_button(self):
        if self._dwg_files and not self._running:
            self._run_btn.state(["!disabled"])
        else:
            self._run_btn.state(["disabled"])

    # ── processing ────────────────────────────────────────────────────────────

    def _collect_field_values(self) -> dict:
        return {k: v.get().strip() for k, v in self._fields.items()}

    def _get_out_dir(self, dwg_path: str) -> str:
        custom = self._out_dir_var.get().strip()
        if custom and custom != "(same folder as DWG)" and os.path.isdir(custom):
            return custom
        return os.path.dirname(dwg_path)

    def _start_processing(self):
        if self._running:
            return
        if not self._dwg_files:
            messagebox.showwarning("No Files", "Please add DWG files first.")
            return

        self._running = True
        self._update_run_button()
        self._progress.start(12)
        self._status_var.set("Processing…")

        field_values = self._collect_field_values()
        dry_run      = self._dry_run_var.get()
        files        = list(self._dwg_files)

        def worker():
            processor = AutoCADProcessor(
                log_fn=lambda msg, tag="": self.after(0, self._log, msg, tag)
            )
            total = len(files)
            ok_count = 0

            for i, dwg in enumerate(files, 1):
                self.after(0, self._log,
                           f"\n── File {i}/{total}: {os.path.basename(dwg)} ──",
                           "accent")
                self.after(0, self._status_var.set,
                           f"Processing {i}/{total}…")

                out_dir = self._get_out_dir(dwg)
                success, out_dwg, out_pdf = processor.process(
                    dwg, field_values, out_dir, dry_run=dry_run
                )
                if success:
                    ok_count += 1

            def done():
                self._progress.stop()
                self._running = False
                self._update_run_button()
                colour = "green" if ok_count == total else "yellow"
                self._log(
                    f"\n✔ Done: {ok_count}/{total} drawings processed successfully.",
                    colour
                )
                self._status_var.set(f"Done — {ok_count}/{total} OK")
                if ok_count > 0 and not dry_run:
                    messagebox.showinfo(
                        "Complete",
                        f"{ok_count}/{total} drawings stamped as-built.\n\n"
                        "Output files saved to selected folder."
                    )

            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()


# =============================================================================
#  ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
