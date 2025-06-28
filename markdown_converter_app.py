# PRE-REQUISITES:
# pip install tkinterweb tkinterdnd2 markdown pdfkit
import os
from pathlib import Path
import re
import json
import sys
import markdown
import pdfkit
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import threading
# MODIFIED LINES 16-17
import threading
from tkwebview2.tkwebview2 import WebView2 # Using a modern WebView2-based renderer instead of tkinterweb
from tkinterdnd2 import DND_FILES, TkinterDnD
import urllib.parse
import base64

# Helper to find asset path for PyInstaller
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_mapping_from_file(input_file: str = resource_path("assets/emoji_mapping.json")) -> dict:
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        # Return an empty dict if the file is missing or corrupted
        return {}
    
# create a function that replaces the emoji glyphs with their corresponding html svg derived from the mapping
def find_color_folder(base_path: Path) -> Path:
    color_folders = []
    for root, dirs, files in os.walk(base_path):
        if os.path.basename(root).lower() == "flat": #'Color' or the one with complex gradients looks fucked up so use the flat ones.
            parent = Path(root).parent
            if parent.name.lower() == "default":
                # Prefer this immediately
                return Path(root)
            color_folders.append(Path(root))
    # If no 'Color' inside 'Default', return any found 'Color' folder
    return color_folders[0] if color_folders else None

def replace_glyphs_with_svg(text: str, emoji_mapping: dict, img_style: str = "") -> str:
        # Build a regex pattern to match all glyphs in one pass
        if not emoji_mapping:
            return text

        glyphs = list(emoji_mapping.keys())
        # Sort by length descending to avoid partial matches
        glyphs.sort(key=len, reverse=True)
        pattern = re.compile('|'.join(map(re.escape, glyphs)))

        # Cache SVG content to avoid repeated disk reads
        svg_cache = {}

        def remove_svg_clip_path(svg_content: str) -> str:
            # Remove <clipPath> definitions in <defs>
            svg_content = re.sub(
                r'<defs>\s*<clipPath[^>]*>.*?</clipPath>\s*</defs>', '', svg_content, flags=re.DOTALL
            )
            # Remove clip-path attributes like clip-path="url(#...)"
            svg_content = re.sub(
                r'\sclip-path="url\(#.*?\)"', '', svg_content
            )
            return svg_content

        def repl(match):
            glyph = match.group(0)
            if glyph in svg_cache:
                svg_content = svg_cache[glyph]
            else:
                folder_name = emoji_mapping[glyph]
                # MODIFICATION: Use resource_path to locate assets relative to the script/executable
                fluentui_path = resource_path("assets/fluentui_assets")
                base_assets = Path(fluentui_path) / folder_name
                color_folder = find_color_folder(base_assets)
                svg_file = None
                if color_folder:
                    svg_file = next(color_folder.glob('*.svg'), None)
                if svg_file:
                    with open(svg_file, 'r', encoding='utf-8') as svg_f:
                        svg_content = svg_f.read()
                    # Clean up and prepare SVG content
                    svg_content = re.sub(r'[\r\n\t]', ' ', svg_content)
                    svg_content = re.sub(r'\s{2,}', ' ', svg_content).strip()
                    
                    # Ensure xmlns attribute is present for compatibility
                    if 'xmlns=' not in svg_content:
                        svg_content = svg_content.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"')
                    
                    # Remove clip-paths as workaround for wkhtmltopdf SVG rendering issues
                    svg_content = remove_svg_clip_path(svg_content)

                    svg_cache[glyph] = svg_content
                else:
                    svg_cache[glyph] = glyph
                    svg_content = glyph

            # If we have SVG content, embed as background image in a span for better wkhtmltopdf compatibility
            if svg_content.startswith("<svg"):
                # Remove XML declaration if present for cleaner embedding
                svg_clean = re.sub(r'<\?xml.*?\?>', '', svg_content).strip()

                # Remove any existing width/height attributes from the <svg> tag.
                svg_no_size = re.sub(r'\s+width="[^"]+"', '', svg_clean, 1)
                svg_no_size = re.sub(r'\s+height="[^"]+"', '', svg_no_size, 1)

                # Inject the desired high-resolution dimensions to force high-quality rasterization.
                svg_with_size = re.sub(r'<svg', '<svg width="128" height="128"', svg_no_size, 1)

                # Base64 encoding for background-image
                svg_base64 = base64.b64encode(svg_with_size.encode('utf-8')).decode('ascii')
                background_style = (
                    f'display:inline-block;width:1.2em;height:1.2em;margin:0 0.1em;'
                    f'background:url(data:image/svg+xml;base64,{svg_base64}) no-repeat center center;'
                    f'background-size:contain;vertical-align:-0.3em;'
                )
                # Merge with any additional img_style
                if img_style:
                    background_style += img_style
                return f'<span style="{background_style}"></span>'
            else:
                # If it's not SVG, return the original glyph
                return svg_content

        return pattern.sub(repl, text)

class MarkdownToPDFConverter:
    def __init__(self, root):
        self.root = root
        
        # --- Core Application State ---
        self.current_file_path = None
        self.preview_update_job = None
        self.config_path = os.path.join(os.path.expanduser("~"), ".md2pdf_converter_config.json")
        self.themes_dir = os.path.join(os.path.expanduser("~"), ".md2pdf_converter_themes")

        # --- Default Settings & Configurable Variables ---
        self.output_folder = os.path.join(os.path.expanduser("~"), "Desktop", "PDF_Output")
        
        # Options with tk variables for UI binding
        self.table_handling = tk.StringVar(value="smart_fit")
        self.orientation = tk.StringVar(value="portrait")
        self.filename_var = tk.StringVar(value="document")
        self.folder_var = tk.StringVar(value=self.output_folder)
        self.page_size = tk.StringVar(value="A4")
        self.margin_top = tk.StringVar(value="0.8")
        self.margin_bottom = tk.StringVar(value="0.8")
        self.margin_left = tk.StringVar(value="0.6")
        self.margin_right = tk.StringVar(value="0.6")
        self.header_left = tk.StringVar()
        self.header_center = tk.StringVar()
        self.header_right = tk.StringVar()
        self.footer_left = tk.StringVar()
        self.footer_center = tk.StringVar(value="Page [page] of [topage]")
        self.footer_right = tk.StringVar()
        self.generate_toc = tk.BooleanVar(value=True)
        self.auto_fix_markdown = tk.BooleanVar(value=True)
        self.current_theme = tk.StringVar()
        
        # Markdown Extensions
        self.extensions_config = {
            'tables': tk.BooleanVar(value=True), 'extra': tk.BooleanVar(value=True),
            'sane_lists': tk.BooleanVar(value=True), 'fenced_code': tk.BooleanVar(value=True),
            'codehilite': tk.BooleanVar(value=True), 'nl2br': tk.BooleanVar(value=True),
            'toc': tk.BooleanVar(value=True), 'admonition': tk.BooleanVar(value=True),
            'attr_list': tk.BooleanVar(value=True), 'def_list': tk.BooleanVar(value=True),
            'footnotes': tk.BooleanVar(value=True), 'meta': tk.BooleanVar(value=False),
            'smarty': tk.BooleanVar(value=True), 'wikilinks': tk.BooleanVar(value=False)
        }

        # Trigger preview update when page layout or margins change
        self.page_size.trace_add("write", lambda *a: self.schedule_preview_update())
        self.orientation.trace_add("write", lambda *a: self.schedule_preview_update())
        self.margin_top.trace_add("write", lambda *a: self.schedule_preview_update())
        self.margin_bottom.trace_add("write", lambda *a: self.schedule_preview_update())
        self.margin_left.trace_add("write", lambda *a: self.schedule_preview_update())
        self.margin_right.trace_add("write", lambda *a: self.schedule_preview_update())

        # Re-render preview when any other option changes
        self.table_handling.trace_add("write", lambda *a: self.schedule_preview_update())
        self.current_theme.trace_add("write", lambda *a: self.schedule_preview_update())
        self.generate_toc.trace_add("write", lambda *a: self.schedule_preview_update())
        self.auto_fix_markdown.trace_add("write", lambda *a: self.schedule_preview_update())
        # Header/Footer fields
        self.header_left.trace_add("write", lambda *a: self.schedule_preview_update())
        self.header_center.trace_add("write", lambda *a: self.schedule_preview_update())
        self.header_right.trace_add("write", lambda *a: self.schedule_preview_update())
        self.footer_left.trace_add("write", lambda *a: self.schedule_preview_update())
        self.footer_center.trace_add("write", lambda *a: self.schedule_preview_update())
        self.footer_right.trace_add("write", lambda *a: self.schedule_preview_update())
        # Markdown extensions
        for var in self.extensions_config.values():
            var.trace_add("write", lambda *a: self.schedule_preview_update())

        # --- Initial Setup ---
        self.setup_themes()
        
        # --- Emoji Mapping ---
        # MODIFICATION: Load mapping and provide a console warning if it fails.
        self.emoji_mapping = load_mapping_from_file()
        if not self.emoji_mapping:
            print("Warning: Could not load 'assets/emoji_mapping.json'. Emoji-to-SVG replacement will be disabled.")

        self.load_config()
        self.ensure_output_folder()
        
        self.root.title("Markdown to PDF Converter")
        # Geometry is set in load_config
        self.root.minsize(800, 600)
        
        self.setup_ui()
        self.update_window_title()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Handle saving config before closing the app."""
        self.save_config()
        self.root.destroy()
        
    def load_config(self):
        """Load settings from the config file."""
        try:
            with open(self.config_path, 'r') as f:
                config = json.load(f)
            
            self.root.geometry(config.get("geometry", "1200x800"))
            self.output_folder = config.get("output_folder", self.output_folder)
            self.folder_var.set(self.output_folder)

            # Load options
            self.table_handling.set(config.get("table_handling", "smart_fit"))
            self.orientation.set(config.get("orientation", "portrait"))
            self.page_size.set(config.get("page_size", "A4"))
            self.margin_top.set(config.get("margin_top", "0.8"))
            self.margin_bottom.set(config.get("margin_bottom", "0.8"))
            self.margin_left.set(config.get("margin_left", "0.6"))
            self.margin_right.set(config.get("margin_right", "0.6"))
            self.header_left.set(config.get("header_left", ""))
            self.header_center.set(config.get("header_center", ""))
            self.header_right.set(config.get("header_right", ""))
            self.footer_left.set(config.get("footer_left", ""))
            self.footer_center.set(config.get("footer_center", "Page [page] of [topage]"))
            self.footer_right.set(config.get("footer_right", ""))
            self.generate_toc.set(config.get("generate_toc", True))
            self.auto_fix_markdown.set(config.get("auto_fix_markdown", True))
            self.current_theme.set(config.get("current_theme", "default_light.css"))

            for ext_name, value in config.get("extensions", {}).items():
                if ext_name in self.extensions_config:
                    self.extensions_config[ext_name].set(value)

        except (FileNotFoundError, json.JSONDecodeError):
            self.root.geometry("1200x800")
            # If config doesn't exist, defaults are already set.

    def save_config(self):
        """Save current settings to the config file."""
        config = {
            "geometry": self.root.winfo_geometry(),
            "output_folder": self.output_folder,
            "table_handling": self.table_handling.get(),
            "orientation": self.orientation.get(),
            "page_size": self.page_size.get(),
            "margin_top": self.margin_top.get(),
            "margin_bottom": self.margin_bottom.get(),
            "margin_left": self.margin_left.get(),
            "margin_right": self.margin_right.get(),
            "header_left": self.header_left.get(),
            "header_center": self.header_center.get(),
            "header_right": self.header_right.get(),
            "footer_left": self.footer_left.get(),
            "footer_center": self.footer_center.get(),
            "footer_right": self.footer_right.get(),
            "generate_toc": self.generate_toc.get(),
            "auto_fix_markdown": self.auto_fix_markdown.get(),
            "current_theme": self.current_theme.get(),
            "extensions": {name: var.get() for name, var in self.extensions_config.items()}
        }
        with open(self.config_path, 'w') as f:
            json.dump(config, f, indent=4)

    def setup_themes(self):
        """Create themes directory and default themes if they don't exist."""
        if not os.path.exists(self.themes_dir):
            os.makedirs(self.themes_dir)

        default_light_css = """
body { font-family: Barlow, sans-serif; line-height: 1.6; font-size: 16px; margin: 0; font-weight: 100; color: #111; }
h1,h2,h3,h4,h5,h6 { margin: 0.2em 0; page-break-after: avoid; }
p { font-weight: 100; color: #111; margin: 0.5em 0; orphans: 2; widows: 2; }
b, strong { color: #000; font-weight: bold; }
pre { background: #2d2d2d; border-radius: 4px; margin: 0.5em 0; padding: 1em; color: #fff; overflow-x: auto; page-break-inside: avoid; }
code { font-family: 'Fira Code', Consolas, Monaco, monospace; }
:not(pre) > code { background: #f0f0f0; padding: 2px 4px; border-radius: 3px; color: #e83e8c; }
img { max-width: 100%; height: auto; page-break-inside: avoid; }
ul, ol { margin-top: 0; }
blockquote { border-left: 4px solid #ddd; padding-left: 1em; margin-left: 0; color: #666; }
h1 { font-size: 2.2em; color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 0.5rem; }
h2 { font-size: 1.8em; color: #34495e; } h3 { font-size: 1.4em; color: #455a64; }
h4 { font-size: 1.2em; color: #546e7a; } h5 { font-size: 1.1em; color: #607d8b; }
h6 { font-size: 1em; color: #78909c; }
        """
        
        github_dark_css = """
        /* GitHub Dark Theme - Enhanced for Markdown to PDF */
        body {
        font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif,"Apple Color Emoji","Segoe UI Emoji";
        color-scheme: dark;
        background-color: #0d1117;
        color: #c9d1d9;
        line-height: 1.6;
        font-size: 16px;
        margin: 0;
        }
        h1, h2, h3, h4, h5, h6 {
        margin: 24px 0 16px;
        font-weight: 600;
        line-height: 1.25;
        padding-bottom: .3em;
        border-bottom: 1px solid #21262d;
        page-break-after: avoid;
        }
        h1 { font-size: 2em; }
        h2 { font-size: 1.5em; }
        h3 { font-size: 1.25em; }
        h4 { font-size: 1.1em; }
        h5 { font-size: 1em; }
        h6 { font-size: 0.95em; }
        p {
        margin-top: 0;
        margin-bottom: 16px;
        orphans: 2;
        widows: 2;
        }
        b, strong { font-weight: 600; color: #fff; }
        em, i { color: #d2a8ff; }
        hr {
        border: 0;
        border-top: 1px solid #21262d;
        margin: 24px 0;
        }
        a {
        color: #58a6ff;
        text-decoration: underline;
        }
        a:hover {
        color: #79c0ff;
        text-decoration: underline;
        }
        ul, ol {
        margin-top: 0;
        margin-bottom: 16px;
        padding-left: 2em;
        }
        li {
        margin-bottom: 0.25em;
        }
        pre {
        background-color: #161b22;
        word-wrap: normal;
        padding: 16px;
        overflow: auto;
        line-height: 1.45;
        border-radius: 6px;
        color: #c9d1d9;
        font-size: 0.95em;
        margin: 0.5em 0;
        page-break-inside: avoid;
        }
        code {
        font-family: 'Fira Code', Consolas, Monaco, monospace;
        font-size: 85%;
        color: #a5d6ff;
        }
        :not(pre) > code {
        padding: .2em .4em;
        margin: 0;
        background-color: rgba(110,118,129,0.4);
        border-radius: 6px;
        color: #d2a8ff;
        }
        blockquote {
        padding: 0 1em;
        color: #8b949e;
        border-left: .25em solid #30363d;
        margin: 0.5em 0;
        background: rgba(110,118,129,0.08);
        border-radius: 4px;
        }
        img {
        max-width: 100%;
        height: auto;
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 4px;
        page-break-inside: avoid;
        }
        table {
        border-collapse: collapse;
        width: 100%;
        margin: 1em 0;
        background: #161b22;
        color: #c9d1d9;
        page-break-inside: avoid;
        border-radius: 6px;
        overflow: hidden;
        }
        th, td {
        border: 1px solid #30363d;
        text-align: left;
        vertical-align: top;
        padding: 8px;
        }
        th {
        background-color: #21262d !important;
        font-weight: bold;
        color: #c9d1d9 !important;
        }
        tr:nth-child(even) td {
        background-color: #161b22;
        }
        tr:nth-child(odd) td {
        background-color: #161b22;
        }
        thead {
        background: #21262d;
        }
        tfoot {
        background: #161b22;
        font-style: italic;
        }
        mark {
        background: #bb8009;
        color: #fff;
        padding: 0 0.2em;
        border-radius: 2px;
        }
        sup, sub {
        font-size: 0.8em;
        color: #8b949e;
        }
        del {
        color: #8b949e;
        text-decoration: line-through;
        }
        kbd {
        background: #21262d;
        color: #c9d1d9;
        border: 1px solid #30363d;
        border-radius: 3px;
        padding: 2px 4px;
        font-size: 0.95em;
        font-family: inherit;
        }
        details, summary {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 4px;
        padding: 4px 8px;
        margin: 0.5em 0;
        color: #c9d1d9;
        }
        summary {
        cursor: pointer;
        font-weight: bold;
        }
        .admonition, .note, .warning, .tip, .important {
        border-left: 4px solid #d29922;
        background: #161b22;
        padding: 0.5em 1em;
        margin: 1em 0;
        border-radius: 4px;
        color: #c9d1d9;
        }
        .admonition-title {
        font-weight: bold;
        color: #d29922;
        }
        .footnote-ref, .footnote-backref {
        color: #58a6ff;
        text-decoration: underline;
        font-size: 0.9em;
        }
        .footnotes {
        font-size: 0.95em;
        color: #8b949e;
        border-top: 1px solid #30363d;
        margin-top: 2em;
        padding-top: 1em;
        }
        .toc {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 4px;
        padding: 1em;
        margin: 1em 0;
        }
        .toc ul, .toc ol {
        padding-left: 1.5em;
        }
        ::-webkit-scrollbar {
        width: 8px;
        background: #161b22;
        }
        ::-webkit-scrollbar-thumb {
        background: #30363d;
        border-radius: 4px;
        }
        """
        files_to_create = {
            "default_light.css": default_light_css,
            "github_dark.css": github_dark_css
        }
        
        for filename, content in files_to_create.items():
            path = os.path.join(self.themes_dir, filename)
            if not os.path.exists(path):
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(content.strip())
    
    def get_available_themes(self):
        """Scan themes directory for CSS files."""
        if not os.path.exists(self.themes_dir):
            return []
        return sorted([f for f in os.listdir(self.themes_dir) if f.endswith('.css')])

    def ensure_output_folder(self):
        """Create output folder if it doesn't exist"""
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
    
    def setup_ui(self):
        # --- Menu Bar ---
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", accelerator="Ctrl+N", command=self.new_file)
        file_menu.add_command(label="Open...", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="Save", accelerator="Ctrl+S", command=self.save_file)
        file_menu.add_command(label="Save As...", accelerator="Ctrl+Shift+S", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Markdown Extensions...", command=self.open_extensions_dialog)
        
        # --- Main Layout ---
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0) # Control Frame
        self.root.rowconfigure(1, weight=1) # Paned Window
        self.root.rowconfigure(2, weight=0) # Status Bar
        
        # --- Top Control Frame ---
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.grid(row=0, column=0, sticky="ew")
        
        self.convert_btn = ttk.Button(control_frame, text="Convert to PDF", command=self.convert_to_pdf)
        self.convert_btn.pack(side="left", padx=(0, 10))
        ttk.Button(control_frame, text="Paste from Clipboard", command=self.paste_from_clipboard).pack(side="left", padx=(0, 10))
        ttk.Button(control_frame, text="Open Output Folder", command=lambda: os.startfile(self.output_folder)).pack(side="left")

        # --- Paned Window for Editor and Preview ---
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        # --- Left Pane: Editor and Options ---
        left_pane = ttk.Frame(paned_window)
        paned_window.add(left_pane, weight=1)
        left_pane.columnconfigure(0, weight=1)
        left_pane.rowconfigure(0, weight=1)
        
        # Editor
        self.text_area = ScrolledText(left_pane, wrap=tk.WORD, width=70, height=18,
                                     font=("Consolas", 11), undo=True)
        self.text_area.grid(row=0, column=0, sticky="nsew")
        self.text_area.bind("<KeyRelease>", self.schedule_preview_update)
        
        # Drag-and-drop integration
        self.text_area.drop_target_register(DND_FILES)
        self.text_area.dnd_bind('<<Drop>>', self.drop_handler)
        
        # --- Right Pane: Preview and Options ---
        right_pane = ttk.Frame(paned_window)
        paned_window.add(right_pane, weight=1)
        right_pane.columnconfigure(0, weight=1)
        right_pane.rowconfigure(0, weight=1) # Preview
        right_pane.rowconfigure(1, weight=0) # Options Notebook
        
        # Live Preview
        self.html_preview = WebView2(right_pane, width=800, height=600) # NEW: Using a modern WebView2-based renderer
        self.html_preview.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        
        # Options Notebook
        notebook = ttk.Notebook(right_pane)
        notebook.grid(row=1, column=0, sticky="ew")

        # PDF Options Tab
        pdf_options_tab = ttk.Frame(notebook, padding="10")
        notebook.add(pdf_options_tab, text="PDF Options")
        
        # Filename
        fn_frame = ttk.Frame(pdf_options_tab)
        fn_frame.pack(fill="x", pady=2)
        ttk.Label(fn_frame, text="Filename:", width=12).pack(side="left")
        self.filename_entry = ttk.Entry(fn_frame, textvariable=self.filename_var)
        self.filename_entry.pack(side="left", expand=True, fill="x", padx=(0,5))
        ttk.Label(fn_frame, text=".pdf").pack(side="left")

        # Output Folder
        fd_frame = ttk.Frame(pdf_options_tab)
        fd_frame.pack(fill="x", pady=2)
        ttk.Label(fd_frame, text="Output Folder:", width=12).pack(side="left")
        self.folder_label = ttk.Label(fd_frame, textvariable=self.folder_var, relief="sunken", padding=2)
        self.folder_label.pack(side="left", expand=True, fill="x", padx=(0,5))
        ttk.Button(fd_frame, text="Browse", command=self.browse_folder).pack(side="left")

        # Theme Selector
        th_frame = ttk.Frame(pdf_options_tab)
        th_frame.pack(fill="x", pady=2)
        ttk.Label(th_frame, text="Theme:", width=12).pack(side="left")
        self.theme_selector = ttk.Combobox(th_frame, textvariable=self.current_theme, values=self.get_available_themes(), state="readonly")
        self.theme_selector.pack(side="left", expand=True, fill="x")
        self.theme_selector.bind("<<ComboboxSelected>>", self.schedule_preview_update)

        # Page Layout
        pl_frame = ttk.LabelFrame(pdf_options_tab, text="Page Layout", padding=5)
        pl_frame.pack(fill="x", pady=(10,2))
        
        ttk.Label(pl_frame, text="Page Size:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(pl_frame, textvariable=self.page_size, values=["A4", "Letter", "Legal", "A3", "A5"], state="readonly", width=10).grid(row=0, column=1, sticky="w")
        
        ttk.Label(pl_frame, text="Orientation:").grid(row=0, column=2, sticky="w", padx=(10,5), pady=2)
        ttk.Radiobutton(pl_frame, text="Portrait", variable=self.orientation, value="portrait").grid(row=0, column=3, sticky="w")
        ttk.Radiobutton(pl_frame, text="Landscape", variable=self.orientation, value="landscape").grid(row=0, column=4, sticky="w")
        
        # Margins (in)
        m_frame = ttk.LabelFrame(pdf_options_tab, text="Margins (inches)", padding=5)
        m_frame.pack(fill="x", pady=2)
        
        ttk.Label(m_frame, text="Top:").grid(row=0, column=0, padx=5)
        ttk.Entry(m_frame, textvariable=self.margin_top, width=5).grid(row=0, column=1)
        ttk.Label(m_frame, text="Bottom:").grid(row=0, column=2, padx=5)
        ttk.Entry(m_frame, textvariable=self.margin_bottom, width=5).grid(row=0, column=3)
        ttk.Label(m_frame, text="Left:").grid(row=0, column=4, padx=5)
        ttk.Entry(m_frame, textvariable=self.margin_left, width=5).grid(row=0, column=5)
        ttk.Label(m_frame, text="Right:").grid(row=0, column=6, padx=5)
        ttk.Entry(m_frame, textvariable=self.margin_right, width=5).grid(row=0, column=7)

        # Advanced Tab
        adv_options_tab = ttk.Frame(notebook, padding="10")
        notebook.add(adv_options_tab, text="Advanced")
        
        # Header/Footer
        hf_frame = ttk.LabelFrame(adv_options_tab, text="Header & Footer (use [page], [topage], [date])", padding=5)
        hf_frame.pack(fill="x", pady=2)
        
        ttk.Label(hf_frame, text="H. Left:").grid(row=0, column=0, sticky="w")
        ttk.Entry(hf_frame, textvariable=self.header_left).grid(row=0, column=1, sticky="ew")
        ttk.Label(hf_frame, text="H. Center:").grid(row=0, column=2, sticky="w")
        ttk.Entry(hf_frame, textvariable=self.header_center).grid(row=0, column=3, sticky="ew")
        ttk.Label(hf_frame, text="H. Right:").grid(row=0, column=4, sticky="w")
        ttk.Entry(hf_frame, textvariable=self.header_right).grid(row=0, column=5, sticky="ew")
        
        ttk.Label(hf_frame, text="F. Left:").grid(row=1, column=0, sticky="w")
        ttk.Entry(hf_frame, textvariable=self.footer_left).grid(row=1, column=1, sticky="ew")
        ttk.Label(hf_frame, text="F. Center:").grid(row=1, column=2, sticky="w")
        ttk.Entry(hf_frame, textvariable=self.footer_center).grid(row=1, column=3, sticky="ew")
        ttk.Label(hf_frame, text="F. Right:").grid(row=1, column=4, sticky="w")
        ttk.Entry(hf_frame, textvariable=self.footer_right).grid(row=1, column=5, sticky="ew")
        hf_frame.columnconfigure(1, weight=1)
        hf_frame.columnconfigure(3, weight=1)
        hf_frame.columnconfigure(5, weight=1)

        # Table/Misc
        tm_frame = ttk.LabelFrame(adv_options_tab, text="Content Handling", padding=5)
        tm_frame.pack(fill="x", pady=(10,2))
        
        ttk.Label(tm_frame, text="Wide table handling:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        table_frame = ttk.Frame(tm_frame)
        table_frame.grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(table_frame, text="Smart", variable=self.table_handling, value="smart_fit").pack(side="left", padx=(0, 5))
        ttk.Radiobutton(table_frame, text="Small Font", variable=self.table_handling, value="smaller_font").pack(side="left", padx=(0, 5))
        ttk.Radiobutton(table_frame, text="Break Words", variable=self.table_handling, value="break_words").pack(side="left", padx=(0, 5))

        ttk.Checkbutton(tm_frame, text="Generate Table of Contents", variable=self.generate_toc).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Checkbutton(tm_frame, text="Auto-fix common Markdown errors", variable=self.auto_fix_markdown).grid(row=2, column=0, columnspan=2, sticky="w")

        # --- Status Bar ---
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", padding=5)
        self.status_bar.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        
        # --- Progress Bar (initially hidden) ---
        self.progress = ttk.Progressbar(self.root, mode='indeterminate')
        
        # --- Keyboard Shortcuts ---
        self.root.bind("<Control-n>", lambda event: self.new_file())
        self.root.bind("<Control-o>", lambda event: self.open_file())
        self.root.bind("<Control-s>", lambda event: self.save_file())
        self.root.bind("<Control-S>", lambda event: self.save_file_as()) # Captial S for Shift
        
    def open_extensions_dialog(self):
        """Open a Toplevel window to manage Markdown extensions."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Markdown Extensions")
        dialog.geometry("300x450")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Select extensions to enable:", font="-weight bold").pack(pady=10)
        
        ext_frame = ttk.Frame(dialog, padding=10)
        ext_frame.pack(expand=True, fill="both")

        for i, (name, var) in enumerate(self.extensions_config.items()):
            ttk.Checkbutton(ext_frame, text=name, variable=var).pack(anchor="w")

        ttk.Button(dialog, text="Done", command=dialog.destroy).pack(pady=10)
    
    def schedule_preview_update(self, event=None):
        """Schedule a delayed update to the HTML preview pane."""
        if self.preview_update_job:
            self.root.after_cancel(self.preview_update_job)
        self.preview_update_job = self.root.after(500, self._render_preview)

    def _render_preview(self):
        """Renders the markdown text to the HTML preview pane."""
        md_text = self.text_area.get(1.0, tk.END)
        html_body = self._create_html_body(md_text)
        
        table_css = self.get_table_css(html_body)
        
        theme_path = os.path.join(self.themes_dir, self.current_theme.get())
        theme_css = ""
        try:
            if os.path.exists(theme_path):
                with open(theme_path, 'r', encoding='utf-8') as f:
                    theme_css = f.read()
        except Exception:
            theme_css = "body { color: red; font-family: sans-serif; } /* THEME FAILED TO LOAD */"

        PAGE_SIZES_IN = {
            "A4": (8.27, 11.69), "Letter": (8.5, 11.0), "Legal": (8.5, 14.0),
            "A3": (11.69, 16.53), "A5": (5.83, 8.27)
        }
        page_w_in, page_h_in = PAGE_SIZES_IN.get(self.page_size.get(), (8.27, 11.69))
        if self.orientation.get() == 'landscape':
            page_w_in, page_h_in = page_h_in, page_w_in

        try:
            margin_t = float(self.margin_top.get())
            margin_b = float(self.margin_bottom.get())
            margin_l = float(self.margin_left.get())
            margin_r = float(self.margin_right.get())
        except (ValueError, tk.TclError): # Handle invalid or empty fields
            margin_t, margin_b, margin_l, margin_r = 0.8, 0.8, 0.6, 0.6
        
        content_w_in = page_w_in - margin_l - margin_r

        preview_specific_css = f"""
        html {{
            padding: 2em 0;          /* Vertical spacing for the shadow effect */
            display: flex;
            justify-content: center;
        }}
        body {{
            /* Set page width based on paper size minus horizontal margins */
            width: {content_w_in}in;
            max-width: 98%;          /* Prevent overflow on narrow windows and allow slight padding */

            /* Use padding to simulate the document margins */
            padding: {margin_t}in {margin_r}in {margin_b}in {margin_l}in;

            /* Visual styling to make it look like a page */
            box-shadow: 0 0.5rem 2rem rgba(0,0,0,0.4);
            box-sizing: border-box;  /* Include padding and border in the element's total width and height */
            margin: 0 !important;    /* Override any margin the theme might set */
        }}
        """

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                /* Base theme styles from user selection */
                {theme_css}
                
                /* Dynamic styles for tables */
                {table_css}

                /* Preview-only styles to simulate page layout */
                @media screen {{
                    {preview_specific_css}
                }}
            </style>
        </head>
        <body>
        {html_body}
        </body>
        </html>
        """
        self.html_preview.load_html(full_html)
    
    def drop_handler(self, event):
        """Handle file drop event."""
        filepath = event.data.strip('{}') # Clean up path string from dnd
        if filepath.endswith(('.md', '.txt')):
            self.open_file(filepath=filepath)
        else:
            messagebox.showwarning("Invalid File", "Please drop a Markdown (.md) or Text (.txt) file.")
    
    def new_file(self):
        if self.text_area.get(1.0, tk.END).strip():
            if not messagebox.askyesno("Confirm", "Clear current content?"):
                return
        self.current_file_path = None
        self.text_area.delete(1.0, tk.END)
        self.filename_var.set("document")
        self.update_window_title()
        self.status_var.set("New file created.")
        self.schedule_preview_update()

    def open_file(self, filepath=None):
        if not filepath:
            filepath = filedialog.askopenfilename(
                defaultextension=".md",
                filetypes=[("Markdown Files", "*.md"), ("Text Files", "*.txt"), ("All Files", "*.*")]
            )
        if not filepath:
            return
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, content)
            self.current_file_path = filepath
            filename = os.path.basename(filepath)
            self.filename_var.set(os.path.splitext(filename)[0])
            self.update_window_title()
            self.status_var.set(f"Opened: {filename}")
            self.schedule_preview_update()
        except Exception as e:
            messagebox.showerror("Error Opening File", str(e))

    def save_file(self):
        if not self.current_file_path:
            self.save_file_as()
        else:
            try:
                with open(self.current_file_path, 'w', encoding='utf-8') as f:
                    f.write(self.text_area.get(1.0, tk.END))
                self.update_window_title()
                self.status_var.set(f"Saved: {os.path.basename(self.current_file_path)}")
            except Exception as e:
                messagebox.showerror("Error Saving File", str(e))
    
    def save_file_as(self):
        filepath = filedialog.asksaveasfilename(
            initialfile=os.path.basename(self.current_file_path) if self.current_file_path else "untitled.md",
            defaultextension=".md",
            filetypes=[("Markdown Files", "*.md"), ("All Files", "*.*")]
        )
        if not filepath:
            return
        self.current_file_path = filepath
        self.save_file()

    def update_window_title(self):
        base_title = "Markdown to PDF Converter"
        if self.current_file_path:
            self.root.title(f"{os.path.basename(self.current_file_path)} - {base_title}")
        else:
            self.root.title(base_title)
            
    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder)
        if folder:
            self.output_folder = folder
            self.folder_var.set(folder)
    
    def clear_text(self):
        self.text_area.delete(1.0, tk.END)
        self.status_var.set("Text cleared")
        self.schedule_preview_update()
    
    def paste_from_clipboard(self):
        try:
            clipboard_text = self.root.clipboard_get()
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, clipboard_text)
            self.status_var.set("Text pasted from clipboard")
            self.schedule_preview_update()
        except tk.TclError:
            messagebox.showwarning("Warning", "No text found in clipboard")

    def correct_table_spacing(self, md_text: str) -> str:
        """Adds blank lines before tables if they're missing, which is required for proper markdown rendering."""
        table_row_pattern = r'^\s*\|.*\|\s*$'
        lines = md_text.split('\n')
        corrected_lines = []
        for i, line in enumerate(lines):
            if re.match(table_row_pattern, line):
                is_table_start = True
                if i > 0:
                    prev_line = lines[i-1].strip()
                    if (re.match(table_row_pattern, prev_line) or 
                        re.match(r'^\s*\|[\s\-\:]*\|\s*$', prev_line)):
                        is_table_start = False
                if is_table_start and i > 0:
                    prev_line = lines[i-1].strip()
                    if prev_line and not re.match(r'^\s*$', prev_line):
                        corrected_lines.append('')
            corrected_lines.append(line)
        return '\n'.join(corrected_lines)

    def correct_table_separator_spacing(self, md_text: str) -> str:
        """Ensures table separators (|---|---| lines) have proper spacing around them."""
        separator_pattern = r'^\s*\|[\s\-\:\|]*\|\s*$'
        lines = md_text.split('\n')
        corrected_lines = []
        for i, line in enumerate(lines):
            if re.match(separator_pattern, line) and re.search(r'[\-]+', line):
                corrected_lines.append(line)
            else:
                corrected_lines.append(line)
        return '\n'.join(corrected_lines)
    
    def correct_markdown_table_list_spacing(self, md_text: str) -> str:
        """Corrects markdown formatting errors where lists or headings immediately
        follow a table without a blank line, causing rendering issues."""
        pattern = re.compile(
            r"(^\s*\|.*\|\s*$)\n(^\s*(?:[\*\-\+]|\d+\.(?!\S)|#+)\s+.*$)",
            re.MULTILINE
        )
        previous_text = None
        max_iterations = 10
        count = 0
        corrected_text = md_text
        while corrected_text != previous_text and count < max_iterations:
            previous_text = corrected_text
            corrected_text = pattern.sub(r"\1\n\n\2", previous_text)
            count += 1
            if previous_text == corrected_text:
                break
        return corrected_text
    
    def correct_general_list_and_heading_spacing(self, md_text: str) -> str:
        """Corrects markdown formatting errors where lists or headings immediately
        follow a paragraph-like line without a blank line.
        Ensures that list items are not incorrectly separated from each other.
        """
        pattern = re.compile(
            r"("
            r"^"
            r"(?!(?:[ \t]+.*)\n\s*(?:[\*\-\+]|\d+\.(?!\S))\s+)"
            r"[ \t]*"
            r"(?!(?:[\*\-\+]|\d+\.(?!\S))\s+)"
            r"(?!(?:  (?:[\*\-\+]|\d+\.)[ \t]+|  \#+[ \t]+|  [ \t]*\||  [ \t]*>|  [ \t]*(?:---|\*\*\*|___)[ \t]*$|  [ \t]*(?:```|~~~)))"
            r"(?![ \t]*$)"
            r".+"
            r"$"
            r")"
            r"\n"
            r"("
            r"^[ \t]*(?:[\*\-\+]|\d+\.(?!\S)|\#+)\s+.*$"
            r")",
            re.MULTILINE | re.VERBOSE
        )
        previous_text = None
        max_iterations = 10 
        count = 0
        corrected_text = md_text
        while corrected_text != previous_text and count < max_iterations:
            previous_text = corrected_text
            corrected_text = pattern.sub(r"\1\n\n\2", previous_text)
            count += 1
            if previous_text == corrected_text:
                break
        return corrected_text

    def process_relative_image_paths(self, html_content: str, base_path: str) -> str:
        if not base_path:
            return html_content
        
        def replacer(match):
            url = match.group(2)
            if re.match(r'^(https?://|file://|data:|/|\\|[A-Za-z]:\\)', url):
                return match.group(0)
            
            abs_path = os.path.join(base_path, url)
            abs_path = os.path.normpath(abs_path).replace('\\', '/')
            return f'{match.group(1)}file:///{abs_path}{match.group(3)}'

        pattern = re.compile(r'(<img[^>]*src=)(["\'])(.*?)\2')
        return pattern.sub(replacer, html_content)
        
    def analyze_table_width(self, html_content):
        import re
        table_pattern = r'<table[^>]*>(.*?)</table>'
        tables = re.findall(table_pattern, html_content, re.DOTALL)
        max_columns = 0
        for table in tables:
            header_match = re.search(r'<tr[^>]*>(.*?)</tr>', table, re.DOTALL)
            if header_match:
                cells = re.findall(r'<t[hd][^>]*>', header_match.group(1))
                max_columns = max(max_columns, len(cells))
        return max_columns
    
    def get_table_css(self, html_content):
        table_handling = self.table_handling.get()
        max_columns = self.analyze_table_width(html_content)
        is_landscape = self.orientation.get() == "landscape"

        # Only set background colors if using the default light theme
        theme = self.current_theme.get()
        if theme == "default_light.css":
            th_bg = "background-color: #f4f4f4;"
            th_color = ""
            td_bg = ""
        else:
            th_bg = ""
            th_color = ""
            td_bg = ""

        base_css = (
            f"table {{ border-collapse: collapse; width: 100%; margin: 1em 0; page-break-inside: avoid; }} "
            f"th, td {{ border: 1px solid #ddd; text-align: left; vertical-align: top; padding: 8px; {td_bg} }} "
            f"th {{ {th_bg} font-weight: bold; {th_color} }}"
        )

        if table_handling == "smart_fit":
            if max_columns > 8 or (max_columns > 6 and not is_landscape):
                return base_css + "table { font-size: 0.7em; } th, td { padding: 4px 6px; word-wrap: break-word; hyphens: auto; max-width: 120px; min-width: 60px; }"
            elif max_columns > 5 or (max_columns > 4 and not is_landscape):
                return base_css + "table { font-size: 0.85em; } th, td { padding: 6px 8px; word-wrap: break-word; hyphens: auto; max-width: 150px; }"
            else:
                return base_css + "th, td { word-wrap: break-word; hyphens: auto; }"
        elif table_handling == "smaller_font":
            return base_css + "table { font-size: 0.7em; } th, td { padding: 4px 6px; word-wrap: break-word; hyphens: auto; }"
        else:  # break_words
            return base_css + "table { table-layout: fixed; } th, td { word-wrap: break-word; word-break: break-all; hyphens: auto; overflow-wrap: break-word; }"
    
    def _create_html_body(self, md_text):
        """Processes markdown text to a clean HTML body."""
        # Replace emoji glyphs with SVGs if mapping is available
        # The new replace_glyphs_with_svg function contains all necessary styling,
        # so we no longer need to pass an img_style parameter.
        md_text = replace_glyphs_with_svg(md_text, self.emoji_mapping)
        
        if self.auto_fix_markdown.get():
            md_text = re.sub(r'^(#+)\s*\*\*(.*?)\*\*', r'\1 \2', md_text, flags=re.MULTILINE)
            md_text = self.correct_table_spacing(md_text)
            md_text = self.correct_table_separator_spacing(md_text)
            md_text = self.correct_markdown_table_list_spacing(md_text)
            md_text = self.correct_general_list_and_heading_spacing(md_text)

        enabled_extensions = [name for name, var in self.extensions_config.items() if var.get()]
        html_body = markdown.markdown(md_text, extensions=enabled_extensions)
        
        if self.current_file_path:
            base_dir = os.path.dirname(self.current_file_path)
            html_body = self.process_relative_image_paths(html_body, base_dir)
            
        return html_body
        
    def convert_markdown_to_pdf(self, md_text, output_path):
        html_body = self._create_html_body(md_text)

        table_css = self.get_table_css(html_body)
        
        theme_path = os.path.join(self.themes_dir, self.current_theme.get())
        theme_css = ""
        if os.path.exists(theme_path):
            with open(theme_path, 'r', encoding='utf-8') as f:
                theme_css = f.read()

        html_template = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        {theme_css}
        {table_css}
    </style>
</head>
<body>
{html_body}
</body>
</html>
"""
        temp_html_path = os.path.join(self.output_folder, "temp_output.html")
        with open(temp_html_path, "w", encoding="utf-8") as f:
            f.write(html_template)
            
        try:
            possible_paths = [
                r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',
                r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe',
                'wkhtmltopdf'
            ]
            
            config = None
            for path in possible_paths:
                if path == 'wkhtmltopdf': 
                    try:
                        config = pdfkit.configuration()
                        _ = config.wkhtmltopdf
                        break
                    except OSError:
                        config = None
                        continue
                elif os.path.exists(path):
                    config = pdfkit.configuration(wkhtmltopdf=path)
                    break
            
            if not config:
                raise IOError("Could not find wkhtmltopdf in known locations or PATH.")

            pdf_options = {
                'encoding': "UTF-8",
                'dpi': 600,
                'image-quality': 100,
                'image-dpi': 600,
                'page-size': self.page_size.get(),
                'orientation': self.orientation.get().title(),
                'margin-top': f'{self.margin_top.get()}in',
                'margin-bottom': f'{self.margin_bottom.get()}in',
                'margin-left': f'{self.margin_left.get()}in',
                'margin-right': f'{self.margin_right.get()}in',
                'enable-local-file-access': None,
                'header-left': self.header_left.get() or None,
                'header-center': self.header_center.get() or None,
                'header-right': self.header_right.get() or None,
                'footer-left': self.footer_left.get() or None,
                'footer-center': self.footer_center.get() or None,
                'footer-right': self.footer_right.get() or None,
                'header-font-size': '9',
                'footer-font-size': '9',
                'header-spacing': '5',
                'footer-spacing': '5',
                'print-media-type': None,
                'no-outline': None,
                'disable-smart-shrinking': None,
                'zoom': '1',
                'use-xserver': None,
            }
            
            pdf_options = {k: v for k, v in pdf_options.items() if v is not None}

            if self.generate_toc.get():
                pdfkit.from_file(temp_html_path, output_path, configuration=config, options=pdf_options, toc={})
            else:
                pdfkit.from_file(temp_html_path, output_path, configuration=config, options=pdf_options)

        finally:
            # Always clean up the temporary file.
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
                
    def convert_to_pdf(self):
        md_text = self.text_area.get(1.0, tk.END).strip()
        if not md_text:
            messagebox.showwarning("Warning", "Please enter some markdown text")
            return
        
        filename = self.filename_var.get().strip()
        invalid_chars = r'[<>:"/\\|?*]'
        filename = re.sub(invalid_chars, '_', filename)
        if not filename:
            filename = "document"
        
        if not filename.lower().endswith('.pdf'):
            filename += '.pdf'
        
        output_path = os.path.join(self.output_folder, filename)
        if os.path.exists(output_path):
            if not messagebox.askyesno("Confirm Overwrite", f"The file '{filename}' already exists. Do you want to overwrite it?"):
                self.status_var.set("Conversion cancelled.")
                return
                
        threading.Thread(target=self._convert_thread, args=(md_text, output_path), daemon=True).start()
    
    def _convert_thread(self, md_text, output_path):
        try:
            self.root.after(0, self._start_conversion)
            self.convert_markdown_to_pdf(md_text, output_path)
            self.root.after(0, lambda: self._conversion_complete(output_path))
        except Exception as e:
            self.root.after(0, lambda err=e: self._conversion_error(str(err)))
    
    def _start_conversion(self):
        self.status_var.set("Converting to PDF...")
        self.convert_btn.config(state='disabled')
        self.progress.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 5))
        self.root.rowconfigure(3, weight=0)
        self.progress.start()
    
    def _conversion_complete(self, output_path):
        self.progress.stop()
        self.progress.grid_remove()
        self.root.rowconfigure(3, weight=0) 
        self.convert_btn.config(state='normal')
        self.status_var.set(f"PDF saved successfully: {os.path.basename(output_path)}")
        if messagebox.askyesno("Success", f"PDF created successfully!\n\nWould you like to open it?"):
            os.startfile(output_path)
    
    def _conversion_error(self, error_msg):
        self.progress.stop()
        self.progress.grid_remove()
        self.root.rowconfigure(3, weight=0)
        self.convert_btn.config(state='normal')
        self.status_var.set("Conversion failed")
        
        error_dialog = f"Failed to convert to PDF.\n\nError: {error_msg}\n\n"
        if "wkhtmltopdf" in error_msg.lower():
            error_dialog += "Make sure wkhtmltopdf is installed:\n"
            error_dialog += "1. Download from: https://wkhtmltopdf.org/downloads.html\n"
            error_dialog += "2. Install to default location (or ensure it's in your system PATH)\n"
            error_dialog += "3. Restart this application"
        
        messagebox.showerror("Conversion Error", error_dialog)

def main():
    root = TkinterDnD.Tk()
    app = MarkdownToPDFConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()