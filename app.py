import copy
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
import winreg
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont

from lua_table import LuaParseError, dump_lua_assignment, parse_lua_assignment
try:
    from tkcalendar import Calendar
except Exception:
    Calendar = None


def _default_logbook_path() -> str:
    """Типичный путь к logbook.lua DCS Mission Editor для текущего пользователя (Saved Games)."""
    return str(Path.home() / "Saved Games" / "DCS" / "MissionEditor" / "logbook.lua")


DEFAULT_LOGBOOK_PATH = _default_logbook_path()


def _application_resource_dir() -> Path:
    """
    Каталог с ресурсами (иконки и т.д.) рядом с исходным app.py.
    В EXE от PyInstaller данные распаковываются в sys._MEIPASS, а не рядом с .exe.
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


REGISTRY_KEY = r"Software\Logbooker"
REG_DCS_INSTALL_PATH = "DCSInstallPath"
REG_LOGBOOK_PATH = "LogbookPath"
REG_CAMPAIGN_DOUBLE_CLICK = "CampaignDoubleClick"
# Значения в реестре: "edit" | "folder"
BTN_PAD_X = 4

# ---------------- Icon vertical alignment tuning ----------------
# You can change these values (pixels) to shift each icon up/down inside its own image.
# Positive values move the icon down, negative values move it up.
ICON_Y_OFFSETS_PX = {
    "browse": 0,
    "load": 0,
    "save": 0,
    "profile": 0,
    "settings": 0,
    "edit_campaign": 0,
    "delete_campaign": 0,
    "folder": 0,
    "add_record": 0,
    "duplicate_record": 0,
    "delete_record": 0,
    "calendar": 0,
    "apply_profile": 0,
    "cancel": 0,
}
# ---------------------------------------------------------------


class LogbookEditorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DCS Logbook Editor")
        self.root.geometry("1800x1000")
        self._app_icon_temp_path: Path | None = None
        self._set_app_icon()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self._setup_dark_theme()

        self.file_path_var = tk.StringVar(value=DEFAULT_LOGBOOK_PATH)
        self.status_var = tk.StringVar(value="Выберите файл и нажмите 'Загрузить'.")

        self.assignment_name = "logbook"
        self.logbook_data = None
        self.current_player = None
        self.selected_game_key = None
        self.selected_history_key = None
        self.game_filter_var = tk.StringVar(value="")
        self.history_filter_var = tk.StringVar(value="")
        self.profile_window = None
        self.settings_window = None
        self.player_fields = {}
        self.stats_text = None
        self.dcs_install_path_var = tk.StringVar(value="")
        self.history_field_vars = {}
        self.history_skipped_var = tk.BooleanVar(value=False)
        self.history_form_target_key = None
        self.suppress_history_auto_apply = False
        self.default_games_widths = {}
        self.default_history_widths = {}
        self.max_games_widths = {}
        self.max_history_widths = {}
        self.campaign_double_click_var = tk.StringVar(value="folder")
        self._load_settings_from_registry()
        self._icon_cache = {}
        self._icon_font = None

        self._build_ui()
        self._setup_entry_clipboard_support()
        self.game_filter_var.trace_add("write", self._on_game_filter_change)
        self.history_filter_var.trace_add("write", self._on_history_filter_change)
        self.root.after(0, self.load_file)

    def _setup_dark_theme(self) -> None:
        bg = "#252526"
        panel = "#2d2d30"
        fg = "#d4d4d4"
        accent = "#3e3e42"
        select = "#094771"
        entry_bg = "#333337"
        self.theme_fg = fg
        self.theme_entry_bg = entry_bg

        self.root.configure(bg=bg)
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", background=panel, foreground=fg, fieldbackground=entry_bg)
        style.configure("TFrame", background=panel)
        style.configure("TLabelframe", background=panel, foreground=fg)
        style.configure("TLabelframe.Label", background=panel, foreground=fg)
        style.configure("TLabel", background=panel, foreground=fg)
        style.configure("TButton", background=accent, foreground=fg, bordercolor="#4a4a50")
        style.map("TButton", background=[("active", "#4a4a50"), ("pressed", "#3a3a3f")])
        style.configure("TEntry", fieldbackground=entry_bg, foreground=fg, insertcolor=fg)
        style.configure("TCheckbutton", background=panel, foreground=fg)
        style.map("TCheckbutton", background=[("active", panel)])
        style.configure("TPanedwindow", background=bg)
        style.configure("Treeview", background=entry_bg, foreground=fg, fieldbackground=entry_bg, bordercolor="#4a4a50")
        style.map("Treeview", background=[("selected", select)], foreground=[("selected", "#ffffff")])
        style.configure(
            "Treeview.Heading",
            background="#3f3f46",
            foreground="#f3f3f3",
            bordercolor="#5a5a61",
            relief="flat",
            font=("Segoe UI", 8, "bold"),
            padding=(8, 6),
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#505058"), ("pressed", "#47474f")],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )

    def _get_icon_font(self) -> ImageFont.ImageFont:
        if self._icon_font is not None:
            return self._icon_font

        candidates = ["seguiemj.ttf", "seguiemj.ttf", "segoeui.ttf", "arial.ttf"]
        fonts_dir = Path(r"C:\Windows\Fonts")
        for name in candidates:
            font_path = fonts_dir / name
            if font_path.exists():
                try:
                    self._icon_font = ImageFont.truetype(str(font_path), 16)
                    return self._icon_font
                except Exception:
                    continue
        self._icon_font = ImageFont.load_default()
        return self._icon_font

    def _get_icon_image(self, icon_key: str, size: int = 16) -> ctk.CTkImage:
        cache_key = (icon_key, size)
        if cache_key in self._icon_cache:
            return self._icon_cache[cache_key]

        glyph_map = {
            "browse": "📂",
            "load": "📥",
            "save": "💾",
            "profile": "👤",
            "settings": "⚙",
            "edit_campaign": "✏",
            "delete_campaign": "🗑",
            "folder": "📁",
            "add_record": "➕",
            "duplicate_record": "📑",
            "delete_record": "🗑",
            "calendar": "📅",
            "apply_profile": "✅",
            "cancel": "✕",
        }
        glyph = glyph_map.get(icon_key, "•")

        offset_y = ICON_Y_OFFSETS_PX.get(icon_key, 0)
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        font = self._get_icon_font()

        bbox = draw.textbbox((0, 0), glyph, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x = (size - text_w) // 2 - bbox[0]
        y = (size - text_h) // 2 - bbox[1] + int(offset_y)

        # Use white fill so icons look consistent even if the font is monochrome.
        draw.text((x, y), glyph, font=font, fill=(255, 255, 255, 255))

        icon = ctk.CTkImage(light_image=img, dark_image=img, size=(size, size))
        self._icon_cache[cache_key] = icon
        return icon

    def _setup_entry_clipboard_support(self) -> None:
        self.entry_menu = tk.Menu(self.root, tearoff=0)
        self.entry_menu.add_command(label="Копировать", command=self._copy_focused_entry)
        self.entry_menu.add_command(label="Вставить", command=self._paste_focused_entry)
        self.entry_menu.add_command(label="Вырезать", command=self._cut_focused_entry)
        self.entry_menu.add_separator()
        self.entry_menu.add_command(label="Выделить всё", command=self._select_all_focused_entry)
        self._entry_menu_target = None

        # Class-level bindings are more reliable with CTkEntry internal widgets.
        self.root.bind_class("Entry", "<Control-c>", self._on_copy_shortcut, add="+")
        self.root.bind_class("Entry", "<Control-v>", self._on_paste_shortcut, add="+")
        self.root.bind_class("Entry", "<Control-x>", self._on_cut_shortcut, add="+")
        self.root.bind_class("Entry", "<Control-a>", self._on_select_all_shortcut, add="+")
        self.root.bind_class("Entry", "<Shift-Insert>", self._on_paste_shortcut, add="+")
        self.root.bind_class("Text", "<Control-c>", self._on_copy_shortcut, add="+")
        self.root.bind_class("Text", "<Control-v>", self._on_paste_shortcut, add="+")
        self.root.bind_class("Text", "<Control-x>", self._on_cut_shortcut, add="+")
        self.root.bind_class("Text", "<Control-a>", self._on_select_all_shortcut, add="+")
        self.root.bind_class("Text", "<Shift-Insert>", self._on_paste_shortcut, add="+")
        # Fallback for non-latin keyboard layouts (e.g. Russian).
        self.root.bind_all("<Control-KeyPress>", self._on_control_keypress_fallback, add="+")

        self.root.bind_all("<Button-3>", self._show_entry_context_menu, add="+")

    def _center_toplevel(self, window: tk.Toplevel, width: int | None = None, height: int | None = None) -> None:
        try:
            window.update_idletasks()
            screen_w = window.winfo_screenwidth()
            screen_h = window.winfo_screenheight()

            w = width if width is not None else window.winfo_width()
            h = height if height is not None else window.winfo_height()

            x = int((screen_w - w) / 2)
            y = int((screen_h - h) / 2)
            window.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            return

    def _fit_toplevel_to_content(self, window: tk.Toplevel, min_width: int = 0) -> None:
        """Подгоняет размер окна под содержимое (без лишней пустоты снизу)."""
        try:
            window.update_idletasks()
            w = window.winfo_reqwidth()
            h = window.winfo_reqheight()
            if min_width and w < min_width:
                w = min_width
            self._center_toplevel(window, w, h)
        except Exception:
            return

    def _on_treeview_xscroll(self, scrollbar: ctk.CTkScrollbar, first: str, last: str) -> None:
        """Горизонтальный скролл только при переполнении: при полном диапазоне xview скрываем полосу."""
        scrollbar.set(first, last)
        try:
            f, lt = float(first), float(last)
            if lt - f >= 0.999:
                scrollbar.grid_remove()
            else:
                scrollbar.grid(row=1, column=0, sticky="ew")
        except (ValueError, tk.TclError):
            pass

    def _refresh_treeview_xscroll_visibility(self, tree: ttk.Treeview, scrollbar: ctk.CTkScrollbar) -> None:
        try:
            lo, hi = tree.xview()
            self._on_treeview_xscroll(scrollbar, str(lo), str(hi))
        except tk.TclError:
            pass

    def _toolbar_button_width_for(self, text: str, icon_size: int = 16) -> int:
        """Ширина CTkButton под текст и иконку слева без лишних полей (логические пиксели, как у CTk)."""
        font = ctk.CTkFont()
        try:
            sz = int(font.cget("size"))
        except (TypeError, ValueError):
            sz = 14
        weight = font.cget("weight")
        if weight not in ("normal", "bold"):
            weight = "normal"
        tkf = tkfont.Font(family=font.cget("family"), size=-abs(sz), weight=weight)
        text_w = tkf.measure(text)
        spacing = 6  # CTkButton._image_label_spacing
        pad = 26  # внутренние поля слева/справа (угол, border_spacing)
        w = text_w + icon_size + spacing + pad
        w = max(w, 28)
        return w + (w % 2)

    def _clear_logbook_data(self) -> None:
        """Сброс данных логбука: пустые таблицы, редактор и поля профиля (если уже созданы)."""
        self.logbook_data = None
        self.current_player = None
        self.selected_game_key = None
        self.selected_history_key = None
        self.history_form_target_key = None
        self.assignment_name = "logbook"
        self._refresh_games_tree()
        self._refresh_history_tree()
        self._load_history_form({}, None)
        if self.player_fields:
            for k in self.player_fields:
                self.player_fields[k].set("")
        self._refresh_stats()

    def _prompt_file_not_found(self, path_display: str) -> str:
        """Если logbook.lua не найден: «Указать путь» или «Отмена». Возвращает 'browse' или 'cancel'."""
        result = {"v": "cancel"}

        dlg = tk.Toplevel(self.root)
        dlg.title("Файл не найден")
        dlg.configure(bg="#252526")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.grab_set()

        container = ctk.CTkFrame(dlg, corner_radius=12)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        path_text = path_display.strip() if path_display.strip() else "(путь не указан)"
        msg = (
            "Файл logbook.lua не найден по указанному пути:\n\n"
            f"{path_text}\n\n"
            "Укажите файл вручную или отмените загрузку."
        )
        ctk.CTkLabel(container, text=msg, wraplength=460, justify="left").pack(anchor="w", padx=14, pady=(14, 10))

        btn_row = ctk.CTkFrame(container, corner_radius=8)
        btn_row.pack(fill=tk.X, padx=10, pady=(0, 14))

        def on_browse() -> None:
            result["v"] = "browse"
            dlg.destroy()

        def on_cancel() -> None:
            result["v"] = "cancel"
            dlg.destroy()

        ctk.CTkButton(
            btn_row,
            text="Указать путь к файлу",
            image=self._get_icon_image("browse", 16),
            compound="left",
            width=200,
            corner_radius=8,
            command=on_browse,
        ).pack(side=tk.LEFT, padx=6, pady=8)
        ctk.CTkButton(
            btn_row,
            text="Отмена",
            image=self._get_icon_image("cancel", 16),
            compound="left",
            width=120,
            corner_radius=8,
            command=on_cancel,
        ).pack(side=tk.LEFT, padx=6, pady=8)

        dlg.protocol("WM_DELETE_WINDOW", on_cancel)
        dlg.geometry("500x240")
        self._center_toplevel(dlg, 500, 240)
        dlg.wait_window()
        return result["v"]

    def _get_focused_text_widget(self):
        widget = self.root.focus_get()
        if widget is None:
            widget = self._entry_menu_target
        if widget is None:
            return None
        if hasattr(widget, "get") and hasattr(widget, "insert"):
            return widget
        return None

    def _on_copy_shortcut(self, _event=None):
        return self._copy_focused_entry()

    def _on_paste_shortcut(self, _event=None):
        return self._paste_focused_entry()

    def _on_cut_shortcut(self, _event=None):
        return self._cut_focused_entry()

    def _on_select_all_shortcut(self, _event=None):
        return self._select_all_focused_entry()

    def _on_control_keypress_fallback(self, event):
        # При русской раскладке <Control-c> не срабатывает; keysym часто Cyrillic_es, а не «с».
        # На Windows надёжно опираемся на keycode (физ. клавиши A/C/V/X = VK 65/67/86/88).
        if sys.platform == "win32":
            try:
                vk = int(event.keycode)
            except (TypeError, ValueError):
                vk = None
            if vk is not None:
                if vk == 67:
                    return self._copy_focused_entry()
                if vk == 86:
                    return self._paste_focused_entry()
                if vk == 88:
                    return self._cut_focused_entry()
                if vk == 65:
                    return self._select_all_focused_entry()
        ks = event.keysym or ""
        # Латиница, символы и стандартные keysyms Tk для тех же клавиш (JCUKEN / др.).
        if ks in (
            "c",
            "C",
            "с",
            "С",
            "Cyrillic_es",
            "Cyrillic_ES",
        ):
            return self._copy_focused_entry()
        if ks in (
            "v",
            "V",
            "м",
            "М",
            "Cyrillic_em",
            "Cyrillic_EM",
        ):
            return self._paste_focused_entry()
        if ks in (
            "x",
            "X",
            "ч",
            "Ч",
            "Cyrillic_che",
            "Cyrillic_CHE",
        ):
            return self._cut_focused_entry()
        if ks in (
            "a",
            "A",
            "ф",
            "Ф",
            "Cyrillic_ef",
            "Cyrillic_EF",
        ):
            return self._select_all_focused_entry()
        return

    def _copy_focused_entry(self):
        widget = self._get_focused_text_widget()
        if widget is None:
            return
        try:
            if widget.winfo_class() == "Text":
                text = widget.get("sel.first", "sel.last")
            else:
                text = widget.selection_get()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            return "break"
        except Exception:
            return

    def _paste_focused_entry(self):
        widget = self._get_focused_text_widget()
        if widget is None:
            return
        try:
            text = self.root.clipboard_get()
        except Exception:
            return
        try:
            if widget.winfo_class() == "Text":
                if widget.tag_ranges("sel"):
                    widget.delete("sel.first", "sel.last")
                widget.insert("insert", text)
            else:
                try:
                    widget.delete("sel.first", "sel.last")
                except Exception:
                    pass
                widget.insert("insert", text)
            return "break"
        except Exception:
            return

    def _cut_focused_entry(self):
        widget = self._get_focused_text_widget()
        if widget is None:
            return
        self._copy_focused_entry()
        try:
            if widget.winfo_class() == "Text":
                widget.delete("sel.first", "sel.last")
            else:
                widget.delete("sel.first", "sel.last")
            return "break"
        except Exception:
            return

    def _select_all_focused_entry(self):
        widget = self._get_focused_text_widget()
        if widget is None:
            return
        try:
            if widget.winfo_class() == "Text":
                widget.tag_add("sel", "1.0", "end-1c")
            else:
                widget.selection_range(0, "end")
                if hasattr(widget, "icursor"):
                    widget.icursor("end")
            return "break"
        except Exception:
            return

    def _show_entry_context_menu(self, event):
        widget = event.widget
        if widget is None:
            return
        if hasattr(widget, "event_generate"):
            self._entry_menu_target = widget
            self.entry_menu.tk_popup(event.x_root, event.y_root)
            self.entry_menu.grab_release()

    def _entry_event(self, virtual_event: str) -> None:
        widget = self._entry_menu_target if self._entry_menu_target else self.root.focus_get()
        if widget is None:
            return
        try:
            widget.event_generate(virtual_event)
        except Exception:
            return

    def _build_ui(self) -> None:
        top = ctk.CTkFrame(self.root, corner_radius=10)
        top.pack(fill=tk.X, padx=8, pady=(8, 0))

        top.grid_columnconfigure(1, weight=1)
        top_file_label = ctk.CTkLabel(top, text="Файл logbook.lua:", cursor="hand2", text_color="#8bb9ff")
        top_file_label.grid(row=0, column=0, padx=(10, 6), pady=8, sticky="w")
        top_file_label.bind("<Button-1>", self.open_logbook_file)
        top_file_label.bind("<Enter>", lambda _e: top_file_label.configure(text_color="#b9d6ff"))
        top_file_label.bind("<Leave>", lambda _e: top_file_label.configure(text_color="#8bb9ff"))
        ctk.CTkEntry(top, textvariable=self.file_path_var, corner_radius=8).grid(row=0, column=1, padx=(0, 10), pady=8, sticky="ew")

        top_buttons = ctk.CTkFrame(top, fg_color="transparent")
        top_buttons.grid(row=0, column=2, padx=(0, 10), pady=8, sticky="e")
        ctk.CTkButton(
            top_buttons,
            text="Обзор",
            image=self._get_icon_image("browse", 16),
            compound="left",
            width=90,
            corner_radius=8,
            command=self.browse_file,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X)
        ctk.CTkButton(
            top_buttons,
            text="Загрузить",
            image=self._get_icon_image("load", 16),
            compound="left",
            width=100,
            corner_radius=8,
            command=self.load_file,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X)
        ctk.CTkButton(
            top_buttons,
            text="Сохранить",
            image=self._get_icon_image("save", 16),
            compound="left",
            width=self._toolbar_button_width_for("Сохранить"),
            corner_radius=8,
            command=self.save_file,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X)
        ctk.CTkButton(
            top_buttons,
            text="Профиль игрока",
            image=self._get_icon_image("profile", 16),
            compound="left",
            width=self._toolbar_button_width_for("Профиль игрока"),
            corner_radius=8,
            command=self.open_profile_window,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X)
        ctk.CTkButton(
            top_buttons,
            text="Настройки",
            image=self._get_icon_image("settings", 16),
            compound="left",
            width=120,
            corner_radius=8,
            command=self.open_settings_window,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X)

        content = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        content.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.content_pane = content

        campaigns_col = ttk.Frame(content, padding=6)
        missions_col = ttk.Frame(content, padding=6)
        editor_col = ttk.Frame(content, padding=6)
        content.add(campaigns_col, weight=3)
        content.add(missions_col, weight=3)
        content.add(editor_col, weight=4)
        self.root.after(100, self._set_initial_editor_width)

        ctk.CTkLabel(campaigns_col, text="Кампании игрока").pack(anchor=tk.W)
        game_filter_row = ctk.CTkFrame(campaigns_col, corner_radius=8)
        game_filter_row.pack(fill=tk.X, pady=(2, 0))
        ctk.CTkLabel(game_filter_row, text="Поиск по имени .cmp:").pack(side=tk.LEFT, padx=8, pady=6)
        game_search_box = ctk.CTkFrame(game_filter_row, corner_radius=8, fg_color=self.theme_entry_bg)
        game_search_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4, pady=6)
        ctk.CTkEntry(game_search_box, textvariable=self.game_filter_var, corner_radius=8, border_width=0, fg_color=self.theme_entry_bg).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0), pady=2
        )
        ctk.CTkButton(
            game_search_box,
            text="✕",
            width=24,
            height=24,
            corner_radius=6,
            fg_color="#4a4a4f",
            hover_color="#5a5a60",
            text_color="#b0b0b5",
            command=self.clear_game_filter,
        ).pack(side=tk.RIGHT, padx=4, pady=2)

        games_cols = ("id", "campaign", "status", "created")
        games_tree_wrap = ttk.Frame(campaigns_col)
        games_tree_wrap.pack(fill=tk.BOTH, expand=True, pady=6)
        self.games_tree = ttk.Treeview(games_tree_wrap, columns=games_cols, show="headings", height=20)
        for col, txt, w in (
            ("id", "ID", 60),
            ("campaign", "Campaign", 220),
            ("status", "Status", 100),
            ("created", "Created", 170),
        ):
            self.games_tree.heading(col, text=txt)
            self.games_tree.column(col, width=w, minwidth=60, anchor=tk.W, stretch=False)
            self.default_games_widths[col] = w
        self.max_games_widths = {"id": 80, "campaign": 340, "status": 140, "created": 220}
        games_y_scroll = ctk.CTkScrollbar(
            games_tree_wrap,
            orientation="vertical",
            command=self.games_tree.yview,
            corner_radius=8,
            width=16,
            fg_color="#1f3044",
            button_color="#3b8ed0",
            button_hover_color="#36719f",
        )
        games_x_scroll = ctk.CTkScrollbar(
            games_tree_wrap,
            orientation="horizontal",
            command=self.games_tree.xview,
            corner_radius=8,
            height=16,
            fg_color="#1f3044",
            button_color="#3b8ed0",
            button_hover_color="#36719f",
        )
        self.games_tree.configure(
            yscrollcommand=games_y_scroll.set,
            xscrollcommand=lambda a, b, sx=games_x_scroll: self._on_treeview_xscroll(sx, a, b),
        )
        self.games_tree.grid(row=0, column=0, sticky="nsew")
        games_y_scroll.grid(row=0, column=1, sticky="ns")
        games_x_scroll.grid(row=1, column=0, sticky="ew")
        games_x_scroll.grid_remove()
        games_tree_wrap.rowconfigure(0, weight=1)
        games_tree_wrap.columnconfigure(0, weight=1)
        self._games_x_scroll = games_x_scroll
        self.games_tree.bind("<Configure>", lambda _e: self._refresh_treeview_xscroll_visibility(self.games_tree, games_x_scroll), add="+")
        self.root.after_idle(lambda: self._refresh_treeview_xscroll_visibility(self.games_tree, games_x_scroll))
        self.games_tree.bind("<<TreeviewSelect>>", self.on_game_select)
        self.games_tree.bind(
            "<Double-1>",
            lambda e: self._on_tree_separator_double_click(
                self.games_tree, e, self.default_games_widths, self.max_games_widths
            ),
        )
        self.games_tree.bind("<Double-Button-1>", self.on_game_double_click, add="+")

        game_buttons = ctk.CTkFrame(campaigns_col, corner_radius=8)
        game_buttons.pack(fill=tk.X, pady=(0, 6))
        ctk.CTkButton(
            game_buttons,
            text="Редактировать кампанию",
            image=self._get_icon_image("edit_campaign", 16),
            compound="left",
            corner_radius=8,
            command=self.edit_game,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)
        ctk.CTkButton(
            game_buttons,
            text="Удалить кампанию",
            image=self._get_icon_image("delete_campaign", 16),
            compound="left",
            corner_radius=8,
            command=self.delete_game,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)
        ctk.CTkButton(
            game_buttons,
            text="Открыть в папке",
            image=self._get_icon_image("folder", 16),
            compound="left",
            corner_radius=8,
            command=self.open_campaign_folder,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)

        ctk.CTkLabel(missions_col, text="История выбранной кампании").pack(anchor=tk.W)
        history_filter_row = ctk.CTkFrame(missions_col, corner_radius=8)
        history_filter_row.pack(fill=tk.X, pady=(2, 0))
        ctk.CTkLabel(history_filter_row, text="Поиск в истории:").pack(side=tk.LEFT, padx=8, pady=6)
        history_search_box = ctk.CTkFrame(history_filter_row, corner_radius=8, fg_color=self.theme_entry_bg)
        history_search_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4, pady=6)
        ctk.CTkEntry(
            history_search_box,
            textvariable=self.history_filter_var,
            corner_radius=8,
            border_width=0,
            fg_color=self.theme_entry_bg,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0), pady=2)
        ctk.CTkButton(
            history_search_box,
            text="✕",
            width=24,
            height=24,
            corner_radius=6,
            fg_color="#4a4a4f",
            hover_color="#5a5a60",
            text_color="#b0b0b5",
            command=self.clear_history_filter,
        ).pack(side=tk.RIGHT, padx=4, pady=2)
        hist_cols = ("id", "stage", "mission", "result", "aa", "ag", "deaths", "datetime", "skipped")
        history_tree_wrap = ttk.Frame(missions_col)
        history_tree_wrap.pack(fill=tk.BOTH, expand=True, pady=6)
        self.history_tree = ttk.Treeview(history_tree_wrap, columns=hist_cols, show="headings", height=20)
        for col, txt, w in (
            ("id", "ID", 45),
            ("stage", "Stage", 55),
            ("mission", "Mission", 120),
            ("result", "Result", 55),
            ("aa", "AA", 45),
            ("ag", "AG", 45),
            ("deaths", "Deaths", 60),
            ("datetime", "Datetime", 180),
            ("skipped", "Пропущено", 70),
        ):
            self.history_tree.heading(col, text=txt)
            self.history_tree.column(col, width=w, minwidth=50, anchor=tk.W, stretch=False)
            self.default_history_widths[col] = w
        self.max_history_widths = {
            "id": 70,
            "stage": 80,
            "result": 90,
            "aa": 70,
            "ag": 70,
            "deaths": 90,
            "datetime": 240,
            "skipped": 90,
        }
        history_y_scroll = ctk.CTkScrollbar(
            history_tree_wrap,
            orientation="vertical",
            command=self.history_tree.yview,
            corner_radius=8,
            width=16,
            fg_color="#1f3044",
            button_color="#3b8ed0",
            button_hover_color="#36719f",
        )
        history_x_scroll = ctk.CTkScrollbar(
            history_tree_wrap,
            orientation="horizontal",
            command=self.history_tree.xview,
            corner_radius=8,
            height=16,
            fg_color="#1f3044",
            button_color="#3b8ed0",
            button_hover_color="#36719f",
        )
        self.history_tree.configure(
            yscrollcommand=history_y_scroll.set,
            xscrollcommand=lambda a, b, sx=history_x_scroll: self._on_treeview_xscroll(sx, a, b),
        )
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        history_y_scroll.grid(row=0, column=1, sticky="ns")
        history_x_scroll.grid(row=1, column=0, sticky="ew")
        history_x_scroll.grid_remove()
        history_tree_wrap.rowconfigure(0, weight=1)
        history_tree_wrap.columnconfigure(0, weight=1)
        self._history_x_scroll = history_x_scroll
        self.history_tree.bind("<Configure>", lambda _e: self._refresh_treeview_xscroll_visibility(self.history_tree, history_x_scroll), add="+")
        self.root.after_idle(lambda: self._refresh_treeview_xscroll_visibility(self.history_tree, history_x_scroll))
        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)
        self.history_tree.bind(
            "<Double-1>",
            lambda e: self._on_tree_separator_double_click(
                self.history_tree, e, self.default_history_widths, self.max_history_widths
            ),
        )

        history_buttons = ctk.CTkFrame(missions_col, corner_radius=8)
        history_buttons.pack(fill=tk.X, pady=(0, 6))
        ctk.CTkButton(
            history_buttons,
            text="Добавить запись",
            image=self._get_icon_image("add_record", 16),
            compound="left",
            corner_radius=8,
            command=self.add_empty_history_entry,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)
        ctk.CTkButton(
            history_buttons,
            text="Дублировать запись",
            image=self._get_icon_image("duplicate_record", 16),
            compound="left",
            corner_radius=8,
            command=self.duplicate_history_entry,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)
        ctk.CTkButton(
            history_buttons,
            text="Удалить запись",
            image=self._get_icon_image("delete_record", 16),
            compound="left",
            corner_radius=8,
            command=self.delete_history_item,
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)

        editor_box = ctk.CTkFrame(editor_col, corner_radius=12)
        editor_box.pack(fill=tk.BOTH, expand=True)
        ctk.CTkLabel(editor_box, text="Редактор миссии").pack(anchor=tk.W, padx=10, pady=(10, 4))

        field_specs = (
            ("stage", "Stage"),
            ("mission", "Mission"),
            ("datetime", "Datetime"),
            ("result", "Result"),
            ("aaKills", "AA Kills"),
            ("agKills", "AG Kills"),
            ("deathsCount", "Deaths"),
        )
        for key, title in field_specs:
            row = ctk.CTkFrame(editor_box, corner_radius=8)
            row.pack(fill=tk.X, pady=3)
            ctk.CTkLabel(row, text=title, width=90).pack(side=tk.LEFT, padx=8, pady=6)
            var = tk.StringVar(value="")
            self.history_field_vars[key] = var
            var.trace_add("write", self._apply_history_form_live)
            ctk.CTkEntry(row, textvariable=var, corner_radius=8).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), pady=6)
            if key == "datetime":
                ctk.CTkButton(
                    row,
                    text="",
                    image=self._get_icon_image("calendar", 16),
                    compound="left",
                    width=34,
                    corner_radius=8,
                    command=self.open_datetime_picker,
                ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=6)

        skip_row = ctk.CTkFrame(editor_box, corner_radius=8)
        skip_row.pack(fill=tk.X, pady=4)
        ctk.CTkCheckBox(skip_row, text="Пропущено", variable=self.history_skipped_var).pack(side=tk.LEFT, padx=8, pady=6)
        self.history_skipped_var.trace_add("write", self._apply_history_form_live)

        bottom = ctk.CTkFrame(self.root, corner_radius=10)
        bottom.pack(fill=tk.X, padx=8, pady=(0, 8))
        ctk.CTkLabel(bottom, textvariable=self.status_var, anchor="w").pack(fill=tk.X, padx=10, pady=8)

    def browse_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Выберите logbook.lua",
            filetypes=[("Lua files", "*.lua"), ("All files", "*.*")],
            initialfile="logbook.lua",
        )
        if selected:
            self.file_path_var.set(selected)
            self._persist_logbook_path_to_registry()

    def open_logbook_file(self, _event=None) -> None:
        path = Path(self.file_path_var.get().strip())
        if not path.exists():
            messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
            return
        try:
            os.startfile(str(path))
        except Exception as exc:
            messagebox.showerror("Ошибка открытия файла", str(exc))

    def load_file(self) -> None:
        path = Path(self.file_path_var.get().strip())
        if not path.exists():
            action = self._prompt_file_not_found(str(path))
            if action == "cancel":
                self._clear_logbook_data()
                self.file_path_var.set("")
                self._persist_logbook_path_to_registry()
                self.status_var.set("Файл не загружен. Укажите путь к logbook.lua кнопкой «Загрузить» или «Обзор».")
                return
            selected = filedialog.askopenfilename(
                title="Выберите logbook.lua",
                filetypes=[("Lua files", "*.lua"), ("All files", "*.*")],
                initialfile="logbook.lua",
            )
            if not selected:
                self.status_var.set("Файл не выбран. Укажите путь к logbook.lua.")
                return
            self.file_path_var.set(selected)
            self._persist_logbook_path_to_registry()
            path = Path(self.file_path_var.get().strip())
            if not path.exists():
                self.load_file()
                return
        try:
            content = path.read_text(encoding="utf-8", errors="replace")
            name, data = parse_lua_assignment(content)
        except LuaParseError as exc:
            messagebox.showerror("Ошибка парсинга", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Ошибка чтения", str(exc))
            return

        self.assignment_name = name
        self.logbook_data = data

        players = self.logbook_data.get("players", {})
        if 1 not in players:
            messagebox.showerror("Ошибка", "В файле нет players[1].")
            return
        self.current_player = players[1]
        self.selected_game_key = None
        self.selected_history_key = None

        self._fill_profile_fields()
        self._refresh_games_tree()
        self._refresh_history_tree()
        self._refresh_stats()
        self.status_var.set(f"Загружено: {path}")
        self._persist_logbook_path_to_registry()

    def _fill_profile_fields(self) -> None:
        if not self.current_player:
            return
        self._ensure_profile_vars()
        for key in ("name", "callsign", "squadron", "rank", "rankName"):
            self.player_fields[key].set(str(self.current_player.get(key, "")))
        self.player_fields["currentPlayerName"].set(str(self.logbook_data.get("currentPlayerName", "")))

    def apply_profile_changes(self) -> None:
        if not self.current_player:
            return
        for key in ("name", "callsign", "squadron", "rank", "rankName"):
            self.current_player[key] = self.player_fields[key].get()
        self.logbook_data["currentPlayerName"] = self.player_fields["currentPlayerName"].get()
        self.status_var.set("Изменения профиля применены в памяти. Не забудьте сохранить файл.")

    def _refresh_games_tree(self) -> None:
        for item in self.games_tree.get_children():
            self.games_tree.delete(item)
        games = self.current_player.get("games", {}) if self.current_player else {}
        filter_value = self.game_filter_var.get().strip().lower()
        for gid in sorted(games.keys()):
            game = games[gid]
            cmp_name = self._campaign_cmp_filename_for_filter(str(game.get("campaign", "")))
            if filter_value:
                if not cmp_name or filter_value not in cmp_name:
                    continue
            self.games_tree.insert(
                "",
                tk.END,
                iid=str(gid),
                values=(
                    gid,
                    self._campaign_display_name(str(game.get("campaign", ""))),
                    game.get("status", ""),
                    game.get("created", ""),
                ),
            )
        self._autosize_tree_columns(self.games_tree, self.default_games_widths, self.max_games_widths)
        self.root.after_idle(lambda: self._refresh_treeview_xscroll_visibility(self.games_tree, self._games_x_scroll))

    def _campaign_display_name(self, campaign_path: str) -> str:
        normalized = campaign_path.replace("\\", "/")
        return normalized.rsplit("/", 1)[-1] if normalized else ""

    def _campaign_cmp_filename_for_filter(self, campaign_path: str) -> str:
        """Имя файла .cmp (нижний регистр) для поиска; пустая строка, если в пути нет имени с расширением .cmp."""
        raw = str(campaign_path).strip() if campaign_path else ""
        if not raw:
            return ""
        base = raw.replace("\\", "/").rsplit("/", 1)[-1]
        if base.lower().endswith(".cmp"):
            return base.lower()
        return ""

    def _refresh_history_tree(self) -> None:
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        if self.selected_game_key is None:
            self.root.after_idle(lambda: self._refresh_treeview_xscroll_visibility(self.history_tree, self._history_x_scroll))
            return
        game = self.current_player["games"].get(self.selected_game_key, {})
        history = game.get("history", {})
        filter_value = self.history_filter_var.get().strip().lower()
        for hid in sorted(history.keys()):
            item = history[hid]
            search_blob = " ".join(
                [
                    str(hid),
                    str(item.get("stage", "")),
                    str(item.get("mission", "")),
                    str(item.get("datetime", "")),
                    str(item.get("result", "")),
                    str(item.get("aaKills", "")),
                    str(item.get("agKills", "")),
                    str(item.get("deathsCount", "")),
                    str(item.get("skipped", "")),
                ]
            ).lower()
            if filter_value and filter_value not in search_blob:
                continue
            self.history_tree.insert(
                "",
                tk.END,
                iid=str(hid),
                values=(
                    hid,
                    item.get("stage", ""),
                    item.get("mission", ""),
                    item.get("result", ""),
                    item.get("aaKills", ""),
                    item.get("agKills", ""),
                    item.get("deathsCount", ""),
                    item.get("datetime", ""),
                    item.get("skipped", False),
                ),
            )
        self._autosize_tree_columns(self.history_tree, self.default_history_widths, self.max_history_widths)
        self.root.after_idle(lambda: self._refresh_treeview_xscroll_visibility(self.history_tree, self._history_x_scroll))

    def clear_game_filter(self) -> None:
        self.game_filter_var.set("")
        self._refresh_games_tree()

    def clear_history_filter(self) -> None:
        self.history_filter_var.set("")
        self._refresh_history_tree()

    def _on_game_filter_change(self, *_args) -> None:
        self._refresh_games_tree()

    def _on_history_filter_change(self, *_args) -> None:
        self._refresh_history_tree()

    def _refresh_stats(self) -> None:
        if self.stats_text is None:
            return
        self.stats_text.delete("1.0", tk.END)
        if not self.current_player:
            return
        stats = self.current_player.get("statistics", {})
        keys = ["missionsCount", "campaignsCount", "flightHours", "aaKills", "agKills", "deaths", "totalScore", "killRatio"]
        lines = []
        for key in keys:
            if key in stats:
                lines.append(f"{key}: {stats[key]}")
        self.stats_text.insert("1.0", "\n".join(lines) if lines else "Нет сводной статистики.")

    def _ensure_profile_vars(self) -> None:
        if self.player_fields:
            return
        for field in ("name", "callsign", "squadron", "rank", "rankName", "currentPlayerName"):
            self.player_fields[field] = tk.StringVar(value="")

    def open_profile_window(self) -> None:
        if self.profile_window is not None and self.profile_window.winfo_exists():
            self.profile_window.lift()
            self.profile_window.focus_force()
            return

        self._ensure_profile_vars()
        self.profile_window = tk.Toplevel(self.root)
        self.profile_window.title("Профиль игрока и статистика")
        self.profile_window.geometry("650x640")
        self.profile_window.configure(bg="#252526")
        self._center_toplevel(self.profile_window, 650, 640)

        container = ctk.CTkFrame(self.profile_window, corner_radius=12)
        container.pack(fill=tk.BOTH, expand=True)

        profile = ctk.CTkFrame(container, corner_radius=10)
        profile.pack(fill=tk.X, pady=(10, 8), padx=10)
        ctk.CTkLabel(profile, text="Основные поля").pack(anchor=tk.W, padx=10, pady=(8, 2))
        for field in ("name", "callsign", "squadron", "rank", "rankName", "currentPlayerName"):
            row = ctk.CTkFrame(profile, corner_radius=8)
            row.pack(fill=tk.X, pady=2)
            ctk.CTkLabel(row, text=field, width=130).pack(side=tk.LEFT, padx=8, pady=6)
            ctk.CTkEntry(row, textvariable=self.player_fields[field], corner_radius=8).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), pady=6)

        ctk.CTkButton(
            profile,
            text="Применить изменения профиля",
            image=self._get_icon_image("apply_profile", 16),
            compound="left",
            corner_radius=8,
            command=self.apply_profile_changes,
        ).pack(anchor=tk.E, pady=8, padx=8)

        stats_box = ctk.CTkFrame(container, corner_radius=10)
        stats_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        ctk.CTkLabel(stats_box, text="Сводная статистика").pack(anchor=tk.W, padx=10, pady=(8, 2))
        self.stats_text = tk.Text(stats_box, wrap="word", height=24)
        self.stats_text.configure(bg=self.theme_entry_bg, fg=self.theme_fg, insertbackground=self.theme_fg, relief=tk.FLAT, highlightthickness=0)
        self.stats_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        def on_close() -> None:
            self.stats_text = None
            self.profile_window.destroy()
            self.profile_window = None

        self.profile_window.protocol("WM_DELETE_WINDOW", on_close)
        self._fill_profile_fields()
        self._refresh_stats()

    def _load_settings_from_registry(self) -> None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY, 0, winreg.KEY_READ) as key:
                try:
                    value, _ = winreg.QueryValueEx(key, REG_DCS_INSTALL_PATH)
                    self.dcs_install_path_var.set(str(value))
                except OSError:
                    self.dcs_install_path_var.set("")
                try:
                    lp, _ = winreg.QueryValueEx(key, REG_LOGBOOK_PATH)
                    p = str(lp).strip()
                    if p:
                        self.file_path_var.set(p)
                except OSError:
                    pass
                try:
                    dc, _ = winreg.QueryValueEx(key, REG_CAMPAIGN_DOUBLE_CLICK)
                    v = str(dc).strip().lower()
                    if v in ("edit", "folder"):
                        self.campaign_double_click_var.set(v)
                except OSError:
                    pass
        except OSError:
            self.dcs_install_path_var.set("")

    def _save_settings_to_registry(self) -> None:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY) as key:
            winreg.SetValueEx(key, REG_DCS_INSTALL_PATH, 0, winreg.REG_SZ, self.dcs_install_path_var.get().strip())
            winreg.SetValueEx(key, REG_LOGBOOK_PATH, 0, winreg.REG_SZ, self.file_path_var.get().strip())
            v = self.campaign_double_click_var.get().strip().lower()
            if v not in ("edit", "folder"):
                v = "folder"
            winreg.SetValueEx(key, REG_CAMPAIGN_DOUBLE_CLICK, 0, winreg.REG_SZ, v)

    def _persist_logbook_path_to_registry(self) -> None:
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY) as key:
                winreg.SetValueEx(key, REG_LOGBOOK_PATH, 0, winreg.REG_SZ, self.file_path_var.get().strip())
        except OSError:
            pass

    def delete_saved_paths_from_registry(self) -> None:
        if not messagebox.askyesno(
            "Подтвердите",
            "Удалить из реестра сохранённые пути к logbook.lua и к установке DCS?\n\n"
            "Поля в этом окне будут сброшены: путь к логбуку — стандартный для текущего пользователя, путь DCS — пустой.",
        ):
            return
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REGISTRY_KEY, 0, winreg.KEY_WRITE) as key:
                try:
                    winreg.DeleteValue(key, REG_DCS_INSTALL_PATH)
                except OSError:
                    pass
                try:
                    winreg.DeleteValue(key, REG_LOGBOOK_PATH)
                except OSError:
                    pass
        except OSError:
            pass
        self.dcs_install_path_var.set("")
        self.file_path_var.set(_default_logbook_path())
        self.status_var.set("Сохранённые пути удалены из реестра.")

    def open_settings_window(self) -> None:
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.settings_window.focus_force()
            return

        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("Настройки")
        self.settings_window.configure(bg="#252526")

        container = ctk.CTkFrame(self.settings_window, corner_radius=12)
        container.pack(fill=tk.X, expand=False, padx=10, pady=10)

        row_logbook = ctk.CTkFrame(container, corner_radius=8)
        row_logbook.pack(fill=tk.X, padx=10, pady=(14, 8))
        ctk.CTkLabel(row_logbook, text="Путь к logbook.lua:").pack(side=tk.LEFT, padx=8, pady=8)
        ctk.CTkEntry(row_logbook, textvariable=self.file_path_var, corner_radius=8).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=8
        )
        ctk.CTkButton(
            row_logbook,
            text="Обзор",
            image=self._get_icon_image("browse", 16),
            compound="left",
            width=90,
            corner_radius=8,
            command=self.browse_file,
        ).pack(side=tk.LEFT, padx=6, pady=8)

        row = ctk.CTkFrame(container, corner_radius=8)
        row.pack(fill=tk.X, padx=10, pady=(0, 8))
        ctk.CTkLabel(row, text="Путь установки DSC world:").pack(side=tk.LEFT, padx=8, pady=8)
        ctk.CTkEntry(row, textvariable=self.dcs_install_path_var, corner_radius=8).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=8)
        ctk.CTkButton(
            row,
            text="Обзор",
            image=self._get_icon_image("browse", 16),
            compound="left",
            width=90,
            corner_radius=8,
            command=self.browse_dcs_install_path,
        ).pack(side=tk.LEFT, padx=6, pady=8)

        row_dclick = ctk.CTkFrame(container, corner_radius=8)
        row_dclick.pack(fill=tk.X, padx=10, pady=(0, 8))
        ctk.CTkLabel(row_dclick, text="Двойной клик по строке кампании:").pack(anchor="w", padx=8, pady=(4, 2))
        dclick_opts = ctk.CTkFrame(row_dclick, fg_color="transparent")
        dclick_opts.pack(fill=tk.X, padx=8, pady=(0, 6))
        ctk.CTkRadioButton(
            dclick_opts,
            text="Открывать редактирование",
            variable=self.campaign_double_click_var,
            value="edit",
        ).pack(anchor="w", pady=2)
        ctk.CTkRadioButton(
            dclick_opts,
            text="Открывать папку кампании",
            variable=self.campaign_double_click_var,
            value="folder",
        ).pack(anchor="w", pady=2)

        actions = ctk.CTkFrame(container, corner_radius=8)
        actions.pack(fill=tk.X, padx=10, pady=8)
        ctk.CTkButton(
            actions,
            text="Удалить сохранённые пути из реестра",
            image=self._get_icon_image("delete_campaign", 16),
            compound="left",
            width=300,
            corner_radius=8,
            command=self.delete_saved_paths_from_registry,
        ).pack(side=tk.LEFT, padx=6, pady=8)
        ctk.CTkButton(
            actions,
            text="Сохранить",
            image=self._get_icon_image("save", 16),
            compound="left",
            width=120,
            corner_radius=8,
            command=self.save_settings,
        ).pack(side=tk.RIGHT, padx=6, pady=8)
        ctk.CTkButton(
            actions,
            text="Отмена",
            image=self._get_icon_image("cancel", 16),
            compound="left",
            width=120,
            corner_radius=8,
            command=self.settings_window.destroy,
        ).pack(side=tk.RIGHT, padx=6, pady=8)

        def on_close() -> None:
            self.settings_window.destroy()
            self.settings_window = None

        self.settings_window.protocol("WM_DELETE_WINDOW", on_close)
        self._fit_toplevel_to_content(self.settings_window, min_width=760)

    def browse_dcs_install_path(self) -> None:
        selected = filedialog.askdirectory(title="Выберите папку установки DCS World")
        if selected:
            self.dcs_install_path_var.set(selected)

    def save_settings(self) -> None:
        try:
            self._save_settings_to_registry()
            self.status_var.set("Настройки сохранены в реестр Windows.")
            if self.settings_window and self.settings_window.winfo_exists():
                self.settings_window.destroy()
                self.settings_window = None
        except Exception as exc:
            messagebox.showerror("Ошибка сохранения настроек", str(exc))

    def on_game_select(self, _event=None) -> None:
        sel = self.games_tree.selection()
        if not sel:
            self.selected_game_key = None
            self.selected_history_key = None
            self._refresh_history_tree()
            return
        self.selected_game_key = int(sel[0])
        self.selected_history_key = None
        self._refresh_history_tree()

    def on_game_double_click(self, _event=None) -> None:
        if self.selected_game_key is None:
            return
        mode = self.campaign_double_click_var.get().strip().lower()
        if mode == "edit":
            self.edit_game()
        else:
            self.open_campaign_folder()

    def on_history_select(self, _event=None) -> None:
        sel = self.history_tree.selection()
        self.selected_history_key = int(sel[0]) if sel else None
        if self.selected_history_key is None or self.selected_game_key is None:
            return
        history = self.current_player["games"][self.selected_game_key].setdefault("history", {})
        item = history.get(self.selected_history_key)
        if item is not None:
            self._load_history_form(item, target_key=self.selected_history_key)

    def _next_int_key(self, mapping: dict) -> int:
        keys = [k for k in mapping.keys() if isinstance(k, int)]
        return (max(keys) + 1) if keys else 1

    def edit_game(self) -> None:
        if self.selected_game_key is None:
            messagebox.showinfo("Инфо", "Сначала выберите кампанию.")
            return
        game = self.current_player["games"][self.selected_game_key]
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Редактирование кампании {self.selected_game_key}")
        dialog.configure(bg="#252526")
        dialog.transient(self.root)
        dialog.grab_set()

        container = ctk.CTkFrame(dialog, corner_radius=12)
        container.pack(fill=tk.X, expand=False, padx=10, pady=10)

        field_vars = {}
        for field in ("player", "campaign", "status", "created"):
            row = ctk.CTkFrame(container, corner_radius=8)
            row.pack(fill=tk.X, padx=10, pady=4)
            ctk.CTkLabel(row, text=field, width=110).pack(side=tk.LEFT, padx=8, pady=8)
            var = tk.StringVar(value=str(game.get(field, "")))
            field_vars[field] = var
            ctk.CTkEntry(row, textvariable=var, corner_radius=8).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), pady=8)
            if field == "created":
                ctk.CTkButton(
                    row,
                    text="",
                    image=self._get_icon_image("calendar", 16),
                    compound="left",
                    width=34,
                    corner_radius=8,
                    command=lambda v=var: self.open_datetime_picker_for_var(v, dialog),
                ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=8)

        btn_row = ctk.CTkFrame(container, corner_radius=8)
        btn_row.pack(fill=tk.X, padx=10, pady=(8, 10))

        ctk.CTkButton(
            btn_row,
            text="Открыть в папке",
            image=self._get_icon_image("folder", 16),
            compound="left",
            corner_radius=8,
            command=lambda: self.open_campaign_folder(field_vars["campaign"].get()),
        ).pack(side=tk.LEFT, padx=BTN_PAD_X, pady=8)

        def save_game_edit() -> None:
            for field in ("player", "campaign", "status", "created"):
                game[field] = field_vars[field].get()
            self._refresh_games_tree()
            self.status_var.set(f"Кампания {self.selected_game_key} обновлена (в памяти).")
            dialog.destroy()

        ctk.CTkButton(
            btn_row,
            text="Сохранить",
            image=self._get_icon_image("save", 16),
            compound="left",
            width=120,
            corner_radius=8,
            command=save_game_edit,
        ).pack(side=tk.RIGHT, padx=BTN_PAD_X, pady=8)
        ctk.CTkButton(
            btn_row,
            text="Отмена",
            image=self._get_icon_image("cancel", 16),
            compound="left",
            width=120,
            corner_radius=8,
            command=dialog.destroy,
        ).pack(side=tk.RIGHT, padx=BTN_PAD_X, pady=8)

        self._fit_toplevel_to_content(dialog, min_width=760)

    def open_datetime_picker_for_var(self, target_var: tk.StringVar, parent_window=None) -> None:
        if Calendar is None:
            messagebox.showerror(
                "Календарь недоступен",
                "Не найден модуль tkcalendar.\nУстановите:\npython -m pip install tkcalendar",
            )
            return

        current_value = target_var.get().strip()
        base_dt = self._parse_datetime_value(current_value) or datetime.now()

        picker = tk.Toplevel(parent_window if parent_window else self.root)
        picker.title("Выбор даты и времени")
        picker.geometry("320x360")
        picker.transient(parent_window if parent_window else self.root)
        picker.grab_set()
        self._center_toplevel(picker, 320, 360)

        cal = Calendar(
            picker,
            selectmode="day",
            year=base_dt.year,
            month=base_dt.month,
            day=base_dt.day,
            date_pattern="yyyy-mm-dd",
        )
        cal.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        time_row = ttk.Frame(picker, padding=(8, 0, 8, 8))
        time_row.pack(fill=tk.X)
        ttk.Label(time_row, text="Время:").pack(side=tk.LEFT)

        hour_var = tk.StringVar(value=f"{base_dt.hour:02d}")
        min_var = tk.StringVar(value=f"{base_dt.minute:02d}")
        sec_var = tk.StringVar(value=f"{base_dt.second:02d}")

        ttk.Entry(time_row, textvariable=hour_var, width=3).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Label(time_row, text=":").pack(side=tk.LEFT)
        ttk.Entry(time_row, textvariable=min_var, width=3).pack(side=tk.LEFT, padx=2)
        ttk.Label(time_row, text=":").pack(side=tk.LEFT)
        ttk.Entry(time_row, textvariable=sec_var, width=3).pack(side=tk.LEFT, padx=2)

        btn_row = ttk.Frame(picker, padding=(8, 0, 8, 8))
        btn_row.pack(fill=tk.X)

        def apply_datetime() -> None:
            try:
                h = max(0, min(23, int(hour_var.get())))
                m = max(0, min(59, int(min_var.get())))
                s = max(0, min(59, int(sec_var.get())))
            except ValueError:
                messagebox.showerror("Ошибка", "Время должно быть числом.")
                return

            selected_date = cal.get_date()  # yyyy-mm-dd
            try:
                chosen = datetime.strptime(selected_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Ошибка", "Некорректная дата.")
                return

            final_dt = chosen.replace(hour=h, minute=m, second=s)
            target_var.set(final_dt.strftime("%a %b %d %H:%M:%S %Y"))
            picker.destroy()

        ttk.Button(btn_row, text="✅ Применить", command=apply_datetime).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_row, text="✕ Отмена", command=picker.destroy).pack(side=tk.RIGHT, padx=2)

    def open_campaign_folder(self, campaign_path_override: str | None = None) -> None:
        if self.selected_game_key is None:
            messagebox.showinfo("Инфо", "Сначала выберите кампанию.")
            return

        if campaign_path_override is not None:
            campaign_path_raw = str(campaign_path_override).strip()
        else:
            game = self.current_player["games"][self.selected_game_key]
            campaign_path_raw = str(game.get("campaign", "")).strip()
        if campaign_path_raw == "":
            messagebox.showerror("Ошибка", "В кампании отсутствует путь к .cmp файлу.")
            return

        cmp_path = Path(campaign_path_raw.replace("/", os.sep))
        is_relative_dcs = campaign_path_raw.startswith("./") or not cmp_path.is_absolute()
        if campaign_path_raw.startswith("./"):
            cmp_path = Path(campaign_path_raw[2:].replace("/", os.sep))

        if is_relative_dcs:
            dcs_root = self.dcs_install_path_var.get().strip()
            if not dcs_root:
                messagebox.showerror(
                    "Нужны настройки",
                    "Для относительного пути кампании сначала укажите путь установки DCS World в кнопке 'Настройки'.",
                )
                return
            cmp_path = Path(dcs_root) / cmp_path

        folder = cmp_path.parent
        if not folder.exists():
            messagebox.showerror(
                "Папка не найдена",
                f"Не удалось найти папку кампании:\n{folder}\n\nПуть взят из поля campaign в logbook.lua.",
            )
            return

        try:
            os.startfile(str(folder))
            self.status_var.set(f"Открыта папка кампании: {folder}")
        except Exception as exc:
            messagebox.showerror("Ошибка открытия", str(exc))

    def add_game_from_template(self) -> None:
        games = self.current_player["games"]
        if self.selected_game_key is not None and self.selected_game_key in games:
            new_game = copy.deepcopy(games[self.selected_game_key])
        elif games:
            src_id = sorted(games.keys())[-1]
            new_game = copy.deepcopy(games[src_id])
        else:
            new_game = {"created": "", "player": "", "campaign": "", "status": "Aктивная", "history": {1: {"stage": 1, "mission": ""}}}

        new_id = self._next_int_key(games)
        created_default = datetime.now().strftime("%a %b %d %H:%M:%S %Y")
        created = simpledialog.askstring("Новая кампания", "created:", initialvalue=new_game.get("created") or created_default)
        if created is None:
            return
        campaign = simpledialog.askstring("Новая кампания", "campaign:", initialvalue=new_game.get("campaign", ""))
        if campaign is None:
            return
        status = simpledialog.askstring("Новая кампания", "status:", initialvalue=new_game.get("status", "Aктивная"))
        if status is None:
            return
        player = simpledialog.askstring("Новая кампания", "player:", initialvalue=new_game.get("player", self.current_player.get("name", "")))
        if player is None:
            return

        new_game["created"] = created
        new_game["campaign"] = campaign
        new_game["status"] = status
        new_game["player"] = player
        games[new_id] = new_game
        self._refresh_games_tree()
        self.games_tree.selection_set(str(new_id))
        self.on_game_select()
        self.status_var.set(f"Добавлена кампания {new_id} по образцу.")

    def delete_game(self) -> None:
        if self.selected_game_key is None:
            return
        if not messagebox.askyesno("Подтвердите", f"Удалить кампанию {self.selected_game_key}?"):
            return
        del self.current_player["games"][self.selected_game_key]
        self.selected_game_key = None
        self._refresh_games_tree()
        self._refresh_history_tree()
        self.status_var.set("Кампания удалена (в памяти).")

    def edit_history_item(self) -> None:
        if self.selected_game_key is None or self.selected_history_key is None:
            messagebox.showinfo("Инфо", "Выберите запись истории.")
            return
        history = self.current_player["games"][self.selected_game_key].setdefault("history", {})
        item = history[self.selected_history_key]
        self._load_history_form(item, target_key=self.selected_history_key)
        self.status_var.set(f"Запись {self.selected_history_key} загружена в редактор справа.")

    def add_history_from_template(self) -> None:
        if self.selected_game_key is None:
            messagebox.showinfo("Инфо", "Сначала выберите кампанию.")
            return
        game = self.current_player["games"][self.selected_game_key]
        history = game.setdefault("history", {})
        if self.selected_history_key is not None and self.selected_history_key in history:
            new_item = copy.deepcopy(history[self.selected_history_key])
        elif history:
            new_item = copy.deepcopy(history[sorted(history.keys())[-1]])
        else:
            new_item = {"stage": 1, "mission": ""}
        self._load_history_form(new_item, target_key=None)
        self.status_var.set("Шаблон записи загружен в редактор справа.")

    def add_empty_history_entry(self) -> None:
        if self.selected_game_key is None:
            messagebox.showinfo("Инфо", "Сначала выберите кампанию.")
            return
        game = self.current_player["games"][self.selected_game_key]
        history = game.setdefault("history", {})
        stage_values = [v.get("stage", 0) for v in history.values() if isinstance(v, dict)]
        max_stage = max([s for s in stage_values if isinstance(s, int)], default=0)
        next_stage = max_stage + 1 if max_stage > 0 else 1
        new_key = self._next_int_key(history)
        new_item = {
            "stage": next_stage,
            "mission": "",
        }
        history[new_key] = new_item
        self.selected_history_key = new_key
        self._refresh_history_tree()
        if str(new_key) in self.history_tree.get_children():
            self.history_tree.selection_set(str(new_key))
            self.history_tree.focus(str(new_key))
        self._load_history_form(new_item, target_key=new_key)
        self.status_var.set(f"Добавлена новая пустая запись {new_key}. Заполните поля справа.")

    def duplicate_history_entry(self) -> None:
        if self.selected_game_key is None:
            messagebox.showinfo("Инфо", "Сначала выберите кампанию.")
            return
        if self.selected_history_key is None:
            messagebox.showinfo("Инфо", "Выберите запись истории для дублирования.")
            return
        game = self.current_player["games"][self.selected_game_key]
        history = game.setdefault("history", {})
        old_key = self.selected_history_key
        if old_key not in history:
            return
        new_key = self._next_int_key(history)
        history[new_key] = copy.deepcopy(history[old_key])
        self.selected_history_key = new_key
        self._refresh_history_tree()
        if str(new_key) in self.history_tree.get_children():
            self.history_tree.selection_set(str(new_key))
            self.history_tree.focus(str(new_key))
        self._load_history_form(history[new_key], target_key=new_key)
        self.status_var.set(f"Запись {old_key} продублирована: новая запись {new_key} (в конце списка).")

    def add_next_mission(self) -> None:
        if self.selected_game_key is None:
            messagebox.showinfo("Инфо", "Сначала выберите кампанию.")
            return
        game = self.current_player["games"][self.selected_game_key]
        history = game.setdefault("history", {})
        if not history:
            template = {"stage": 1, "mission": ""}
            next_stage = 1
        else:
            last_key = sorted(history.keys())[-1]
            template = copy.deepcopy(history[last_key])
            stage_values = [v.get("stage", 0) for v in history.values() if isinstance(v, dict)]
            max_stage = max([s for s in stage_values if isinstance(s, int)], default=0)
            next_stage = max_stage + 1 if max_stage > 0 else len(history) + 1

        new_id = self._next_int_key(history)
        template["stage"] = next_stage
        if "mission" in template and isinstance(template["mission"], str):
            template["mission"] = self._increment_mission_name(template["mission"], next_stage)
        template.pop("datetime", None)
        template.pop("result", None)
        template.pop("aaKills", None)
        template.pop("agKills", None)
        template.pop("deathsCount", None)
        template.pop("skipped", None)
        self._load_history_form(template, target_key=None)
        self.status_var.set("Шаблон следующей миссии загружен. Изменения применяются при вводе.")

    def _increment_mission_name(self, mission_name: str, fallback_stage: int) -> str:
        digits = ""
        end = len(mission_name) - 1
        while end >= 0 and mission_name[end].isdigit():
            digits = mission_name[end] + digits
            end -= 1
        if digits:
            width = len(digits)
            incremented = int(digits) + 1
            return f"{mission_name[: end + 1]}{incremented:0{width}d}"
        return f"{mission_name} ({fallback_stage})"

    def _load_history_form(self, item: dict, target_key: int | None) -> None:
        self.suppress_history_auto_apply = True
        self.history_form_target_key = target_key
        for key in ("stage", "mission", "datetime", "result", "aaKills", "agKills", "deathsCount"):
            self.history_field_vars[key].set("" if key not in item else str(item.get(key, "")))
        self.history_skipped_var.set(bool(item.get("skipped", False)))
        self.suppress_history_auto_apply = False

    def _apply_history_form_live(self, *_args) -> None:
        if self.suppress_history_auto_apply:
            return
        if self.selected_game_key is None:
            return
        game = self.current_player["games"][self.selected_game_key]
        history = game.setdefault("history", {})

        mission = self.history_field_vars["mission"].get().strip()
        if mission == "":
            self.status_var.set("Для автосохранения миссии заполните поле Mission.")
            return

        def parse_int(field_name: str) -> int | None:
            raw = self.history_field_vars[field_name].get().strip()
            if raw == "":
                return None
            try:
                return int(raw)
            except ValueError:
                return 0

        stage = parse_int("stage")
        if stage is None:
            stage = 1

        new_item = {"stage": stage, "mission": mission}
        for fld in ("result", "aaKills", "agKills", "deathsCount"):
            val = parse_int(fld)
            if val is not None:
                new_item[fld] = val
        dt_val = self.history_field_vars["datetime"].get().strip()
        if dt_val:
            new_item["datetime"] = dt_val
        if self.history_skipped_var.get():
            new_item["skipped"] = True

        if self.history_form_target_key is None:
            key = self._next_int_key(history)
            history[key] = new_item
            self.selected_history_key = key
            self.status_var.set(f"Добавлена запись истории {key}. Изменения применяются сразу.")
        else:
            key = self.history_form_target_key
            history[key] = new_item
            self.selected_history_key = key
            self.status_var.set(f"Запись истории {key} обновлена.")

        self._refresh_history_tree()
        if str(key) in self.history_tree.get_children():
            self.history_tree.selection_set(str(key))
            self.history_tree.focus(str(key))
        self._load_history_form(new_item, target_key=key)

    def _parse_datetime_value(self, value: str) -> datetime | None:
        if not value:
            return None
        patterns = (
            "%a %b %d %H:%M:%S %Y",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
        )
        for pattern in patterns:
            try:
                return datetime.strptime(value, pattern)
            except ValueError:
                continue
        return None

    def open_datetime_picker(self) -> None:
        if Calendar is None:
            messagebox.showerror(
                "Календарь недоступен",
                "Не найден модуль tkcalendar.\nУстановите:\npython -m pip install tkcalendar",
            )
            return

        current_value = self.history_field_vars["datetime"].get().strip()
        base_dt = self._parse_datetime_value(current_value) or datetime.now()

        picker = tk.Toplevel(self.root)
        picker.title("Выбор даты и времени")
        picker.geometry("320x360")
        picker.transient(self.root)
        picker.grab_set()
        self._center_toplevel(picker, 320, 360)

        cal = Calendar(
            picker,
            selectmode="day",
            year=base_dt.year,
            month=base_dt.month,
            day=base_dt.day,
            date_pattern="yyyy-mm-dd",
        )
        cal.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        time_row = ttk.Frame(picker, padding=(8, 0, 8, 8))
        time_row.pack(fill=tk.X)
        ttk.Label(time_row, text="Время:").pack(side=tk.LEFT)

        hour_var = tk.StringVar(value=f"{base_dt.hour:02d}")
        min_var = tk.StringVar(value=f"{base_dt.minute:02d}")
        sec_var = tk.StringVar(value=f"{base_dt.second:02d}")

        ttk.Entry(time_row, textvariable=hour_var, width=3).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Label(time_row, text=":").pack(side=tk.LEFT)
        ttk.Entry(time_row, textvariable=min_var, width=3).pack(side=tk.LEFT, padx=2)
        ttk.Label(time_row, text=":").pack(side=tk.LEFT)
        ttk.Entry(time_row, textvariable=sec_var, width=3).pack(side=tk.LEFT, padx=2)

        btn_row = ttk.Frame(picker, padding=(8, 0, 8, 8))
        btn_row.pack(fill=tk.X)

        def apply_datetime() -> None:
            try:
                h = max(0, min(23, int(hour_var.get())))
                m = max(0, min(59, int(min_var.get())))
                s = max(0, min(59, int(sec_var.get())))
            except ValueError:
                messagebox.showerror("Ошибка", "Время должно быть числом.")
                return

            selected_date = cal.get_date()  # yyyy-mm-dd
            try:
                chosen = datetime.strptime(selected_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Ошибка", "Некорректная дата.")
                return

            final_dt = chosen.replace(hour=h, minute=m, second=s)
            self.history_field_vars["datetime"].set(final_dt.strftime("%a %b %d %H:%M:%S %Y"))
            picker.destroy()

        ttk.Button(btn_row, text="✅ Применить", command=apply_datetime).pack(side=tk.RIGHT, padx=2)
        ttk.Button(btn_row, text="✕ Отмена", command=picker.destroy).pack(side=tk.RIGHT, padx=2)

    def _autosize_tree_columns(
        self,
        tree: ttk.Treeview,
        min_widths: dict[str, int],
        max_widths: dict[str, int] | None = None,
    ) -> None:
        for col in tree["columns"]:
            self._autosize_single_column(tree, col, min_widths, max_widths)

    def _autosize_single_column(
        self,
        tree: ttk.Treeview,
        col: str,
        min_widths: dict[str, int],
        max_widths: dict[str, int] | None = None,
    ) -> None:
        font = tkfont.nametofont("TkDefaultFont")
        header_width = font.measure(str(tree.heading(col, "text"))) + 24
        computed_width = max(header_width, min_widths.get(col, 60))
        for item_id in tree.get_children():
            cell_value = tree.set(item_id, col)
            cell_width = font.measure(str(cell_value)) + 24
            if cell_width > computed_width:
                computed_width = cell_width
        if max_widths and col in max_widths:
            computed_width = min(computed_width, max_widths[col])
        tree.column(col, width=computed_width)

    def _on_tree_separator_double_click(
        self,
        tree: ttk.Treeview,
        event,
        min_widths: dict[str, int],
        max_widths: dict[str, int] | None = None,
    ):
        region = tree.identify_region(event.x, event.y)
        if region not in ("separator", "heading"):
            return
        col_token = tree.identify_column(event.x)
        if not col_token.startswith("#"):
            return
        try:
            col_idx = int(col_token[1:]) - 1
        except ValueError:
            return
        columns = tree["columns"]
        if col_idx < 0 or col_idx >= len(columns):
            return
        self._autosize_single_column(tree, columns[col_idx], min_widths, max_widths)
        return "break"

    def _set_initial_editor_width(self) -> None:
        if not hasattr(self, "content_pane"):
            return
        try:
            total_width = self.content_pane.winfo_width()
            if total_width > 950:
                first_col_width = 674
                second_col_width = 720
                sash0 = first_col_width
                sash1 = sash0 + second_col_width
                self.content_pane.sashpos(1, sash1)
                self.content_pane.sashpos(0, sash0)
        except Exception:
            return

    def _square_rgba_icon(self, im: Image.Image) -> Image.Image:
        """Квадратная иконка на прозрачном фоне — Windows ожидает квадратные кадры в ICO."""
        if im.mode != "RGBA":
            im = im.convert("RGBA")
        w, h = im.size
        if w == h:
            return im
        side = max(w, h)
        canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        canvas.paste(im, ((side - w) // 2, (side - h) // 2), im)
        return canvas

    def _build_hires_windows_ico(self, source: Path) -> Path | None:
        """
        Собирает ICO с набором размеров (16–512 px) для панели задач Windows.
        Крупный PNG (например iconMY.png 581×581) используется как есть; апскейл только если
        сторона меньше 256 px.
        """
        try:
            resample = Image.Resampling.LANCZOS
        except AttributeError:
            resample = Image.LANCZOS  # type: ignore[attr-defined]
        try:
            im = Image.open(source)
            if getattr(im, "n_frames", 1) > 1:
                best = None
                best_area = 0
                for i in range(im.n_frames):
                    im.seek(i)
                    frame = im.copy()
                    a = frame.size[0] * frame.size[1]
                    if a > best_area:
                        best_area = a
                        best = frame
                im = best if best is not None else im.copy()
            else:
                im = im.copy()
            im = self._square_rgba_icon(im.convert("RGBA"))
            w, h = im.size
            max_dim = max(w, h)
            if max_dim < 256:
                scale = 256 / max_dim
                im = im.resize(
                    (max(1, int(round(w * scale))), max(1, int(round(h * scale)))),
                    resample,
                )
            fd, tmp = tempfile.mkstemp(suffix=".ico")
            os.close(fd)
            out = Path(tmp)
            sizes = [
                (16, 16),
                (20, 20),
                (24, 24),
                (32, 32),
                (40, 40),
                (48, 48),
                (64, 64),
                (128, 128),
                (256, 256),
                (512, 512),
            ]
            im.save(out, format="ICO", sizes=sizes)
            return out
        except Exception:
            return None

    def _apply_windows_wm_icons(self, ico_path: Path) -> None:
        """
        Tk/wm iconbitmap часто оставляет для окна только мелкий HICON — панель задач тянет его
        и размывает. Явно задаём ICON_BIG (512/256) и ICON_SMALL (32/16) через WM_SETICON.
        """
        if sys.platform != "win32":
            return
        path = os.path.abspath(str(ico_path))
        if not os.path.isfile(path):
            return
        try:
            import ctypes

            user32 = ctypes.windll.user32
            WM_SETICON = 0x0080
            ICON_SMALL = 0
            ICON_BIG = 1
            IMAGE_ICON = 1
            LR_LOADFROMFILE = 0x00000010
            LR_DEFAULTSIZE = 0x00000040

            self.root.update_idletasks()
            tk_hwnd = int(self.root.winfo_id())
            if not tk_hwnd:
                return

            def load_sz(cx: int, cy: int):
                h = user32.LoadImageW(0, path, IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
                return int(h) if h else 0

            h_big = load_sz(512, 512) or load_sz(256, 256) or load_sz(128, 128)
            if not h_big:
                h_big = user32.LoadImageW(0, path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
                h_big = int(h_big) if h_big else 0

            h_small = load_sz(32, 32) or load_sz(24, 24) or load_sz(16, 16)
            if not h_small:
                h_small = user32.LoadImageW(0, path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
                h_small = int(h_small) if h_small else 0

            if not h_big and not h_small:
                return

            for hwnd in (tk_hwnd, int(user32.GetParent(tk_hwnd) or 0)):
                if not hwnd:
                    continue
                if h_big:
                    user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
                if h_small:
                    user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_small)
        except Exception:
            return

    def _set_app_icon(self) -> None:
        """Иконка окна: в Windows для заголовка и панели задач надёжнее .ico + AppUserModelID (см. main)."""
        base = _application_resource_dir()
        iconMY_png = base / "iconMY.png"
        iconMY_ico = base / "iconMY.ico"
        app_png = base / "app_icon.png"
        app_ico = base / "app_icon.ico"
        try:
            if sys.platform == "win32":
                # PNG раньше ICO: из PNG собираем многоразмерный ICO; при отдельном ICO — пересобираем его.
                for src in (iconMY_png, app_png, iconMY_ico, app_ico):
                    if not src.exists():
                        continue
                    built = self._build_hires_windows_ico(src)
                    if built is not None:
                        self._app_icon_temp_path = built
                        self.root.iconbitmap(default=str(built))
                        p = built

                        def apply_icons() -> None:
                            self._apply_windows_wm_icons(p)

                        self.root.after_idle(apply_icons)
                        self.root.after(150, apply_icons)
                        return
            else:
                for png in (iconMY_png, app_png):
                    if png.exists():
                        self._window_icon = tk.PhotoImage(file=str(png))
                        self.root.iconphoto(True, self._window_icon)
                        return
                for ico in (iconMY_ico, app_ico):
                    if ico.exists():
                        self.root.iconbitmap(str(ico))
                        return
        except Exception:
            return

    def delete_history_item(self) -> None:
        if self.selected_game_key is None or self.selected_history_key is None:
            return
        if not messagebox.askyesno("Подтвердите", f"Удалить запись {self.selected_history_key}?"):
            return
        history = self.current_player["games"][self.selected_game_key].setdefault("history", {})
        del history[self.selected_history_key]
        self.selected_history_key = None
        self._refresh_history_tree()
        self.status_var.set("Запись истории удалена (в памяти).")

    def save_file(self) -> None:
        if not self.logbook_data:
            messagebox.showinfo("Инфо", "Сначала загрузите файл.")
            return
        path = Path(self.file_path_var.get().strip())
        if not path.exists():
            messagebox.showerror("Ошибка", f"Файл не найден:\n{path}")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = path.with_name(f"{path.stem}.backup_{timestamp}{path.suffix}")

        try:
            os.makedirs(path.parent, exist_ok=True)
            shutil.copy2(path, backup_path)
            dumped = dump_lua_assignment(self.assignment_name, self.logbook_data)
            path.write_text(dumped, encoding="utf-8")
        except Exception as exc:
            messagebox.showerror("Ошибка сохранения", str(exc))
            return

        self.status_var.set(f"Сохранено. Бэкап: {backup_path.name}")
        messagebox.showinfo("Готово", f"Файл сохранён.\nБэкап создан:\n{backup_path}")


def main() -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            # Чёткие иконки на высоком DPI (в т.ч. панель задач).
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(3)
            except Exception:
                try:
                    ctypes.windll.shcore.SetProcessDpiAwareness(2)
                except Exception:
                    try:
                        ctypes.windll.user32.SetProcessDPIAware()
                    except Exception:
                        pass
            # Иначе Windows сопоставляет процесс с python.exe — чужая иконка на панели.
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "Logbooker.DCSLogbookEditor.Application.1"
            )
        except Exception:
            pass
    root = tk.Tk()
    app = LogbookEditorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
