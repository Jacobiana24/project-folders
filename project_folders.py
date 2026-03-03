#!/usr/bin/env python3
"""
Project Folders — Quick-access tool for Stockport Council project folder paths.

Reads Obsidian markdown notes to display project buttons grouped by effort level.
Clicking a button copies the Project_Folder path to clipboard and opens it in Explorer.
"""

import json
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
import frontmatter

# ─── Constants ────────────────────────────────────────────────────────────────

CONFIG_FILE = "folders_config.json"
DEFAULT_CONFIG = {
    "vault_path": r"C:\Users\jacob.hand\OneDrive - Stockport Metropolitan Borough Council\Documents\Jacob Hand SMBC PKM\Slip Box",
    "window_width": 900,
    "window_height": 600,
    "window_x": 150,
    "window_y": 150,
}

EFFORT_ORDER = [
    ("1 - Owned", "Owned"),
    ("2 - High", "High"),
    ("3 - Medium", "Medium"),
    ("4 - Low", "Low"),
]
EFFORT_LABELS = {k: v for k, v in EFFORT_ORDER}

BTN_COLOR = "#667eea"
BTN_HOVER = "#5568d3"
BTN_COPIED = "#28a745"

# ─── Utility Functions ────────────────────────────────────────────────────────


def config_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    return base / CONFIG_FILE


def load_config() -> dict:
    cp = config_path()
    if cp.exists():
        try:
            with open(cp, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)
            return cfg
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as e:
        print(f"Warning: could not save config: {e}")


def open_folder(path_str: str):
    """Open a folder in Windows Explorer. Handles missing folders gracefully."""
    p = Path(path_str)
    try:
        if p.exists():
            os.startfile(str(p))
        else:
            # Try opening the parent if the exact path doesn't exist
            parent = p
            while parent != parent.parent and not parent.exists():
                parent = parent.parent
            if parent.exists():
                os.startfile(str(parent))
    except Exception as e:
        print(f"Warning: could not open folder: {e}")


# ─── Project Data ─────────────────────────────────────────────────────────────


class ProjectNote:
    def __init__(self, filepath: Path):
        self.filepath = filepath
        self.name = filepath.stem
        self.post: frontmatter.Post | None = None
        self.load()

    def load(self):
        with open(self.filepath, "r", encoding="utf-8") as f:
            self.post = frontmatter.load(f)

    @property
    def meta(self) -> dict:
        return self.post.metadata

    @property
    def cls(self) -> str:
        return str(self.meta.get("Class", ""))

    @property
    def status(self) -> str:
        return str(self.meta.get("Status", ""))

    @property
    def effort(self) -> str:
        return str(self.meta.get("Effort", ""))

    @property
    def effort_group(self) -> str:
        e = self.effort.strip().strip('"').strip("'")
        return EFFORT_LABELS.get(e, "Unassigned")

    @property
    def project_folder(self) -> str | None:
        val = self.meta.get("Project_Folder")
        if val:
            return str(val).strip().strip('"').strip("'")
        return None

def scan_projects(vault_path: str) -> list[ProjectNote]:
    """Scan vault for active Project notes with a Project_Folder."""
    projects = []
    vp = Path(vault_path)
    if not vp.exists():
        return projects

    for md in sorted(vp.rglob("*.md")):
        try:
            p = ProjectNote(md)
            if (
                p.cls == "Project"
                and p.status == "Active"
                and p.project_folder
            ):
                projects.append(p)
        except Exception as e:
            print(f"Warning: could not load {md}: {e}")

    return projects


# ─── GUI Application ─────────────────────────────────────────────────────────


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.config = load_config()
        self.title("Project Folders")
        self.geometry(
            f"{self.config['window_width']}x{self.config['window_height']}"
            f"+{self.config['window_x']}+{self.config['window_y']}"
        )
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.projects: list[ProjectNote] = []
        self._copied_reset_ids: dict[str, str] = {}  # track after() ids for button resets

        self._build_ui()
        self._initial_load()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        # Top bar
        top = ctk.CTkFrame(self, height=40, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(6, 0))

        self.count_label = ctk.CTkLabel(
            top, text="", font=ctk.CTkFont(size=13), anchor="w"
        )
        self.count_label.pack(side="left")

        ctk.CTkButton(
            top, text="⚙ Settings", width=100, height=28,
            command=self._open_settings,
            fg_color="#555555", hover_color="#666666"
        ).pack(side="right")

        ctk.CTkButton(
            top, text="🔄 Refresh", width=100, height=28,
            command=self._refresh,
            fg_color="#555555", hover_color="#666666"
        ).pack(side="right", padx=(0, 6))

        # Scrollable button area
        self.scroll = ctk.CTkScrollableFrame(self)
        self.scroll.pack(fill="both", expand=True, padx=10, pady=8)

        # Status bar
        self.status_frame = ctk.CTkFrame(self, height=28, corner_radius=0, fg_color="#1a1a2e")
        self.status_frame.pack(fill="x", side="bottom")
        self.status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            self.status_frame, text="Ready",
            font=ctk.CTkFont(size=12), text_color="#aaaaaa", anchor="w"
        )
        self.status_label.pack(fill="x", padx=12, pady=3)

    def _initial_load(self):
        vault = self.config["vault_path"]
        if not Path(vault).exists():
            from tkinter import messagebox
            messagebox.showerror(
                "Vault Not Found",
                f"Projects folder not found:\n{vault}\n\nUse Settings to set the correct path."
            )
        self._refresh()

    def _refresh(self):
        self.projects = scan_projects(self.config["vault_path"])
        self._populate_buttons()

    def _populate_buttons(self):
        for w in self.scroll.winfo_children():
            w.destroy()

        if not self.projects:
            self.count_label.configure(text="No matching projects found")
            ctk.CTkLabel(
                self.scroll,
                text="No active projects with Project_Folder found.\nCheck Settings to verify the vault path and client filter.",
                font=ctk.CTkFont(size=13), text_color="#aaaaaa",
            ).pack(pady=40)
            return

        count = len(self.projects)
        self.count_label.configure(
            text=f"📁  {count} active project{'s' if count != 1 else ''} — click to copy path & open folder"
        )

        # Group by effort
        groups: dict[str, list[ProjectNote]] = {}
        for label in [v for _, v in EFFORT_ORDER] + ["Unassigned"]:
            groups[label] = []

        for p in self.projects:
            g = p.effort_group
            if g not in groups:
                groups["Unassigned"].append(p)
            else:
                groups[g].append(p)

        for group_label, projs in groups.items():
            if not projs:
                continue

            # Group heading
            ctk.CTkLabel(
                self.scroll, text=group_label,
                font=ctk.CTkFont(size=14, weight="bold"), anchor="w"
            ).pack(fill="x", padx=4, pady=(12, 4))

            # Button container for wrapping
            container = ctk.CTkFrame(self.scroll, fg_color="transparent")
            container.pack(fill="x", padx=4, pady=2)

            for p in sorted(projs, key=lambda x: x.name):
                btn = ctk.CTkButton(
                    container,
                    text=p.name,
                    height=40,
                    font=ctk.CTkFont(size=12, weight="bold"),
                    fg_color=BTN_COLOR,
                    hover_color=BTN_HOVER,
                    corner_radius=6,
                    command=lambda proj=p: self._on_click(proj),
                )
                # Auto-size width to text content
                btn.pack(side="left", padx=4, pady=4)

    def _on_click(self, project: ProjectNote):
        folder_path = project.project_folder
        if not folder_path:
            return

        # Copy to clipboard
        self.clipboard_clear()
        self.clipboard_append(folder_path)

        # Open in Explorer
        open_folder(folder_path)

        # Update status bar
        self.status_label.configure(
            text=f"✓ Copied & opened: {project.name}",
            text_color="#28a745"
        )

        # Find the button and flash it green
        self._flash_button(project.name)

        # Reset status bar after 3 seconds
        self.after(3000, lambda: self.status_label.configure(
            text="Ready", text_color="#aaaaaa"
        ))

    def _flash_button(self, project_name: str):
        """Briefly turn the button green to confirm the action."""
        # Search through scroll children for the button
        for container in self.scroll.winfo_children():
            if not isinstance(container, ctk.CTkFrame):
                continue
            for widget in container.winfo_children():
                if isinstance(widget, ctk.CTkButton) and widget.cget("text") == project_name:
                    original_text = widget.cget("text")
                    widget.configure(text="✓ Copied!", fg_color=BTN_COPIED)

                    # Cancel any existing reset timer for this button
                    if project_name in self._copied_reset_ids:
                        self.after_cancel(self._copied_reset_ids[project_name])

                    # Reset after 2 seconds
                    after_id = self.after(2000, lambda w=widget, t=original_text: w.configure(
                        text=t, fg_color=BTN_COLOR
                    ))
                    self._copied_reset_ids[project_name] = after_id
                    return

    # ── Settings ──────────────────────────────────────────────────────────

    def _open_settings(self):
        from tkinter import messagebox

        dialog = ctk.CTkToplevel(self)
        dialog.title("Settings")
        dialog.geometry("500x180")
        dialog.transient(self)
        dialog.grab_set()

        # Vault path
        ctk.CTkLabel(dialog, text="Vault Path:").pack(padx=16, pady=(16, 4), anchor="w")

        path_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        path_frame.pack(fill="x", padx=16, pady=(0, 8))

        path_entry = ctk.CTkEntry(path_frame, width=400)
        path_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        path_entry.insert(0, self.config["vault_path"])

        def browse():
            from tkinter import filedialog
            d = filedialog.askdirectory(initialdir=self.config["vault_path"])
            if d:
                path_entry.delete(0, "end")
                path_entry.insert(0, d)

        ctk.CTkButton(path_frame, text="Browse", width=80, command=browse).pack(side="right")

        def save():
            new_path = path_entry.get().strip()
            if not new_path or not Path(new_path).exists():
                messagebox.showerror("Invalid Path", "The vault path does not exist.", parent=dialog)
                return
            self.config["vault_path"] = new_path
            save_config(self.config)
            self._refresh()
            dialog.destroy()

        ctk.CTkButton(dialog, text="Save", width=100, command=save).pack(pady=8)

    # ── Window Lifecycle ──────────────────────────────────────────────────

    def _on_close(self):
        try:
            geo = self.geometry()
            parts = geo.replace("+", "x").split("x")
            if len(parts) >= 4:
                self.config["window_width"] = int(parts[0])
                self.config["window_height"] = int(parts[1])
                self.config["window_x"] = int(parts[2])
                self.config["window_y"] = int(parts[3])
        except Exception:
            pass
        save_config(self.config)
        self.destroy()


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
