# The RoutineManagerGUI class provides a GUI for managing app launch routines.
# It can be invoked by the voice assistant on the command "open routines".

import os
import json
import subprocess
import ctypes
import win32com.client
from tkinter import messagebox
import customtkinter
from tkinter import ttk

# Set appearance (dark theme with default blue color)
customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("blue")

class RoutineManagerGUI(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("App Launch Routines Manager")
        self.geometry("800x600")

        self.available_apps = self.get_start_menu_shortcuts()

        # Path to routines JSON file
        self.routines_file = os.path.join(os.getcwd(), "routines.json")
        self.routines = {}
        self.load_routines()

        # Set up main frames
        top_frame = customtkinter.CTkFrame(self)
        top_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)
        
        left_frame = customtkinter.CTkFrame(top_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        
        right_frame = customtkinter.CTkFrame(top_frame)
        right_frame.pack(side="right", fill="y", padx=10, pady=10)
        
        # Treeview for routines and their apps
        self.tree = ttk.Treeview(left_frame, columns=("Admin","Order"), show="tree headings")
        self.tree.heading("#0", text="Routine / App")
        self.tree.heading("Admin", text="Admin?")
        self.tree.heading("Order", text="Order")
        self.tree.column("Admin", width=60, anchor="center")
        self.tree.column("Order", width=50, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(left_frame, orient="horizontal", command=self.tree.xview)
        hsb.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=hsb.set)

        # Buttons in the right frame
        self.add_button = customtkinter.CTkButton(right_frame, text="Add Routine", command=self.open_new_routine_window)
        self.add_button.pack(pady=5)
        self.run_button = customtkinter.CTkButton(right_frame, text="Run Routine", command=self.run_selected_routine, state="disabled")
        self.run_button.pack(pady=5)
        self.delete_button = customtkinter.CTkButton(right_frame, text="Delete Routine", command=self.delete_selected_routine, state="disabled")
        self.delete_button.pack(pady=5)
        self.close_button = customtkinter.CTkButton(right_frame, text="Close", command=self.destroy)
        self.close_button.pack(pady=5)
        
        # Status log textbox at bottom
        self.log_text = customtkinter.CTkTextbox(self, height=150)
        self.log_text.configure(state="disabled")
        self.log_text.pack(side="bottom", fill="x", padx=10, pady=10)

        # Bind selection event
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # Populate the tree with routines
        self.load_treeview_data()

    def load_routines(self):
        # Load routines from JSON
        try:
            if os.path.exists(self.routines_file):
                with open(self.routines_file, 'r') as f:
                    self.routines = json.load(f)
            else:
                self.routines = {}
        except (json.JSONDecodeError, FileNotFoundError):
            self.routines = {}

    def save_routines(self):
        # Save routines to JSON
        try:
            with open(self.routines_file, 'w') as f:
                json.dump(self.routines, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save routines: {e}")

    def load_treeview_data(self):
        # Refresh the tree display of routines
        for item in self.tree.get_children():
            self.tree.delete(item)
        for routine, apps in self.routines.items():
            parent = self.tree.insert("", "end", text=routine, values=("", ""))
            for idx, app in enumerate(apps):
                app_name = app.get("app", "")
                admin_flag = "Yes" if app.get("admin", False) else "No"
                order = idx + 1
                self.tree.insert(parent, "end", text=app_name, values=(admin_flag, order))

    def on_tree_select(self, event):
        # Enable or disable buttons based on selection
        selected = self.tree.selection()
        if not selected:
            self.run_button.configure(state="disabled")
            self.delete_button.configure(state="disabled")
            return
        sel = selected[0]
        parent = self.tree.parent(sel)
        routine_id = parent if parent else sel
        if routine_id:
            self.run_button.configure(state="normal")
            self.delete_button.configure(state="normal")
        else:
            self.run_button.configure(state="disabled")
            self.delete_button.configure(state="disabled")
    
    def open_new_routine_window(self):
        """Opens a window to create a new routine with dropdown-based app selection."""
        if hasattr(self, 'new_window') and self.new_window.winfo_exists():
            self.new_window.focus()
            return

        self.new_window = customtkinter.CTkToplevel(self)
        self.new_window.lift()
        self.new_window.attributes("-topmost", True)
        self.new_window.after(200, lambda: self.new_window.attributes("-topmost", False))
        self.new_window.focus_force()
        self.new_window.title("Create New Routine")
        self.new_window.geometry("520x460")

        # Routine Name
        name_frame = customtkinter.CTkFrame(self.new_window)
        name_frame.pack(fill="x", padx=10, pady=(15, 5))

        name_label = customtkinter.CTkLabel(name_frame, text="Routine Name:")
        name_label.pack(side="left", padx=(0, 10))

        self.name_entry = customtkinter.CTkEntry(name_frame)
        self.name_entry.pack(side="left", fill="x", expand=True)

        # App Selection Frame
        self.apps_frame = customtkinter.CTkScrollableFrame(self.new_window, width=480, height=230)
        self.apps_frame.pack(padx=10, pady=(10, 5), fill="both", expand=True)

        self.app_rows = []
        self.add_app_row()

        # App Add Button
        add_app_button = customtkinter.CTkButton(self.new_window, text="Add Another App", command=self.add_app_row)
        add_app_button.pack(pady=8)

        # Action Buttons
        button_frame = customtkinter.CTkFrame(self.new_window)
        button_frame.pack(fill="x", padx=15, pady=(10, 15))

        save_button = customtkinter.CTkButton(button_frame, text="Save Routine", command=self.save_new_routine)
        save_button.pack(side="left", padx=5, expand=True)

        cancel_button = customtkinter.CTkButton(button_frame, text="Cancel", command=self.new_window.destroy)
        cancel_button.pack(side="right", padx=5, expand=True)

    def add_app_row(self):
        """Adds a new app row with dropdown and controls."""
        frame = customtkinter.CTkFrame(self.apps_frame)
        frame.pack(fill="x", pady=4, padx=5)

        # App dropdown
        app_combo = customtkinter.CTkComboBox(
            frame,
            values=list(self.available_apps.keys()),
            width=260,
        )
        app_combo.set("Select App")
        app_combo.pack(side="left", padx=(5, 8), fill="x", expand=True)

        # Admin checkbox
        admin_var = customtkinter.CTkCheckBox(frame, text="Admin")
        admin_var.pack(side="left", padx=4)

        # Row reorder and remove buttons
        up_btn = customtkinter.CTkButton(frame, text="↑", width=30, command=lambda f=frame: self.move_app_row_up(f))
        up_btn.pack(side="left", padx=2)
        down_btn = customtkinter.CTkButton(frame, text="↓", width=30, command=lambda f=frame: self.move_app_row_down(f))
        down_btn.pack(side="left", padx=2)
        rm_btn = customtkinter.CTkButton(frame, text="✕", width=40, command=lambda f=frame: self.remove_app_row(f))
        rm_btn.pack(side="left", padx=(2, 4))

        self.app_rows.append((frame, app_combo, admin_var))

    
    def remove_app_row(self, frame):
        # Remove an app row from the new routine form
        for i, (frm, entry, chk) in enumerate(self.app_rows):
            if frm == frame:
                frm.destroy()
                self.app_rows.pop(i)
                break

    def move_app_row_up(self, frame):
        # Move an app row up
        for i, (frm, entry, chk) in enumerate(self.app_rows):
            if frm == frame and i > 0:
                self.app_rows[i], self.app_rows[i-1] = self.app_rows[i-1], self.app_rows[i]
                break
        self.reorder_app_rows()

    def move_app_row_down(self, frame):
        # Move an app row down
        for i, (frm, entry, chk) in enumerate(self.app_rows):
            if frm == frame and i < len(self.app_rows)-1:
                self.app_rows[i], self.app_rows[i+1] = self.app_rows[i+1], self.app_rows[i]
                break
        self.reorder_app_rows()

    def reorder_app_rows(self):
        # Re-pack app rows after reordering
        for frm, entry, chk in self.app_rows:
            frm.pack_forget()
        for frm, entry, chk in self.app_rows:
            frm.pack(fill="x", pady=2)

    def save_new_routine(self):
        name = self.name_entry.get().strip()

        if not name:
            messagebox.showwarning("Input Error", "Routine name cannot be empty.")
            return

        # Case-insensitive duplicate check
        existing_names = [n.lower() for n in self.controller.routines.keys()]
        if name.lower() in existing_names:
            messagebox.showerror("Duplicate Routine", f"A routine named '{name}' already exists.")
            return

        apps = []
        for frame, app_combo, admin_check in self.app_rows:
            app_name = app_combo.get().strip()
            if app_name and app_name.lower() != "select app":
                apps.append({
                    "name": app_name,
                    "admin": bool(admin_check.get())
                })

        if not apps:
            messagebox.showwarning("Input Error", "Please add at least one valid app.")
            return

        self.controller.routines[name] = apps
        self.controller.save_routines()
        messagebox.showinfo("Success", f"Routine '{name}' saved successfully.")
        self.new_window.destroy()

    def delete_selected_routine(self):
        selected = self.tree.selection()
        if not selected:
            return
        sel = selected[0]
        parent = self.tree.parent(sel)
        if parent:
            sel = parent
        routine_name = self.tree.item(sel, "text")
        if messagebox.askyesno("Delete Routine", f"Are you sure you want to delete '{routine_name}'?"):
            if routine_name in self.routines:
                del self.routines[routine_name]
                self.save_routines()
                self.load_treeview_data()
                self.log(f"Routine '{routine_name}' deleted.")

    def run_selected_routine(self):
        selected = self.tree.selection()
        if not selected:
            return
        sel = selected[0]
        parent = self.tree.parent(sel)
        if parent:
            sel = parent
        routine_name = self.tree.item(sel, "text")
        apps = self.routines.get(routine_name, [])
        self.log(f"Running routine '{routine_name}'...")
        for app_info in apps:
            app = app_info.get("app")
            admin = app_info.get("admin", False)
            try:
                if admin:
                    # Launch with admin privileges (Windows)
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", app, None, None, 1)
                    self.log(f" Launched (admin): {app}")
                else:
                    # Launch normally
                    subprocess.Popen(app, shell=True)
                    self.log(f" Launched: {app}")
            except Exception as e:
                self.log(f" Error launching {app}: {e}")

    def get_start_menu_shortcuts(self):
        # Fetch all .lnk shortcuts from system and user start menu
        start_menu_dirs = [
            os.path.join(os.environ["PROGRAMDATA"], r"Microsoft\Windows\Start Menu\Programs"),
            os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs"),
        ]

        app_dict = {}

        for start_dir in start_menu_dirs:
            for root, _, files in os.walk(start_dir):
                for file in files:
                    if file.endswith(".lnk"):
                        full_path = os.path.join(root, file)
                        try:
                            shell = win32com.client.Dispatch("WScript.Shell")
                            shortcut = shell.CreateShortCut(full_path)
                            target_path = shortcut.Targetpath
                            if target_path and os.path.exists(target_path):
                                app_name = os.path.splitext(file)[0]
                                app_dict[app_name] = target_path
                        except Exception:
                            continue
        return dict(sorted(app_dict.items()))  # sort alphabetically


    def log(self, message):
        # Append message to log textbox
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.configure(state="disabled")
        try:
            self.log_text.yview_moveto(1.0)
        except:
            pass

# Example integration with voice assistant:
# For instance, if the assistant hears "open routines", you could call:
#     gui = RoutineManagerGUI()
#     gui.mainloop()

if __name__ == "__main__":
    app = RoutineManagerGUI()
    app.mainloop()
