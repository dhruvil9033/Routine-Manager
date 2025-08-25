import os
import win32com.client

def get_all_start_menu_shortcuts():
    start_menu_dirs = [
        os.path.join(os.environ["PROGRAMDATA"], r"Microsoft\Windows\Start Menu\Programs"),
        os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs"),
    ]

    shortcuts = {}

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
                            shortcuts[app_name] = target_path
                    except Exception:
                        continue

    return shortcuts

# Example usage
if __name__ == "__main__":
    shortcuts = get_all_start_menu_shortcuts()
    for name, path in list(shortcuts.items())[:10]:  # print first 10
        print(f"{name}: {path}")
