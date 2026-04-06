from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
import time


def parse_env_file(env_path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not env_path.exists():
        return values

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        values[key] = value
    return values


def parse_bool(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


def resolve_path(project_root: Path, raw_path: str) -> Path:
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = project_root / path
    return path.resolve()


def chunked(items: list[Path], size: int) -> list[list[Path]]:
    return [items[index:index + size] for index in range(0, len(items), size)]


def load_state(state_path: Path) -> dict[str, int]:
    if not state_path.exists():
        return {"next_batch_index": 0}
    try:
        return json.loads(state_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"next_batch_index": 0}


def save_state(state_path: Path, batch_index: int) -> None:
    state_path.write_text(
        json.dumps({"next_batch_index": batch_index}, indent=2),
        encoding="utf-8",
    )


def open_in_chrome(url: str) -> None:
    subprocess.run(["open", "-a", "Google Chrome", url], check=False)


def run_applescript(script: str) -> None:
    subprocess.run(
        ["osascript", "-e", script],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )


def activate_chrome() -> None:
    script = '''
tell application "Google Chrome"
    activate
end tell
'''
    run_applescript(script)


def run_chrome_javascript(javascript: str) -> str:
    escaped_javascript = json.dumps(javascript.strip())
    script = f'''
tell application "Google Chrome"
    if (count of windows) = 0 then error "Google Chrome is not open."
    return execute active tab of front window javascript {escaped_javascript}
end tell
'''
    result = subprocess.run(
        ["osascript", "-e", script],
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def click_button_in_chrome(label: str) -> str:
    javascript = f"""
(() => {{
  const targetLabel = {json.dumps(label)};
  const candidates = Array.from(document.querySelectorAll('button,span,a,div,[role="button"]'));
  const match = candidates.find((node) => node.textContent && node.textContent.trim() === targetLabel);
  if (!match) {{
    return 'missing';
  }}
  const clickable = match.closest('button,a,[role="button"]') || match;
  clickable.scrollIntoView({{block: 'center', inline: 'center'}});
  clickable.click();
  return 'clicked';
}})()
""".strip()
    return run_chrome_javascript(javascript)


def get_browse_target_info(label: str) -> str:
    javascript = f"""
(() => {{
  const targetLabel = {json.dumps(label)};
  const candidates = Array.from(document.querySelectorAll('span,button,a,div,label,[role="button"]'));
  const textMatch = candidates.find((node) => node.textContent && node.textContent.trim() === targetLabel);
  if (!textMatch) {{
    return '';
  }}
  const clickable = textMatch.closest('button,a,label,[role="button"]') || textMatch;
  const associatedInput =
    (clickable.tagName === 'LABEL' && clickable.htmlFor && document.getElementById(clickable.htmlFor)) ||
    clickable.querySelector('input[type="file"]') ||
    textMatch.querySelector && textMatch.querySelector('input[type="file"]') ||
    document.querySelector('input[type="file"]');
  const target = associatedInput || clickable;
  target.scrollIntoView({{block: 'center', inline: 'center'}});
  if (target.tagName === 'INPUT' && target.type === 'file') {{
    target.hidden = false;
    target.disabled = false;
    target.tabIndex = 0;
    target.style.display = 'block';
    target.style.visibility = 'visible';
    target.style.opacity = '1';
    target.style.pointerEvents = 'auto';
    target.style.position = 'fixed';
    target.style.top = '120px';
    target.style.left = '120px';
    target.style.width = '320px';
    target.style.height = '44px';
    target.style.zIndex = '2147483647';
  }}
  if (typeof target.focus === 'function') {{
    target.focus();
  }}
  target.style.outline = '3px solid red';
  return JSON.stringify({{
    tag: target.tagName,
    type: target.type || '',
    text: clickable.textContent ? clickable.textContent.trim() : '',
    focused: document.activeElement === target
  }});
}})()
""".strip()
    return run_chrome_javascript(javascript)


def press_key_in_chrome(key_code: int) -> None:
    script = f'''
tell application "Google Chrome"
    activate
end tell
delay 0.5
tell application "System Events"
    key code {key_code}
end tell
'''
    run_applescript(script)


def press_keys_in_chrome(key_codes: list[int]) -> None:
    activate_chrome()
    time.sleep(0.5)
    for key_code in key_codes:
        press_key_in_chrome(key_code)
        time.sleep(0.3)


def select_batch_files_in_dialog(import_folder: Path, batch: list[Path]) -> None:
    escaped_path = str(import_folder).replace("\\", "\\\\").replace('"', '\\"')
    escaped_first_name = batch[0].name.replace("\\", "\\\\").replace('"', '\\"')
    extend_selection_steps = max(len(batch) - 1, 0)
    script = f'''
tell application "Google Chrome"
    activate
end tell
delay 0.8
tell application "System Events"
    keystroke "g" using {{command down, shift down}}
    delay 1.2
    tell process "Google Chrome"
        if exists sheet 1 of window 1 then
            tell sheet 1 of window 1
                if exists text field 1 then
                    set value of text field 1 to "{escaped_path}"
                end if
            end tell
        end if
    end tell
    delay 0.3
    key code 36
    delay 1.5
    keystroke "{escaped_first_name}"
    delay 1.0
    repeat {extend_selection_steps} times
        key code 125 using {{shift down}}
        delay 0.15
    end repeat
    delay 0.5
    key code 36
end tell
'''
    run_applescript(script)


def prompt_token(message: str, expected: str) -> None:
    while True:
        response = input(f"{message} Type '{expected}' to continue: ").strip().lower()
        if response == expected.lower():
            return
        print(f"Did not continue. Expected '{expected}'.")


def debug_pause(enabled: bool, message: str) -> None:
    if not enabled:
        return
    prompt_token(f"DEBUG: {message}", "go")


def main() -> None:
    project_root = Path(__file__).resolve().parent
    env = parse_env_file(project_root / ".env")

    import_url = env.get("GARMIN_IMPORT_URL", "https://connect.garmin.com/app/import-data")
    import_folder = resolve_path(project_root, env.get("GARMIN_IMPORT_FOLDER", "garmin_gpx_output"))
    file_pattern = env.get("GARMIN_IMPORT_FILE_PATTERN", "*.gpx")
    batch_size = int(env.get("GARMIN_IMPORT_BATCH_SIZE", "10"))
    state_path = resolve_path(project_root, env.get("GARMIN_IMPORT_STATE_FILE", ".garmin_import_state.json"))
    reset_progress = parse_bool(env.get("GARMIN_IMPORT_RESET_PROGRESS", "false"))
    automate_dialog = parse_bool(env.get("GARMIN_IMPORT_AUTOMATE_DIALOG", "true"))
    import_debug = parse_bool(env.get("IMPORT_DEBUG", "false"))
    browse_click_delay = float(env.get("GARMIN_IMPORT_BROWSE_CLICK_DELAY", "1.5"))
    import_click_delay = float(env.get("GARMIN_IMPORT_IMPORT_CLICK_DELAY", "1.5"))
    batch_wait_seconds = float(env.get("GARMIN_IMPORT_BATCH_WAIT_SECONDS", "8"))

    files = sorted(import_folder.glob(file_pattern))
    if not files:
        raise FileNotFoundError(f"No files matched {file_pattern!r} in {import_folder}")

    batches = chunked(files, batch_size)
    state = {"next_batch_index": 0} if reset_progress else load_state(state_path)
    next_batch_index = state.get("next_batch_index", 0)

    if next_batch_index >= len(batches):
        print("All batches are already marked complete.")
        print("Set GARMIN_IMPORT_RESET_PROGRESS=true in .env to start over.")
        return

    print(f"Opening Garmin import page in Chrome: {import_url}")
    open_in_chrome(import_url)
    prompt_token(
        "Log in if needed and leave the Garmin import page open in front.",
        "ready",
    )

    for batch_index in range(next_batch_index, len(batches)):
        batch = batches[batch_index]

        print(f"\nBatch {batch_index + 1}/{len(batches)}")
        print(f"Source directory: {import_folder}")
        for file_path in batch:
            print(f"  - {file_path.name}")

        if automate_dialog:
            debug_pause(import_debug, "Inspect Garmin before clicking Browse.")
            print("Clicking Garmin's Browse control...")
            browse_target = get_browse_target_info("Browse")
            if not browse_target:
                raise RuntimeError("Could not locate Garmin Browse control in Chrome.")
            if import_debug:
                print(f"Browse target: {browse_target}")
            browse_data = json.loads(browse_target)
            if not browse_data.get("focused"):
                raise RuntimeError(f"Could not focus Garmin Browse control: {browse_data}")
            press_keys_in_chrome([49, 36])
            time.sleep(browse_click_delay)
            debug_pause(import_debug, "The file picker should be open. Confirm before selecting files.")
            print("Selecting batch files directly from the source folder in the open macOS file picker...")
            print("If nothing happens, make sure Accessibility permissions are enabled and the file dialog is frontmost.")
            select_batch_files_in_dialog(import_folder, batch)
            time.sleep(import_click_delay)
            debug_pause(import_debug, "Files should be selected. Confirm before clicking Import Data.")
            print("Clicking Garmin's Import Data button...")
            import_result = click_button_in_chrome("Import Data")
            if import_result != "clicked":
                raise RuntimeError(f"Could not click Garmin Import Data button. Chrome returned: {import_result!r}")
            debug_pause(import_debug, "Import Data was clicked. Confirm before waiting for Garmin to finish the batch.")
            time.sleep(batch_wait_seconds)
        else:
            prompt_token(
                "Dialog automation disabled. Open the Garmin file picker and select the listed files from the source folder manually, then",
                "next" if batch_index < len(batches) - 1 else "done",
            )

        save_state(state_path, batch_index + 1)

    print("All batches processed.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped by user.")
        sys.exit(130)
