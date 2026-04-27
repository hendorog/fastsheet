#!/bin/bash
# Build the Windows .exe and install it for use from the Start Menu.
#
# After running this, Win key → "fastsheet" → Enter launches the latest
# build. Re-running just rebuilds + overwrites the installed copy; the
# Start Menu shortcut is created once and survives subsequent runs.
#
# Run from the repo root:
#   scripts/install-windows.sh
#
# What it does:
#   1. cross-builds the .exe via cargo-xwin
#   2. copies it to %USERPROFILE%\Tools\fastsheet\fastsheet.exe
#      (stable Windows-native path — avoids the WSL UNC slowdown on
#      file I/O AND keeps SmartScreen happy by not running off a
#      network-zone path)
#   3. strips Mark-of-the-Web if any zone identifier is present, so
#      Windows doesn't show the "unrecognized app" warning
#   4. creates a Start Menu shortcut on first run

set -euo pipefail

cd "$(dirname "$0")/.."

# --- 1. build ---
echo "→ building Windows .exe via cargo-xwin..."
npx tauri build --runner cargo-xwin --target x86_64-pc-windows-msvc --no-bundle

# --- 2. resolve Windows-side paths ---
USERPROFILE_WIN="$(cmd.exe /c 'echo %USERPROFILE%' 2>/dev/null | tr -d '\r')"
USERPROFILE="$(wslpath "$USERPROFILE_WIN")"

INSTALL_DIR="$USERPROFILE/Tools/fastsheet"
INSTALL_EXE="$INSTALL_DIR/fastsheet.exe"
START_MENU="$USERPROFILE/AppData/Roaming/Microsoft/Windows/Start Menu/Programs"
SHORTCUT="$START_MENU/fastsheet.lnk"

TARGET_WIN="$USERPROFILE_WIN\\Tools\\fastsheet\\fastsheet.exe"
WORKDIR_WIN="$USERPROFILE_WIN\\Tools\\fastsheet"
SHORTCUT_WIN="$USERPROFILE_WIN\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\fastsheet.lnk"

# Resolve PowerShell — not always on PATH inside a non-login WSL shell,
# but the binary is at a fixed location on every Windows install.
PWSH="/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"

# --- 3. copy + unblock ---
mkdir -p "$INSTALL_DIR"
cp src-tauri/target/x86_64-pc-windows-msvc/release/fastsheet.exe "$INSTALL_EXE"
echo "→ installed: $INSTALL_EXE"

# Strip Mark-of-the-Web (Zone.Identifier ADS) if set. Locally-built
# files don't normally carry it, but Unblock-File is a no-op when it's
# already absent so this is safe to always run.
"$PWSH" -NoProfile -Command "Unblock-File -Path '$TARGET_WIN'" >/dev/null 2>&1 || true

# --- 4. shortcut (idempotent) ---
if [ ! -f "$SHORTCUT" ]; then
  echo "→ creating Start Menu shortcut..."
  "$PWSH" -NoProfile -Command "
    \$ws = New-Object -ComObject WScript.Shell
    \$s = \$ws.CreateShortcut('$SHORTCUT_WIN')
    \$s.TargetPath = '$TARGET_WIN'
    \$s.WorkingDirectory = '$WORKDIR_WIN'
    \$s.Description = 'fastsheet — keyboard-first spreadsheet'
    \$s.Save()
  " >/dev/null
  echo "→ shortcut: $SHORTCUT"
else
  echo "→ shortcut already present at: $SHORTCUT"
fi

echo
echo "Done. Press Win key, type 'fastsheet', hit Enter."
