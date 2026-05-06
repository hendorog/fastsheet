# fastsheet — quickstart

Everything you need to start using fastsheet in 5 minutes. This is a
living document; check the bottom for what's not yet implemented.

## Table of contents

1. [First launch](#first-launch)
2. [The cursor and selection](#the-cursor-and-selection)
3. [Editing cells](#editing-cells)
4. [The `/` menu](#the--menu)
5. [Files: open, save, save-as, backup](#files-open-save-save-as-backup)
6. [Formula trace popup (`/T T`)](#formula-trace-popup--t-t)
7. [Keyboard reference card](#keyboard-reference-card)
8. [Not yet implemented](#not-yet-implemented)

---

## First launch

After launching fastsheet you get a blank grid. There's no file open
yet. Either:

- **`/F R`** (File → Retrieve) — opens the file picker. Type to filter
  the listing or recent files. Enter to open.
- **`/W E Y`** (Worksheet → Erase → Yes) — start a new empty workbook.
  (Yes confirms the wipe; defaults to No so a stray `/W E` won't
  destroy unsaved work.)
- **`F5`** with a path — same as `/F R` but skips straight to typing.

When a file is open the title bar shows the filename. A `●` prefix
means the workbook has unsaved changes. In manual recalculation mode,
the status bar separately shows how many edits are waiting for F9.

## The cursor and selection

- **Arrows** — move one cell in any direction.
- **Ctrl + Arrows** — jump to the next non-empty edge in that
  direction. Same as Excel.
- **Shift + Arrows** — extend the selection by one cell.
- **`.`** (period) — cycle the "free corner" of a multi-cell
  selection. Lets you start growing the range from the opposite
  corner of where you started.
- **`F5`** — goto. Accepts `B22`, `Sheet1!CK99`, `'Sheet 1'!A1`, or a
  defined name (e.g. `LeadTimeTable`).
- **Tab / Shift+Tab** — same as → / ←.
- **Page Up / Page Down** — scroll one viewport.
- **Home / End** — start / end of row.
- **Ctrl+Home** — top-left of the sheet.

The status bar at the bottom-right shows count / sum / average for
multi-cell selections, like Excel.

## Editing cells

- **Type a printable character** — starts editing with that character
  as the first keystroke. Replaces existing content.
- **F2** — edit the existing content of the cell. The formula bar
  shows `=...` for formula cells.
- **Enter** — commit, move down.
- **Tab** — commit, move right.
- **Esc** — cancel, restore original content.
- **F4** while editing — cycle the cell reference under the caret
  through `A1 → $A$1 → A$1 → $A1 → A1` (the standard Excel toggle).
- **Backspace / Delete** — clear the selected cells' contents (when not
  editing).

### F2 edit highlights

When you edit a formula (anything starting with `=`), every cell,
range and **named range** the formula references is outlined on the
grid in a different color. Named ranges also show a small tag with
the name above the highlight box.

This updates live as you type — useful for catching `LeadTimeTabe`
typos vs the actual `LeadTimeTable`.

## The `/` menu

Press `/` to open the Lotus-style menu bar. Each item has a single
capital letter you press to descend; **Esc** backs out one level.
The description bar under the menu shows what the highlighted item
will do.

### Top-level

```
/W   Worksheet     Insert/delete rows+cols, hide, freeze, sheets
/R   Range         Format, name, erase, sort, find/replace
/C   Copy          Copy selection to a destination
/M   Move          Move selection to a destination
/F   File          Retrieve, save, backup
/P   Print         (not implemented)
/G   Graph         (not implemented)
/D   Data          Fill, sort, what-if (partial)
/T   Trace         Formula dependency tools (fastsheet-only)
/Q   Quit          Exit
```

### `/W` Worksheet

```
/W I C    Insert columns at cursor (selection width = count)
/W I R    Insert rows at cursor (selection height = count)
/W I I    Insert cells, shift selected rows right
/W I D    Insert cells, shift selected columns down
/W D C    Delete columns
/W D R    Delete rows
/W D L    Delete cells, shift selected rows left
/W D U    Delete cells, shift selected columns up
/W C S    Column Set-Width — pick cols, type width in px or "auto"
/W C A    Column Auto-fit
/W C H    Hide current column
/W C D    Display (un-hide) all columns on this sheet
/W R S    Row Set-Height — pick rows, type height in px or "auto"
/W R A    Row Auto-fit
/W R H    Hide current row
/W R D    Display (un-hide) all rows
/W E Y    Erase entire worksheet (with Y/N confirm)
/W T B    Titles Both — freeze rows above + cols left of cursor
/W T H    Titles Horizontal — freeze rows above
/W T V    Titles Vertical — freeze cols left
/W T C    Titles Clear — unfreeze
/W S N    Sheet New — append a new sheet
/W S D    Sheet Delete — remove current sheet (with confirm)
/W S R    Sheet Rename — rename current sheet
/W G R    Global Recalculation — Automatic / Manual / Now
/W G      Other global settings     ⚠ not yet
/W W      Window split / unsplit    ⚠ not yet
/W P      Page break settings       ⚠ not yet
```

### `/R` Range

```
/R F F    Format Fixed — N decimals
/R F C    Format Currency — N decimals, $ prefix
/R F ,    Format Comma — N decimals, thousands separator
/R F P    Format Percent — N decimals
/R F D    Format Date — yyyy-mm-dd
/R F T    Format Time — h:mm:ss
/R F G    Format General — reset
/R F B    Format Background — fill colour (#RRGGBB or empty to clear)
/R F X    Format Text — text colour
/R F R A  Border All — every side of every cell
/R F R O  Border Outline — outer edges only
/R F R T  Border Top
/R F R B  Border Bottom
/R F R L  Border Left
/R F R R  Border Right
/R F R N  Border None — remove all
/R L L    Label Left — text alignment
/R L C    Label Center
/R L R    Label Right
/R L G    Label General — reset (numbers right, text left)
Ctrl+1    Format Cells dialog — number, font, border, fill, alignment,
          vertical alignment, and wrap text
/R E      Erase selected cell contents (preserves formatting)
/R C      Clear Formats — clear formatting from the selected cells
/R D      Clear All — clear contents and formatting from the selected cells
/R N C    Name Create — define a name for the current selection
/R N D    Name Delete — delete a defined name
/R N L    Name List — drop the names list into the worksheet
/R V      Value — convert formulas in the selection to literal values
/R T      Trans — transpose the selection in place
/R M      Merge — merge the selected cells
/R G      Unmerge — unmerge any merged cells overlapping the selection
/R S      Search — find and (optionally) replace within the sheet
/R J      Justify text in the selected range
/R P      Protect a range              ⚠ not yet
/R U      Unprotect a range            ⚠ not yet
/R I      Restrict input               ⚠ not yet
```

### `/C` Copy and `/M` Move

```
/C        Copy. Prompts for destination top-left, paste preserves
          the source.
/M        Move. Prompts for destination top-left, source is cleared.
```

Both use the active selection as the source. Arrow keys steer the
destination cursor while the menu is in copy/move mode; Enter
commits, Esc cancels.

### `/F` File

```
/F R      Retrieve — open file picker. .xls and .xlsx supported.
/F S      Save — see "Files" below for the picker dialog
/F C O    Compare against another file — opens a docked diff list
/F C X    Compare exit — close the comparison and clear the dock
/F L      List worksheet files in the directory
/F J      Combine — paste the first sheet of another workbook at the cursor
/F X      Xtract — save the selected range to a new .xlsx file
/F E      Erase a file from disk (with confirm)
/F I      Import a text/CSV/TSV file into cells
/F D      Change directory
/F A      Admin                  ⚠ not yet
```

#### Compare mode

`/F C O` picks a second workbook and diffs it against the active one
(values + formulas; sheets matched by name). Each cell where the
formatted display value or formula text differs gets highlighted in
the grid (red = value, blue = formula, gold = missing on one side)
and listed in a docked panel on the right.

While compare is active:
- `↑ ↓` move through the diff list
- `Enter` jump cursor to the highlighted diff
- `← / →` collapse / expand the current row's sheet group
- `* / /` expand-all / collapse-all sheet groups
- `V` cycle filter: all → value → formula → other → all. The
  header row also has clickable filter buttons with per-bucket
  counts. Buckets:
  - **value**: neither side has a formula; the literal values
    differ.
  - **formula**: the formula text actually differs (e.g. `=A1+B1`
    vs `=A1+B2`). A formula on one side and a literal on the
    other counts here too.
  - **other**: both sides have the *same* formula but it
    evaluates to different values — usually a downstream symptom
    of an upstream input change.
- `H` hide the panel — keeps compare highlights on the grid but
  returns the keyboard to the grid (useful to type around)
- `Esc` exit compare mode (same as `/F C X`)
- The `/T T` trace popup picks up right-side values: each dep
  renders `left | right`, and rows where the two sides disagree are
  flagged in red/green so you can drill into a `#N/A` and see which
  branch produced it on each side.

Sheet matching is by name. Sheets that exist on only one side appear
as headers at the top of the diff list. Style/format/width
differences are not reported — only values and formulas are in
scope.

### `/P` Print

⚠ Not yet implemented. Use Excel for printing.

### `/G` Graph

⚠ Not yet implemented. Use Excel for charts.

### `/D` Data

```
/D F      Fill — fill selection with an arithmetic progression
/D S      Sort — sort selected rows by a column
/D P      Parse — split a selected column into adjacent cells
/D T      Table         ⚠ not yet
/D Q      Query         ⚠ not yet
/D D      Distribution  ⚠ not yet
/D M      Matrix        ⚠ not yet
/D R      Regression    ⚠ not yet
```

### `/T` Trace (fastsheet-only)

The killer feature for figuring out why a formula is broken without
opening Excel.

```
/T T      Trace — popup showing the full dependency chain of the
          current cell's formula
/T G      Goto — pick a top-level dependency of the current cell and
          jump to it
/T N      Names — browse all defined names with their resolved
          locations and jump to one
```

See [Formula trace popup](#formula-trace-popup--t-t) below for
keyboard reference inside the popup.

### `/Q` Quit

```
/Q Y      Quit
/Q N      Cancel
```

## Files: open, save, save-as, backup

### Open

`/F R` opens the file navigator. The first view shows recently opened
files (most-recent first) and the last 7 directories you've used.
Enter on a file opens it; Enter on a directory descends into it.
`..` goes up. `\\` jumps to the WSL UNC root if you're on Windows.

Type to filter across the listing, recent files, and recent dirs.
As soon as you cross into a different directory, both recent lists
collapse and only the current directory's entries show.

If the current workbook has unsaved changes, fastsheet asks before
opening another file, starting a new workbook, or quitting.

### Save

`/F S` behaviour depends on the workbook's history:

- **Brand-new workbook** (never saved) → opens the Save As navigator.
- **Existing path that exists on disk** → shows a small picker:
  ```
  Replace / Save As / Backup / Cancel
  ```
  - **Replace** — overwrite the existing file. ⚠ For .xlsx with
    charts/pivots/drawings, Replace uses the in-place byte patcher so
    those features are preserved. For .xls with macros, the
    VBA storage subtree is preserved.
  - **Save As** — pick a different path via the navigator.
  - **Backup** — copy the existing file to `.bak` / `.bak.N`, then save.
  - **Cancel** — do nothing.
- **Existing path that no longer exists on disk** → straight save (no
  picker — nothing to overwrite).

### What's preserved on save

| Feature | .xlsx | .xls |
|---|:-:|:-:|
| Cell values + formulas | ✓ | ✓ |
| Number formats, fonts, fills, borders, alignment | ✓ | ✓ |
| Defined names | ✓ | ✓ |
| Merged cells, frozen panes, col widths, row heights | ✓ | ✓ |
| Hidden rows / cols | ✓ | ✓ |
| Charts | ✓ (Replace only) | ✗ |
| Pivot tables | ✓ (Replace only) | ✗ |
| Drawings, comments, conditional formatting | ✓ (Replace only) | ✗ |
| VBA / macros | ✓ (Replace only) | ✓ |

For .xlsx, the "Replace" path patches the original file's bytes
in-place, leaving everything we don't model untouched. For Save As,
we go through IronCalc's xlsx writer, which doesn't know about
charts / pivots / drawings (so they're lost). The save picker warns
you about this when relevant.

For .xls, fastsheet's BIFF8 writer is greenfield (we synthesise a
fresh file from the model). VBA / macros are preserved by capturing
the original storage subtree on load and replaying it on save —
self-contained from the rest of the file. Charts / pivots / drawings
are NOT yet preserved (would require BIFF-record-offset preservation;
tracked in `CLAUDE.md` pending list).

## Formula trace popup (`/T T`)

The trace popup shows the full dependency chain of the current
cell's formula — every cell, range, and named range it references,
recursing into formulas inside those cells. Same data you'd get from
manually clicking through Excel for half an hour.

### Keys inside the popup

```
↑ / ↓        Move highlight up / down
← / →        Collapse / expand the highlighted node
*            Expand all
/            Collapse all
Enter        Jump to the highlighted cell / named range — closes
             popup and moves cursor there
H            Hide — collapse the popup to a tiny status bar
             ("Trace: G4 …") so you can interact with the grid.
             Press H again to bring the popup back.
D            Dock / undock — toggle between centered modal and
             right-side panel.
Esc          Close trace and restore the cursor to where it was
             when you opened the popup.
```

As you arrow up and down through the list, the grid behind the
popup automatically:

- Switches to the relevant sheet.
- Scrolls so the highlighted cell is visible.
- Outlines the cell or range in orange (with the name as a tag for
  named-range entries).

Your original cursor position is preserved — Esc returns you to it,
or Enter commits the jump.

### What the icons mean

```
•   regular cell (or literal — no formula)
ƒ   defined name
▦   range
⚠   error value (#N/A, #VALUE!, etc.)
↺   circular reference (already on the path — recursion stopped)
…   depth cap — recurse from this row to see further
```

## Keyboard reference card

```
NAVIGATION
  Arrows                Move one cell
  Ctrl+Arrows           Jump to next non-empty edge
  Shift+Arrows          Extend selection
  F5                    Goto address or named range
  Page Up / Down        Scroll one viewport
  Home / End            Row start / end
  Ctrl+Home             Top-left
  Tab / Shift+Tab       → / ←
  .                     Cycle selection's free corner

EDITING
  printable char        Start editing with that character
  F2                    Edit existing content
  Enter / Tab           Commit and move
  Esc                   Cancel edit
  F4                    Cycle ref absolute/relative under caret
  Backspace / Delete    Clear selected cells

UNDO / REDO
  Ctrl+Z                Undo
  Ctrl+Y / Ctrl+Shift+Z Redo

CLIPBOARD
  Ctrl+C / Ctrl+X       Copy / cut selection
  Ctrl+V                Paste from clipboard

  External clipboard uses display-value TSV. Pasting back inside
  fastsheet from a fastsheet copy preserves raw formulas and adjusts
  relative references to the new location.

RECALC + SAVE
  F9                    Recalculate workbook
  Ctrl+S                /F S (save flow)

MENUS
  /                     Open Lotus menu
  Esc                   Back out one menu level (or close)

TRACE POPUP
  H                     Hide / show
  D                     Dock / undock
  ↑↓                    Move highlight
  ←→                    Fold / unfold
  *  /                  Expand all / collapse all
  Enter                 Jump to highlighted cell
  Esc                   Close trace

SHEETS
  Ctrl+PageUp/PageDown  Switch sheet
  Click tab             Switch sheet
  Right-click tab       Sheet context menu
```

## Not yet implemented

The Lotus menu tree is fully populated with the original 1-2-3
items, but several are stubs that print "Not yet implemented" to the
status bar when you select them:

```
/W G       Worksheet/Global settings other than Recalculation
/W W       Worksheet/Window (split / unsplit panes)
/W P       Worksheet/Page (print breaks)
/R J       Range/Justify
/R P       Range/Protect
/R U       Range/Unprotect
/R I       Range/Input restriction
/F J       File/Combine
/F X       File/Xtract (extract)
/F E       File/Erase from disk
/F L       File/List
/F I       File/Import text
/F D       File/Change directory
/F A       File/Admin
/P *       Print/* (entire branch)
/G *       Graph/* (entire branch)
/D T       Data/Table (what-if)
/D Q       Data/Query
/D D       Data/Distribution
/D M       Data/Matrix
/D R       Data/Regression
/D P       Data/Parse
/S         System (shell out)
```

## Beyond the menu

A few features don't have menu entries yet but are in the codebase:

- **Hidden columns / rows are preserved** on save for .xlsx (stored
  via the side-channel `state.hidden_cols`).
- **MY*-style array UDFs** (`MYUNIQUE`, `MYSORT`, `MYFILTER`,
  `MYSORTBLANK`, `MYTRANSPOSE`) — fastsheet emulates these natively
  for .xls files that use them as pseudo-arrays. Names hard-coded;
  see `vendor/base/src/functions/fastsheet_udfs.rs`.
- **Phase timing** — set `FASTSHEET_PROFILE_LOAD=1` in the
  environment before launching, then check `fastsheet_profile.log`
  next to the .exe for per-phase load/evaluate/boot timings.
