# AS-BUILT STAMPER
### Hargreaves / Exentec Hargreaves

Automates stamping DWG drawings as As-Built by adding a new revision row
to the title block, incrementing the revision letter, and setting all
required fields.

---

## What it does

For each DWG file you select:

1. Opens the drawing in AutoCAD (launches AutoCAD if not running)
2. Scans all text entities to find the revision table
3. Identifies the **latest revision row** (highest rev letter)
4. **Adds a new row directly above** the current latest row — same font, same height, same position
5. Sets the new row fields:
   - **Rev** — auto-incremented (A→B, B→C, …, Z→AA, etc.)
   - **Date** — today's date (dd/mm/yyyy) — editable before running
   - **Prep By** — copied from previous row (or enter your own)
   - **Checked By** — copied from previous row (or enter your own)
   - **Status** — `P7` (editable)
   - **Reason** — `As Built` (editable)
   - **Approved** — same as previous row (or enter your own)
6. Saves the modified drawing as `[original name]_ASBUILT.dwg`
7. Exports a PDF as `[original name]_ASBUILT.pdf`

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows 10/11 | COM automation is Windows-only |
| Python 3.9+ | https://python.org — tick "Add to PATH" |
| AutoCAD | Any version 2013–2025 with COM automation enabled |
| pywin32 | Installed via `setup.bat` |

---

## Installation

1. Copy the `asbuilt_stamper` folder to your machine
2. **Right-click `setup.bat` → Run as Administrator**
3. Wait for the install to complete

---

## Running

Double-click **`Run Stamper.bat`**

Or from a command prompt:
```
python asbuilt_stamper.py
```

---

## How to use

1. **Select DWG Files** — click `+ Add DWG Files` and pick one or more drawings
2. **Fill in Revision Fields**:
   - Rev is auto-detected and incremented — shown as `AUTO`
   - Enter your Prep By, Checked By, Approved initials
   - Date defaults to today — change if needed
   - Status defaults to `P7`
   - Reason defaults to `As Built`
3. **Output folder** — defaults to same folder as the DWG; click Browse to change
4. **Dry Run** — tick to analyse the drawing without making changes (useful for testing)
5. Click **▶ STAMP AS-BUILT**

---

## Title block compatibility

The app finds revision rows by looking for:
- Text entities containing single or double letter revision codes (A, B, C… AA, AB…)
- Grouped with other text on the same Y-coordinate (same row)

It works with **plain TEXT entities** in the title block. If your title block uses
**block attributes (ATTDEF/ATTRIB)**, contact your CAD manager to export as plain text,
or the app can be extended to handle attributes.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `pywin32 not installed` | Run `setup.bat` as Administrator |
| `No running AutoCAD found` | AutoCAD will be launched automatically |
| `No revision rows found` | Ensure rev cells are plain TEXT (not attributes). Try the Dry Run to see what text is detected in the log |
| PDF export not working | Use AutoCAD's Publish command manually after stamping |
| Rev letter wrong | The app reads the revision letter from the last row; if your drawing uses numbers or different codes, edit the `_increment_rev` method |

---

## Output files

```
original_drawing_ASBUILT.dwg   ← modified drawing
original_drawing_ASBUILT.pdf   ← PDF export
```

---

*Built for Hargreaves / Exentec Hargreaves — HPC HVAC As-Built workflow*
