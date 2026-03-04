# OMR Grader

A lightweight desktop app that grades OMR `.bmp` sheets using a provided answer-key `.xlsx` file.

## What This Project Does
- Extracts student names from name bubbles in each OMR image
- Extracts selected options for 100 questions
- Scores each student with:
  - `+1` for correct
  - `-0.25` for wrong
  - `0` for blank
- Exports one result Excel file:
  - `YYYY-MM-DD_HH-MM-SS-result.xlsx`
  - Columns: `Student Name`, `Final Score`, `Correct`, `Wrong`, `Blank`

## How It Works (Concise)
This uses classical computer vision + deterministic logic (no deep learning):
- Fixed ROI geometry (name/answer regions), scaled to image size
- Hough circle detection for answer bubbles
- 1D k-means clustering to map bubble centers to `(block, question-row, option)`
- Connected-components for name bubbles, grouped by x-position, then calibrated y-to-letter mapping (`A-Z`)
- Deterministic scoring against the parsed answer key

## Run From Source
1. Install Python 3.10+.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Start app:
   ```bash
   python app.py
   ```
4. In the app:
   - Pick answer sheet (`.xlsx`)
   - Pick OMR folder (`.bmp` images)
   - Pick save path for output Excel (`.xlsx`)
   - Click Generate

## Package App (for sharing)
Build on the same OS as target device.

### Windows build
```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --windowed --name OMR-Grader app.py
```
Output: `dist/OMR-Grader.exe`

### macOS build
```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --windowed --name OMR-Grader app.py
```
Output: `dist/OMR-Grader`

## Repository Notes
Generated outputs are ignored via `.gitignore`:
- `debug_out/`, `out/`, `__pycache__/`, generated `.xlsx/.csv/.json` outputs
