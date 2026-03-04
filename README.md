# OMR Grader App

Desktop app to grade OMR `.bmp` sheets using one `.xlsx` answer key.

## Output
Each run creates one Excel file named like:
- `YYYY-MM-DD_HH-MM-SS-result.xlsx`

Excel columns (in order):
1. `Student Name`
2. `Final Score`
3. `Correct`
4. `Wrong`
5. `Blank`

---

## For Mom: Run the Packaged App (No Coding)

### On Windows laptop
1. Receive the app folder (or `.zip`) from you.
2. If zipped, right-click and choose **Extract All**.
3. Open the extracted folder.
4. Double-click `OMR-Grader.exe`.
5. If Windows SmartScreen appears:
   - Click **More info**
   - Click **Run anyway**
6. In the app:
   - Pick `📄 Answer Sheet (.xlsx)`
   - Pick `📁 OMR Images Folder (.bmp)`
   - Pick `💾 Save Result As (.xlsx)`
   - Click `✨ Generate Result`
7. Open the saved Excel file from the path shown after completion.

### On macOS laptop
1. Open the provided `OMR-Grader` app/folder.
2. If macOS blocks it (first run):
   - Go to **System Settings > Privacy & Security**
   - Click **Open Anyway** for the app
3. Then use the same 4 steps inside the app:
   - Answer sheet
   - OMR folder
   - Save location
   - Generate

---

## For You: Build the App

Important: build on the same OS you will distribute to.
- Build on Windows for Windows `.exe`
- Build on macOS for macOS app

### 1) Install dependencies
```bash
pip install -r requirements.txt
pip install pyinstaller
```

### 2) Build
```bash
pyinstaller --noconfirm --windowed --name OMR-Grader app.py
```

### 3) Share
- Windows: share `dist/OMR-Grader.exe`
- macOS/Linux: share `dist/OMR-Grader`

---

## Run from Source (Developer)
```bash
pip install -r requirements.txt
python app.py
```
