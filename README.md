# BhugtanEase — EXE Banane Ki Guide

## Yeh kyu zaroori hai?
Windows ka `.exe` file sirf Windows par hi banta hai (Linux/Mac se direct "cross-build"
nahi ho sakta). Isliye yeh poora "build kit" aapko apne ya kisi bhi **Windows PC** par
chalana hai — bas 1 click mein `BhugtanEase.exe` ban jayegi.

## Is folder mein kya hai
- `BhugtanEase_Final.pyw` → aapka asli software (source code)
- `BhugtanEase__Logo.png` → aapka logo
- `BhugtanEase.ico` → wahi logo, EXE icon format (.ico) mein already convert kiya hua
- `requirements.txt` → zaroori Python libraries ki list
- `build.bat` → 1-click build script (EXE + Installer dono banata hai)
- `installer.iss` → installer ka blueprint (Inno Setup script)
- `build_installer.bat` → sirf installer (Setup.exe) banane ka script

## Option B: GitHub se banao (Windows PC ki zaroorat NAHI)

Agar aapke paas Windows PC nahi hai, GitHub Actions cloud mein khud EXE +
Installer bana dega — bilkul free.

1. https://github.com par jaake naya **repository** banao (private ya public, jo chaho).
2. Is poore folder ka content us repo mein upload/push karo. Sabse aasan tarika:
   - GitHub repo page par "Add file" → "Upload files" → yeh saari files/folders
     (`.github` folder bhi, ismein workflow chhupi hai) drag-drop karo → "Commit".
3. Repo ke top par **"Actions"** tab par click karo.
   - Workflow khud-ba-khud chalna shuru ho jayega ("Build BhugtanEase EXE + Installer").
   - Agar nahi chala, to "Run workflow" button dabao.
4. 3-5 minute wait karo (status "green tick ✅" hone tak).
5. Us workflow run ke andar **"Artifacts"** section mein 2 zip milenge:
   - `BhugtanEase-EXE` → isme `BhugtanEase.exe` hai
   - `BhugtanEase-Setup-Installer` → isme `BhugtanEase_Setup.exe` (proper installer) hai
6. Jo zip chahiye usko download karke andar ki `.exe` file customer ko de do.

> Note: GitHub Actions Python 3.8 use karta hai (Windows 7 compatibility ke liye).
> Agar Windows 7 support zaroori nahi hai, `.github/workflows/build.yml` file mein
> `python-version: "3.8"` ko `"3.11"` kar sakte ho.

---

## Option A: Apne Windows PC par banao

1. **Python install karo** (agar already nahi hai):
   - https://www.python.org/downloads/
   - Install karte waqt neeche diya hua box **"Add python.exe to PATH"** zaroor tick karo.

   ### Windows 7 compatibility chahiye to ⚠️ IMPORTANT:
   - Python **3.8.x** install karo (yeh aakhri version hai jo Windows 7 par chalta hai).
     Direct link: https://www.python.org/downloads/release/python-3810/
     ("Windows installer (64-bit)" download karo)
   - Agar aapko sirf Windows 8/10/11 par chalana hai (Windows 7 zaroori nahi),
     to koi bhi naya Python (3.9 se 3.12 tak) chalega.

2. **Inno Setup install karo** (installer/Setup.exe banane ke liye, free hai):
   - https://jrsoftware.org/isdl.php
   - "Inno Setup 6.x" download karke default settings se install kar do.

3. Yeh poora folder (`build` folder ke andar ki saari files) Windows PC par copy karo.

4. `build.bat` par **double-click** karo.
   - Yeh khud venv banayega, sab library install karega, EXE banayega,
     aur uske baad automatically installer (`Setup.exe`) bhi bana dega.
   - 2-5 minute lag sakte hain (internet chahiye, kyunki libraries download hoti hain).

5. Build complete hone ke baad:
   - Isi folder mein **`BhugtanEase.exe`** (single file, bina installer wala)
   - `installer_output` folder ke andar **`BhugtanEase_Setup.exe`** (proper installer —
     wizard, Start Menu/Desktop shortcut, Uninstaller sab saath)

   Customer ko inme se jo bhi chahiye wo de sakte ho — zyada professional dena ho to
   **`BhugtanEase_Setup.exe`** dena behtar hai.

## Customer ke PC par chalne ke liye kya chahiye?
Kuch bhi install karne ki zaroorat NAHI hai — `--onefile` build mein Python aur
saari libraries EXE ke andar hi bundled hoti hain.
- `BhugtanEase.exe` → bas double-click, seedha chal jayega.
- `BhugtanEase_Setup.exe` → double-click karega, install wizard khulega
  ("Next, Next, Install"), Desktop/Start Menu shortcut ban jayega, aur
  Control Panel se uninstall bhi kar sakega — ekdum professional software jaisa.

> Note: Windows kabhi-kabhi naye/unknown EXE par "Windows Defender SmartScreen"
> warning dikhata hai ("More info" → "Run anyway" se chal jata hai) — yeh normal hai
> kyunki EXE digitally signed nahi hai. Agar yeh warning bhi nahi chahiye, to aapko
> ek paid Code Signing Certificate lena padega (alag topic hai, zaroorat ho to bata dena).

## Agar build.bat mein koi error aaye
Error message ka screenshot bhejo, main fix bata dunga. Common issues:
- "pip install" mein internet slow → dobara try karo
- pywin32 postinstall warning → usually ignore kar sakte ho, exe phir bhi ban jati hai
