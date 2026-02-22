# Substack Archiver GUI

A 100% vibe-coded GUI Windows application that downloads Substack newsletter archives to your computer locally, and optionally converts them into an ePub file suitable for reading on a Kindle, Kobo, or any other e-reader.

This is just a vibe-coded GUI - the real work is being done behind the scenes by two excellent open-source tools that other authors have very generously made available — [sbstck-dl](https://github.com/alexferrari88/sbstck-dl) and [Pandoc](https://pandoc.org/).

All this repo's app does is present a GUI that builds and runs the commands for you, showing you exactly what it is doing in real time.

No command-line knowledge is therefore required once everything has been set up the first time.

I have only tested it on a Windows 11 computer. AI wrote most of it. AI also wrote most of this readme. It comes with no functionality or security guarentees whatsoever. Caveat emptor.

<img width="828" height="1024" alt="ss_main" src="https://github.com/user-attachments/assets/6ac17ebd-1c17-4036-b06c-7e1fd1eb0a7c" />


---

## What It Does

1. **Downloads** every post from any Substack newsletter (free or paid) to a folder on your computer, in Markdown, HTML, or plain text format — including images and file attachments if you want them.
2. **Converts** those downloaded files into a single ePub book, with the posts in chronological order and an optional table of contents.


---

## Requirements

You need to install three things before running the app:

### 1. Python 3.10 or later

Download from [python.org](https://www.python.org/downloads/). During installation, make sure you tick **"Add Python to PATH"**.

To check you already have it:
```
python --version
```

### 2. sbstck-dl

A free tool that downloads Substack newsletters. Install it by opening a Command Prompt and running:

```
pip install sbstck-dl
```

Source and documentation: [github.com/alexferrari88/sbstck-dl](https://github.com/alexferrari88/sbstck-dl)

### 3. Pandoc

A free document converter used to build the ePub file. Download the Windows installer from:

[pandoc.org/installing.html](https://pandoc.org/installing.html)

The default installation path (`C:\Program Files\Pandoc\pandoc.exe`) is already pre-filled in the app's Settings tab.

Alternatively, install via Windows Package Manager:
```
winget install --id JohnMacFarlane.Pandoc
```

---

## Installing the App

1. **Download or clone this repository** to a folder on your computer.

2. **Install the Python GUI dependencies** by opening a Command Prompt in the project folder and running:

```
pip install customtkinter
```

3. **Run the app:**

```
python main.py
```

That's it. A settings file (`config.json`) will be created automatically in the same folder on first run.

### Optional: Create a desktop shortcut

Right-click `main.py` in Windows Explorer → **Send to → Desktop (create shortcut)**.
Then right-click the shortcut → **Properties** → change "Open with" to `pythonw.exe` (instead of `python.exe`) to avoid the black console window appearing when you launch from the shortcut.

---

## Using the App

The app has three tabs: **Download**, **ePub Conversion**, and **Settings**, plus a log panel at the bottom that shows live output from any running command.

---

### Download Tab

Use this tab to download a Substack newsletter archive to your computer.

<img width="828" height="1024" alt="ss_main" src="https://github.com/user-attachments/assets/703e635e-de26-4318-adcd-34f4335721a6" />

#### Source — Substack URL

Enter the full URL of the Substack newsletter you want to archive, e.g.:

```
https://example.substack.com/
```

#### Destination

- **Output folder** — Choose where the downloaded files will be saved. Use the Browse button.
- **Format** — Choose the file format for downloaded posts:
  - **Markdown (.md)** — Best choice if you plan to convert to ePub afterwards. This is the recommended format.
  - **HTML (.html)** — Good for reading in a browser.
  - **Plain Text (.txt)** — Strips all formatting.

#### Date Range (optional)

Tick "Filter by date range" to download only posts published within a specific window. Set an **After** date, a **Before** date, or both. Leave unticked to download everything.

#### Image Options

- **Download images locally** — When ticked, images are downloaded to your computer and links in the files are updated to point to the local copies. Without this, images remain as online links and will not be available if you read offline.
  - **Quality** — Choose `low` (424px, smallest files), `medium` (848px), or `high` (1456px, largest files). `low` is recommended for e-readers.
  - **Images subfolder** — The name of the subfolder where images are saved (default: `images`). You rarely need to change this.

#### File Attachments

- **Download file attachments** — Tick this to also download any attached files (PDFs, audio, etc.).
  - **Extensions** — Optionally filter to specific file types, e.g. `pdf,mp3`. Leave blank to download all attachment types.
  - **Files subfolder** — Where attachments are saved (default: `files`).

#### Advanced Options

| Option | Description |
|--------|-------------|
| Add source URL to each post | Appends the original Substack URL at the bottom of each downloaded file — useful for reference |
| Create archive index page | Generates an `index.md` or `index.html` file linking all posts — useful for browser-based browsing |
| Verbose output | Shows more detail in the log |
| Dry run | Builds and displays the command but **does not actually download anything** — useful for checking your settings |
| Rate limit | How many requests per second to send (default: 1). Increase with caution; very high rates may get you temporarily blocked |

#### Paid Content Authentication

If you want to download posts from a Substack you **pay for**, you need to provide your session cookie. This proves to Substack's servers that you are a paying subscriber.

<img width="991" height="245" alt="cookie" src="https://github.com/user-attachments/assets/ee3c516b-1e2b-4b1e-a69c-6001c379e3b7" />

Click **Show Cookie Settings** to expand this section.

**How to find your cookie:**

1. Log into [substack.com](https://substack.com) in your browser
2. Press **F12** to open Developer Tools
3. Go to **Application** → **Cookies** → `https://substack.com`
4. Find the cookie named **`substack.sid`** (or `connect.sid`)
5. Copy its **Value** — it will be a long string of letters and numbers

Paste that value into the **Cookie Value** field. The value is masked by default (like a password); click **Show** to reveal it for checking.

> ⚠️ **Security note:** Your cookie acts like a password for your Substack account. The app never saves it to disk — you will need to paste it in each session.

#### Command Preview

The read-only box at the bottom of the Download tab shows the exact command that will be run when you click **Start Download**. It updates live as you change settings. This is useful for understanding what the app is doing, or for copying the command if you ever want to run it manually.

#### Starting the Download

Click **Start Download**. The button will grey out while the download runs, and output from `sbstck-dl` will appear in the log panel at the bottom of the window.

When the download completes, a dialog appears with three options:


| Button | Action |
|--------|--------|
| **Create ePub →** | Switches to the ePub Conversion tab with the source folder already filled in |
| **Show Files** | Opens Windows Explorer in the folder where files were saved |
| **Return to App** | Dismisses the dialog |

If you used **Dry run**, only "Return to App" is shown (since no files were downloaded).

---

### ePub Conversion Tab

Use this tab to convert a folder of Markdown files into a single ePub book.

<img width="1042" height="1020" alt="epub" src="https://github.com/user-attachments/assets/590b6083-0d18-4136-aeb3-5a7e3a41c559" />

#### Source

- **Source folder** — The folder containing the `.md` files you downloaded. If you clicked "Create ePub →" from the download completion dialog, this is already filled in.

The **Files found** box below lists all the `.md` files detected in the folder (excluding `index.md`), sorted alphabetically. Because `sbstck-dl` names files by date, alphabetical order equals chronological order — so the first file in the list is the oldest post and the last is the newest. This is the order they will appear in the ebook.

#### Destination

- **Output .epub file** — Where to save the finished ePub. Use the Browse button to choose a location and filename. The file must end in `.epub`.

#### Book Metadata

- **Book title** — The title that will appear on the ebook's cover and in your e-reader's library.
- **Author** — The author name that will appear in your e-reader's library.

#### ePub Options

- **Include Table of Contents** — Adds a clickable table of contents to the ebook, with one entry per article. Recommended.
- **Chapter split level** — Controls how Pandoc divides the content into chapters. The default of `1` means each article becomes its own chapter. You rarely need to change this.

#### Command Preview

Shows the Pandoc command that will be run, abbreviated to keep it readable (showing the first file and then the total count, rather than listing every file).

#### Converting

Click **Convert to ePub**. Pandoc will process all the Markdown files and produce a single `.epub` file. Large archives with many images may take a minute or two.

When complete, a dialog appears:

| Button | Action |
|--------|--------|
| **Show File** | Opens Windows Explorer in the folder containing the ePub |
| **Return to App** | Dismisses the dialog |

---

### Settings Tab

#### Executable Paths

If `sbstck-dl` and `pandoc` are on your system PATH (which they will be after a standard installation), you can leave both fields blank and the app will find them automatically.

If the app cannot find them, use the Browse buttons to locate the executable files:

| Tool | Typical Windows location |
|------|--------------------------|
| sbstck-dl | `C:\Users\YourName\AppData\Local\Programs\Python\Python3xx\Scripts\sbstck-dl.exe` |
| pandoc | `C:\Program Files\Pandoc\pandoc.exe` |

Click **Save Settings** to save your paths. They are stored in `config.json` in the app folder and will be remembered between sessions.

---

### The Log Panel

The log panel at the bottom of the window is always visible regardless of which tab you are on. It shows:

- The exact command being run (so you can see what is happening)
- All output from `sbstck-dl` or `pandoc` as it runs, line by line
- Success or error messages when the process finishes

Click **Clear Log** to wipe the log output.

---

## Tips

- **Always use Markdown format** if you plan to convert to ePub. HTML files can be converted but produce less clean results.
- **Start with a dry run** to confirm the command looks right before starting a large download.
- **Low image quality** is recommended for e-readers — it produces much smaller files and the difference is barely visible on e-ink screens.
- The app remembers your last-used URL, output folder, and ePub settings between sessions, so repeated use is faster.
- Your cookie value is **never saved to disk**. You will need to paste it in again each time you open the app.

---

## Troubleshooting

**"Could not find executable: sbstck-dl"**
Run `pip install sbstck-dl` in a Command Prompt, or use the Settings tab to browse to the executable manually.

**"Could not find executable: pandoc"**
Download and install Pandoc from [pandoc.org](https://pandoc.org/installing.html), or use the Settings tab to browse to `pandoc.exe`.

**Download starts but no files appear**
Check the log for error messages. Common causes: the URL is wrong (try copying it directly from your browser), or the Substack is paid and you need to provide a cookie.

**ePub conversion produces an empty or broken file**
Make sure you downloaded in Markdown (`.md`) format. HTML files sometimes cause issues with Pandoc's ePub output.

**The app window is blank or has display issues**
Try running `pip install --upgrade customtkinter` to get the latest version of the GUI library.

---

## Acknowledgements

This app is a low quality vibe-coded GUI wrapper around two excellent open-source tools:

- **[sbstck-dl](https://github.com/alexferrari88/sbstck-dl)** by Alex Ferrari — Substack downloader
- **[Pandoc](https://pandoc.org/)** by John MacFarlane — Universal document converter

---

## Licence

MIT
