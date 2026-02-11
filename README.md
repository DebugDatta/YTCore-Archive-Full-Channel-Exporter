
# 🎥 YTCore Archive – Full Channel Exporter

YTCore Archive turns YouTube channels into organized Excel files. It captures every playlist, short, and live stream so nothing is missed. The result is a clean spreadsheet with video durations, publish dates, and clickable links for easy navigation.

Perfect for tracking courses, research, or archiving content.

---

## 📦 Repository Structure

```
app.py
requirements.txt
README.md
```

* **app.py** – Main Streamlit application
* **requirements.txt** – Required Python dependencies

---

## ✨ What It Does

* Extracts all playlists from a channel
* Scans `/videos`, `/shorts`, `/streams`
* Detects videos not assigned to playlists
* Sorts videos by publish date (newest first)
* Calculates duration per category
* Generates a structured Excel archive

---

## 📊 Excel Output Includes

### Channel_Info

* Export tool
* Export date
* Total duration

### Summary

* Category name
* Video count
* Total duration
* Clickable navigation links

### Playlist / Category Sheets

* Video Title
* Duration (HH:MM:SS)
* Video Link (clickable)
* Publish Date (YYYY-MM-DD)
* Back to Summary navigation

---

## 🚀 Installation & Run

### Clone the Repository

```bash
git clone https://github.com/DebugDatta/YTCore-Archive-Full-Channel-Exporter.git
cd YTCore-Archive-Full-Channel-Exporter
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Run the App

```bash
python -m streamlit run app.py
```

Open the local URL shown in the terminal.

---

## 🛠 Tech Stack

* Python 3.8+
* Streamlit
* yt-dlp
* pandas
* openpyxl
* concurrent.futures

---

## ⚠ Limitations

* Private or members-only videos cannot be accessed
* Region-restricted videos may fail
* Metadata accuracy depends on YouTube availability

---

