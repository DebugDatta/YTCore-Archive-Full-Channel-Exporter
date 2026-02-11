
---

# 🎥 YTCore Archive – Full Channel Exporter

YTCore Archive converts an entire YouTube channel into a structured, navigable Excel archive. It captures playlists, videos, shorts, and streams into a clean, organized spreadsheet.

🌐 **Live App:**
[https://ytcore.streamlit.app/](https://ytcore.streamlit.app/)

---

## ✨ Features

* Extracts all playlists
* Scans `/videos`, `/shorts`, `/streams`
* Detects uncategorized videos
* Sorts by publish date (newest first)
* Calculates duration per category
* Multi-threaded processing
* Clickable Excel links
* Back-to-summary navigation
* Auto-adjusted columns

---

## 📊 Excel Output

**Channel_Info**

* Export date
* Total duration

**Summary**

* Category
* Video count
* Duration
* Clickable links

**Playlist / Category Sheets**

* Video Title
* Duration (HH:MM:SS)
* Video Link
* Publish Date (YYYY-MM-DD)

---

## 🚀 Run Locally

```bash
git clone https://github.com/DebugDatta/YTCore-Archive-Full-Channel-Exporter.git
```

```bash
cd YTCore-Archive-Full-Channel-Exporter
```

```bash
pip install -r requirements.txt
```

```bash
python -m streamlit run app.py
```

---

## 📦 Requirements

```
streamlit
yt-dlp
pandas
openpyxl
```

---

## ⚠ Limitations

* Private / members-only videos not accessible
* Region-restricted videos may fail

---

