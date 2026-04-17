import yt_dlp
import pandas as pd
import streamlit as st
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl.utils import get_column_letter
from datetime import datetime

# 1. Improved Timestamp Helper
def hms(total_seconds):
    """Converts seconds to HH:MM:SS format."""
    if not total_seconds or total_seconds < 0: 
        return "00:00:00"
    h = int(total_seconds // 3600)
    m = int((total_seconds % 3600) // 60)
    s = int(total_seconds % 60)
    return f"{h:02}:{m:02}:{s:02}"

def clean_name(n, maxlen=150):
    forbidden = [':', '\\', '/', '?', '*', '[', ']', '<', '>', '|']
    for char in forbidden: n = n.replace(char, '')
    return "".join(c for c in n if c.isalnum() or c in " _-").strip()[:maxlen] or "data"

def clean_sheet(n, used):
    base = clean_name(n, 25)
    name, counter = base, 1
    while name.lower() in used:
        suffix = f"_{counter}"
        name = base[:31-len(suffix)] + suffix
        counter += 1
    used.add(name.lower())
    return name

def auto_adjust_columns(ws, max_width=60):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                length = len(str(cell.value)) if cell.value else 0
                if length > max_length: max_length = length
            except: pass
        ws.column_dimensions[col_letter].width = min(max_length + 2, max_width)

def get_ydl_opts(flat=True):
    return {
        'quiet': True, 
        'extract_flat': flat, 
        'skip_download': True, 
        'ignoreerrors': True, 
        'ignore_no_formats_error': True,
        'no_warnings': True
    }

def normalize_channel_url(url: str) -> str:
    if not url: return url
    url = url.strip()
    if "?" in url: url = url.split("?")[0]
    return url.rstrip("/")

@st.cache_data(ttl=3600, show_spinner=False)
def get_channel_info_cached(url):
    with yt_dlp.YoutubeDL(get_ydl_opts(False)) as d:
        return d.extract_info(url, download=False)

def extract_tab(url, tab):
    target_url = url.rstrip('/') + f"/{tab}"
    videos = {}
    with yt_dlp.YoutubeDL(get_ydl_opts(True)) as d:
        try:
            info = d.extract_info(target_url, download=False)
            if info and 'entries' in info:
                for v in info['entries']:
                    if not v: continue
                    vid_id = v.get('id')
                    if not vid_id: continue
                    dur = v.get('duration') or 0
                    date_val = v.get('upload_date')
                    try:
                        dt = datetime.strptime(date_val, "%Y%m%d").date() if date_val else None
                    except: dt = None
                    
                    videos[vid_id] = {
                        "Video Title": v.get('title') or "video",
                        "Duration (HH:MM:SS)": hms(dur),
                        "Video Link": f"https://www.youtube.com/watch?v={vid_id}",
                        "Publish Date": dt,
                        "seconds": dur,
                        "type": tab
                    }
        except: pass
    return videos

@st.cache_data(ttl=3600, show_spinner=False)
def extract_tab_cached(url, tab):
    return extract_tab(url, tab)

def get_all_channel_content(url, progress_bar):
    tabs = ["videos", "shorts", "streams"]
    all_content = {}
    for i, tab in enumerate(tabs):
        source = extract_tab_cached(url, tab)
        all_content.update(source)
        progress_bar.progress(0.20 + (i + 1) * 0.15)
    return all_content

def extract_playlist(pl):
    data, total_dur, ids = [], 0, set()
    with yt_dlp.YoutubeDL(get_ydl_opts(True)) as d:
        try:
            info = d.extract_info(pl['url'], download=False)
            if info and 'entries' in info:
                for v in info['entries']:
                    if not v: continue
                    vid_id = v.get('id')
                    dur = v.get('duration') or 0
                    date_val = v.get('upload_date')
                    try:
                        dt = datetime.strptime(date_val, "%Y%m%d").date() if date_val else None
                    except: dt = None
                    
                    total_dur += dur
                    if vid_id: ids.add(vid_id)
                    data.append({
                        "Video Title": v.get('title') or "video",
                        "Duration (HH:MM:SS)": hms(dur),
                        "Video Link": f"https://www.youtube.com/watch?v={vid_id}" if vid_id else "",
                        "Publish Date": dt
                    })
        except: pass
    return pl['name'], pl['url'], data, total_dur, ids

@st.cache_data(ttl=3600, show_spinner=False)
def extract_playlist_cached(pl):
    return extract_playlist(pl)

def build_excel(results, summary, total_channel, metadata):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Channel Metadata Sheet
        pd.DataFrame(list(metadata.items()), columns=["Field", "Value"]).to_excel(writer, sheet_name="Channel_Info", index=False)
        
        # Summary Sheet
        df_sum = pd.DataFrame({
            "Category": [s[1] for s in summary],
            "Video Count": [s[4] for s in summary],
            "Duration": [hms(s[3]) for s in summary],
            "Link": [s[2] for s in summary]
        })
        df_sum.loc[len(df_sum)] = {"Category": "TOTAL CHANNEL DURATION", "Video Count": "", "Duration": hms(total_channel), "Link": ""}
        df_sum.to_excel(writer, sheet_name="Summary", index=False)
        
        wb = writer.book
        for sn, n, pl_url, d, t in results:
            df = pd.DataFrame(d)
            if not df.empty and "Publish Date" in df.columns:
                df = df.sort_values("Publish Date", ascending=False, na_position='last')
            
            # Remove helper column 'seconds' and 'type' if they exist before writing
            cols_to_drop = [c for c in ['seconds', 'type'] if c in df.columns]
            df.drop(columns=cols_to_drop, inplace=True)
            
            df.to_excel(writer, sheet_name=sn, index=False, startrow=1)
            ws = wb[sn]
            ws["A1"] = "⬅ Back to Summary"
            ws["A1"].hyperlink, ws["A1"].style = "#'Summary'!A1", "Hyperlink"
            
            # Format Hyperlinks in Video Link column
            for r in range(3, len(df) + 3):
                cell = ws[f"C{r}"]
                if cell.value and str(cell.value).startswith("http"):
                    cell.hyperlink, cell.style = cell.value, "Hyperlink"
            auto_adjust_columns(ws)
            
        auto_adjust_columns(wb["Channel_Info"])
        ws_sum = wb["Summary"]
        for i, r in enumerate(summary, 2):
            ws_sum[f"A{i}"].hyperlink, ws_sum[f"A{i}"].style = f"#'{r[0]}'!A1", "Hyperlink"
        auto_adjust_columns(ws_sum)
        
    return output.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="YTCore Archive", layout="wide", page_icon="🎥")
st.title("🎥 YTCore Archive: Full Channel Exporter")

input_url = st.text_input("YouTube Channel URL (e.g., https://www.youtube.com/@ChannelName)")

if st.button("Generate Archive", type="primary"):
    if not input_url:
        st.error("Please enter a URL")
    else:
        clean_url = normalize_channel_url(input_url)
        progress = st.progress(0.05)
        
        with st.spinner("Processing full channel archive..."):
            try:
                # Get Basic Info
                channel_info = get_channel_info_cached(clean_url)
                channel_name = clean_name(channel_info.get("title", "channel"))
                
                # Extract Playlists
                y_playlists = []
                with yt_dlp.YoutubeDL(get_ydl_opts(True)) as d:
                    info = d.extract_info(clean_url + '/playlists', download=False)
                    if info and 'entries' in info:
                        for e in info['entries']:
                            if e:
                                y_playlists.append({'name': e.get('title') or "playlist", 'url': e.get('url') or e.get('webpage_url')})

                # Extract All Content (Videos/Shorts/Streams)
                all_content = get_all_channel_content(clean_url, progress)

                results, summary, playlist_video_ids, used_sheets = [], [], set(), {"summary", "channel_info"}

                # Process Playlists using Threads
                with ThreadPoolExecutor(max_workers=8) as ex:
                    futures = [ex.submit(extract_playlist_cached, p) for p in y_playlists]
                    for f in as_completed(futures):
                        n, pl_url, data, total_d, ids = f.result()
                        if not data: continue
                        playlist_video_ids.update(ids)
                        sn = clean_sheet(n, used_sheets)
                        results.append((sn, n, pl_url, data, total_d))
                        summary.append((sn, n, pl_url, total_d, len(data)))

                # Process Remaining "Orphan" Videos
                for label, tab_type in {"Videos": "videos", "Shorts": "shorts", "Streams": "streams"}.items():
                    items = [v for k, v in all_content.items() if k not in playlist_video_ids and v['type'] == tab_type]
                    if items:
                        sn = clean_sheet(label, used_sheets)
                        dur = sum(i['seconds'] for i in items)
                        results.append((sn, label, clean_url, items, dur))
                        summary.append((sn, label, clean_url, dur, len(items)))

                total_duration_secs = sum(s[3] for s in summary)
                metadata = {
                    "Channel Name": channel_name,
                    "Channel URL": clean_url,
                    "Export Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Total Duration": hms(total_duration_secs)
                }

                # Build File
                excel_data = build_excel(results, summary, total_duration_secs, metadata)

                # UI Display
                st.success("Archive Generated Successfully!")
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Channel Statistics")
                    st.table(pd.DataFrame(list(metadata.items()), columns=["Field", "Value"]))
                with col2:
                    st.subheader("Summary per Tab")
                    st.dataframe(pd.DataFrame(summary, columns=["Sheet", "Category", "URL", "Seconds", "Count"]).drop(columns=["Seconds"]))

                st.download_button(
                    label="📥 Download Excel Archive",
                    data=excel_data,
                    file_name=f"{channel_name}_archive_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                progress.progress(1.0)

            except Exception as e:
                st.error(f"An error occurred: {e}")

if st.sidebar.button("Clear Cache"):
    st.cache_data.clear()
    st.sidebar.success("Cache cleared")
