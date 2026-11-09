"""
Instagram Post & Reel IST Timestamp Tracker (MVP Dashboard + Excel Exporter)
=============================================================================
Features:
- Accepts an Instagram username (e.g., `natgeo`), `@handle`, or profile URL (`https://www.instagram.com/natgeo/`).
- Uses the free, open-source `instaloader` Python module (https://instaloader.github.io/as-module.html)
  to iterate through all profile posts and Reels (79 posts, 500+ posts, etc.).
- Converts every post's UTC creation timestamp (`post.date_utc`) to Indian Standard Time (IST / UTC+5:30).
- Displays an interactive Streamlit dashboard (https://docs.streamlit.io/) with media-type filtering
  (Reels/Videos vs. All Posts) and posting-time analytics.
- Exports a formatted Excel (.xlsx) workbook with clickable post links and auto-fitted columns.
- Can also be run directly from the terminal (CLI mode) or Google Colab to generate the Excel file.

Quick Start:
1. Install dependencies:
   pip install streamlit instaloader pandas openpyxl
2. Run the dashboard:
   python -m streamlit run instagram_ist_tracker_app.py
3. (Optional) Run via CLI directly to Excel:
   python instagram_ist_tracker_app.py <username_or_url>
"""

import argparse
from datetime import timezone, timedelta
from io import BytesIO
import sys
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd

# Define India Standard Time (UTC+05:30)
IST_TIMEZONE = timezone(timedelta(hours=5, minutes=30), name="IST")


def safe_int(val: Any, default: Optional[int] = 0) -> Optional[int]:
    """Safely converts a value to int, returning `default` if None or invalid."""
    if val is None:
        return default
    try:
        return int(val)
    except (ValueError, TypeError):
        return default


def safe_attr(obj: Any, attr: str, default: Any = None) -> Any:
    """Safely reads an attribute from an Instaloader object without raising key/network errors."""
    try:
        val = getattr(obj, attr, default)
        return default if val is None else val
    except Exception:
        return default


def extract_username(raw_input: str) -> str:
    """
    Extracts a clean Instagram username from a raw username, @handle, or full Instagram URL.
    Examples:
      - 'https://www.instagram.com/cristiano/?hl=en' -> 'cristiano'
      - 'https://instagram.com/cristiano/reels/'     -> 'cristiano'
      - '@cristiano'                                 -> 'cristiano'
      - 'cristiano'                                  -> 'cristiano'
    """
    cleaned = raw_input.strip()
    if not cleaned:
        return ""

    if "instagram.com" in cleaned.lower():
        if not cleaned.startswith(("http://", "https://")):
            cleaned = "https://" + cleaned
        parsed = urlparse(cleaned)
        path_parts = [p for p in parsed.path.split("/") if p]
        reserved = {"p", "reel", "reels", "tv", "stories", "explore", "accounts"}
        for part in path_parts:
            if part.lower() not in reserved:
                return part.lstrip("@").strip()

    cleaned = cleaned.lstrip("@").split("/")[0].split("?")[0].strip()
    return cleaned


def classify_media_type(typename: str, is_video: bool) -> str:
    """
    Maps Instaloader's GraphQL typename (`GraphVideo`, `GraphImage`, `GraphSidecar`)
    to a clear human-readable category.
    """
    if typename == "GraphVideo" or is_video:
        return "Reel / Video"
    if typename == "GraphSidecar":
        return "Carousel (Multi-Slide)"
    if typename == "GraphImage":
        return "Image"
    return typename or "Post"


def fetch_profile_posts_ist(
    username: str,
    login_user: Optional[str] = None,
    login_password: Optional[str] = None,
    session_file: Optional[str] = None,
    max_posts: Optional[int] = None,
    only_reels_videos: bool = False,
    progress_callback=None,
) -> Tuple[Dict, pd.DataFrame]:
    """
    Fetches posts from an Instagram profile using Instaloader and converts timestamps to IST (UTC+5:30).
    Returns (profile_metadata_dict, posts_dataframe).
    """
    try:
        import instaloader
    except ImportError as exc:
        raise RuntimeError(
            "Missing `instaloader`. Install it with: pip install instaloader"
        ) from exc

    loader = instaloader.Instaloader(
        download_pictures=False,
        download_videos=False,
        download_video_thumbnails=False,
        download_geotags=False,
        download_comments=False,
        save_metadata=False,
        compress_json=False,
        quiet=True,
    )

    # Authenticate if optional credentials or saved session are provided
    if login_user:
        logged_in = False
        try:
            if session_file:
                loader.load_session_from_file(login_user, filename=session_file)
            else:
                loader.load_session_from_file(login_user)
            logged_in = True
        except Exception:
            if login_password:
                try:
                    loader.login(login_user, login_password)
                    try:
                        loader.save_session_to_file()
                    except Exception:
                        pass
                    logged_in = True
                except Exception as login_exc:
                    raise RuntimeError(
                        f"Could not log in as '{login_user}'. Check your username/password or 2FA settings. "
                        f"Details: {login_exc}"
                    ) from login_exc
        if not logged_in:
            raise RuntimeError(
                f"No saved session found for '{login_user}'. Either enter your password in the sidebar "
                f"or run `instaloader --login={login_user}` once in your terminal."
            )

    profile = instaloader.Profile.from_username(loader.context, username)
    raw_media_count = safe_int(safe_attr(profile, "mediacount", None), default=None)

    profile_info: Dict[str, Any] = {
        "username": safe_attr(profile, "username", username),
        "full_name": safe_attr(profile, "full_name", ""),
        "total_media_count": raw_media_count,
        "followers": safe_int(safe_attr(profile, "followers", None), default=None),
        "followees": safe_int(safe_attr(profile, "followees", None), default=None),
        "is_private": bool(safe_attr(profile, "is_private", False)),
        "profile_url": f"https://www.instagram.com/{safe_attr(profile, 'username', username)}/",
        "partial_warning": None,
    }

    if profile_info["is_private"] and not bool(
        safe_attr(profile, "followed_by_viewer", False)
    ):
        raise PermissionError(
            f"@{profile_info['username']} is a private account. To inspect a private account, "
            "you must log in with an account that follows it."
        )

    records: List[Dict] = []
    posts_iterator = profile.get_posts()

    if raw_media_count is not None and max_posts is not None:
        target_total = min(raw_media_count, max_posts)
    elif raw_media_count is not None:
        target_total = raw_media_count
    elif max_posts is not None:
        target_total = max_posts
    else:
        target_total = 0

    try:
        for idx, post in enumerate(posts_iterator, start=1):
            is_video = bool(safe_attr(post, "is_video", False))
            typename = str(safe_attr(post, "typename", "Post"))
            media_category = classify_media_type(typename, is_video)

            if only_reels_videos and media_category != "Reel / Video":
                if max_posts and len(records) >= max_posts:
                    break
                continue

            # Instaloader provides `post.date_utc` as a naive UTC datetime
            utc_dt = post.date_utc.replace(tzinfo=timezone.utc)
            ist_dt = utc_dt.astimezone(IST_TIMEZONE)

            shortcode = str(safe_attr(post, "shortcode", ""))
            url_prefix = "reel" if is_video else "p"
            post_url = f"https://www.instagram.com/{url_prefix}/{shortcode}/"

            caption_raw = str(safe_attr(post, "caption", "") or "")
            caption_clean = caption_raw.replace("\n", " ").strip()
            if len(caption_clean) > 120:
                caption_clean = caption_clean[:117] + "..."

            likes_val = safe_int(safe_attr(post, "likes", 0), default=0)
            comments_val = safe_int(safe_attr(post, "comments", 0), default=0)
            views_val = (
                safe_int(safe_attr(post, "video_view_count", 0), default=0)
                if is_video
                else 0
            )

            records.append(
                {
                    "Post #": len(records) + 1,
                    "Media Type": media_category,
                    "IST Timestamp": ist_dt.strftime("%d %b %Y, %I:%M:%S %p IST"),
                    "IST Date": ist_dt.strftime("%Y-%m-%d"),
                    "IST Time (12-Hr)": ist_dt.strftime("%I:%M:%S %p"),
                    "IST Time (24-Hr)": ist_dt.strftime("%H:%M:%S"),
                    "Day of Week (IST)": ist_dt.strftime("%A"),
                    "Hour of Day (IST)": ist_dt.strftime("%I %p"),
                    "UTC Timestamp": utc_dt.strftime("%Y-%m-%d %H:%M:%S UTC"),
                    "Likes": max(likes_val or 0, 0),
                    "Comments": max(comments_val or 0, 0),
                    "Video Views": max(views_val or 0, 0),
                    "Shortcode": shortcode,
                    "Post URL": post_url,
                    "Caption Preview": caption_clean,
                }
            )

            if progress_callback:
                progress_callback(len(records), target_total, idx)

            if max_posts and len(records) >= max_posts:
                break
    except Exception as loop_exc:
        if records:
            profile_info["partial_warning"] = (
                f"Instagram paused pagination after extracting {len(records)} posts ({loop_exc}). "
                "Add a secondary Instagram login in the left sidebar to fetch the rest."
            )
        else:
            raise

    df = pd.DataFrame(records)
    return profile_info, df


def export_dataframe_to_excel_bytes(df: pd.DataFrame, profile_info: Dict) -> bytes:
    """
    Creates a formatted Excel (.xlsx) file in memory with styled headers,
    clickable post links, and summary metadata.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "IST Post Timestamps"

    total_count_val = profile_info.get("total_media_count")
    total_count_display = (
        f"{total_count_val:,}" if isinstance(total_count_val, int) else f"{len(df):,}"
    )

    # Title Banner
    ws.merge_cells("A1:O1")
    title_cell = ws["A1"]
    title_cell.value = (
        f"Instagram Post & Reel IST Timestamps — @{profile_info.get('username', 'unknown')} "
        f"(Profile Total: {total_count_display} Posts | Extracted: {len(df):,})"
    )
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(
        start_color="1E293B", end_color="1E293B", fill_type="solid"
    )
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    if df.empty:
        ws["A3"] = "No posts matched the selected criteria."
    else:
        headers = list(df.columns)
        header_fill = PatternFill(
            start_color="334155", end_color="334155", fill_type="solid"
        )
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        ws.row_dimensions[2].height = 24
        for col_idx, col_name in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        link_font = Font(name="Calibri", size=10, color="2563EB", underline="single")
        regular_font = Font(name="Calibri", size=10)
        zebra_fill = PatternFill(
            start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
        )

        for row_idx, row_data in enumerate(df.itertuples(index=False), start=3):
            ws.row_dimensions[row_idx].height = 20
            is_even = row_idx % 2 == 0
            for col_idx, value in enumerate(row_data, start=1):
                col_name = headers[col_idx - 1]
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = regular_font
                if is_even:
                    cell.fill = zebra_fill

                if (
                    col_name == "Post URL"
                    and isinstance(value, str)
                    and value.startswith("http")
                ):
                    cell.hyperlink = value
                    cell.font = link_font
                elif col_name in {"Post #", "Likes", "Comments", "Video Views"}:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name in {
                    "IST Date",
                    "IST Time (12-Hr)",
                    "IST Time (24-Hr)",
                    "Day of Week (IST)",
                    "Media Type",
                }:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Auto-fit columns with sensible caps
        for col_idx, col_name in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(col_name)),
                *(len(str(val)) for val in df[col_name].astype(str).head(200)),
            )
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

        ws.freeze_panes = "A3"

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def run_streamlit_dashboard() -> None:
    """Renders the interactive Streamlit web application."""
    import streamlit as st

    st.set_page_config(
        page_title="Instagram IST Timestamp Tracker",
        page_icon="⏱️",
        layout="wide",
    )

    st.title("⏱️ Instagram Post & Reel Timestamp Tracker (IST)")
    st.markdown(
        "Enter any public **Instagram username** or **profile link** to extract all posted Reels, "
        "Videos, Carousels, and Images with their exact **Indian Standard Time (`UTC+5:30`)** timestamps, "
        "and export the full dataset to **Excel (`.xlsx`)**."
    )

    with st.sidebar:
        st.markdown("### 🌓 Appearance")
        is_dark = st.toggle(
            "Dark Mode",
            value=True,
            help="Toggle between Dark Mode and Light Mode",
        )
        st.divider()

        st.header("⚙️ Extraction Settings")
        only_reels = st.checkbox("Fetch Reels / Videos Only", value=False)
        limit_toggle = st.checkbox("Fetch All Posts (No Limit)", value=True)
        max_posts = None
        if not limit_toggle:
            max_posts = st.number_input(
                "Max Posts to Fetch", min_value=1, max_value=5000, value=100, step=10
            )

        st.divider()
        st.subheader("🔐 Optional Instagram Login")
        st.caption(
            "Leave blank to test immediately. If Instagram pauses after ~12–24 posts on a 79+ post profile, "
            "enter a secondary/burner Instagram account here so it can fetch all posts."
        )
        login_user = (
            st.text_input("Instagram Username (Optional)", value="").strip() or None
        )
        login_password = (
            st.text_input(
                "Instagram Password (Optional — auto-saves session)",
                type="password",
                value="",
            ).strip()
            or None
        )
        session_file = (
            st.text_input("Custom Session File Path (Optional)", value="").strip()
            or None
        )

    if is_dark:
        st.markdown(
            """
            <style>
                /* Dark Mode Styling */
                .stApp {
                    background-color: #0B0F19;
                    color: #F3F4F6;
                }
                [data-testid="stSidebar"] {
                    background-color: #111827 !important;
                    border-right: 1px solid #1F2937 !important;
                }
                [data-testid="stSidebar"] * {
                    color: #E5E7EB;
                }
                [data-testid="stMetric"] {
                    background: linear-gradient(135deg, rgba(30, 41, 59, 0.8), rgba(15, 23, 42, 0.8));
                    border: 1px solid rgba(255, 255, 255, 0.1);
                    border-radius: 12px;
                    padding: 16px 20px;
                    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.35);
                }
                [data-testid="stMetricLabel"] {
                    color: #94A3B8 !important;
                    font-size: 0.88rem !important;
                    font-weight: 500 !important;
                }
                [data-testid="stMetricValue"] {
                    color: #38BDF8 !important;
                    font-weight: 700 !important;
                }
                .stTextInput input {
                    background-color: #1E293B !important;
                    color: #F8FAFC !important;
                    border: 1px solid #334155 !important;
                    border-radius: 8px !important;
                }
                .stTextInput input:focus {
                    border-color: #38BDF8 !important;
                    box-shadow: 0 0 0 1px #38BDF8 !important;
                }
                div[data-testid="stExpander"] {
                    background-color: #1E293B !important;
                    border: 1px solid #334155 !important;
                    border-radius: 10px !important;
                }
                hr {
                    border-color: #1F2937 !important;
                }
            </style>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <style>
                /* Light Mode Styling */
                .stApp {
                    background-color: #FFFFFF;
                    color: #1E293B;
                }
                [data-testid="stSidebar"] {
                    background-color: #F8FAFC !important;
                    border-right: 1px solid #E2E8F0 !important;
                }
                [data-testid="stSidebar"] * {
                    color: #1E293B;
                }
                [data-testid="stMetric"] {
                    background: #F8FAFC;
                    border: 1px solid #E2E8F0;
                    border-radius: 12px;
                    padding: 16px 20px;
                    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
                }
                [data-testid="stMetricLabel"] {
                    color: #64748B !important;
                    font-size: 0.88rem !important;
                    font-weight: 500 !important;
                }
                [data-testid="stMetricValue"] {
                    color: #0F172A !important;
                    font-weight: 700 !important;
                }
                .stTextInput input {
                    background-color: #FFFFFF !important;
                    color: #0F172A !important;
                    border: 1px solid #CBD5E1 !important;
                    border-radius: 8px !important;
                }
                div[data-testid="stExpander"] {
                    background-color: #FFFFFF !important;
                    border: 1px solid #E2E8F0 !important;
                    border-radius: 10px !important;
                }
                hr {
                    border-color: #E2E8F0 !important;
                }
            </style>
            """,
            unsafe_allow_html=True,
        )

    col_input, col_btn = st.columns([4, 1])
    with col_input:
        raw_target = st.text_input(
            "Instagram Username or Profile Link",
            placeholder="e.g., natgeo or https://www.instagram.com/natgeo/",
        )
    with col_btn:
        st.write("")
        st.write("")
        fetch_clicked = st.button(
            "🚀 Fetch Timestamps", type="primary", use_container_width=True
        )

    if fetch_clicked:
        username = extract_username(raw_target)
        if not username:
            st.error("Please enter a valid Instagram username or profile URL.")
            return

        progress_bar = st.progress(
            0, text=f"Connecting to Instagram for @{username}..."
        )
        status_text = st.empty()

        def update_progress(fetched_count: int, total_count: int, scanned_count: int):
            if total_count and total_count > 0:
                ratio = min(fetched_count / total_count, 1.0)
                progress_bar.progress(
                    ratio,
                    text=f"Fetched {fetched_count} matching posts (scanned {scanned_count} of ~{total_count} total posts)...",
                )
            else:
                status_text.info(
                    f"Fetched {fetched_count} posts (scanned {scanned_count} items)..."
                )

        try:
            profile_info, df = fetch_profile_posts_ist(
                username=username,
                login_user=login_user,
                login_password=login_password,
                session_file=session_file,
                max_posts=max_posts,
                only_reels_videos=only_reels,
                progress_callback=update_progress,
            )
            progress_bar.empty()
            status_text.empty()

            st.session_state["profile_info"] = profile_info
            st.session_state["posts_df"] = df
        except Exception as exc:
            progress_bar.empty()
            status_text.empty()
            st.error(f"Error fetching data for @{username}: {exc}")
            st.info(
                "💡 **Tip if you see a 401/429 rate-limit error:** Enter a secondary/burner Instagram username "
                "and password in the left sidebar under **Optional Instagram Login**, then click **Fetch Timestamps** again."
            )
            return

    if "posts_df" in st.session_state and "profile_info" in st.session_state:
        profile_info = st.session_state["profile_info"]
        df: pd.DataFrame = st.session_state["posts_df"]

        if profile_info.get("partial_warning"):
            st.warning(profile_info["partial_warning"])

        full_name_display = (
            profile_info.get("full_name") or profile_info.get("username") or "Profile"
        )
        st.subheader(
            f"📊 Profile Summary: @{profile_info.get('username', '')} ({full_name_display})"
        )

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        reels_count = (
            int((df["Media Type"] == "Reel / Video").sum()) if not df.empty else 0
        )
        top_day = (
            df["Day of Week (IST)"].mode().iloc[0]
            if (not df.empty and not df["Day of Week (IST)"].mode().empty)
            else "N/A"
        )
        top_hour = (
            df["Hour of Day (IST)"].mode().iloc[0]
            if (not df.empty and not df["Hour of Day (IST)"].mode().empty)
            else "N/A"
        )

        total_count_val = profile_info.get("total_media_count")
        total_posts_str = (
            f"{total_count_val:,}"
            if isinstance(total_count_val, int)
            else f"{len(df):,}"
        )

        kpi1.metric("Profile Total Posts", total_posts_str)
        kpi2.metric("Posts Extracted", f"{len(df):,}")
        kpi3.metric("Reels / Videos Extracted", f"{reels_count:,}")
        kpi4.metric("Top Posting Window (IST)", f"{top_day} @ {top_hour}")

        if df.empty:
            st.warning(
                "No posts found for this profile with the current filter settings."
            )
            return

        # Quick Dashboard Filter & Export Bar
        filter_col, excel_col, csv_col = st.columns([2, 1, 1])
        with filter_col:
            selected_types = st.multiselect(
                "Filter Table by Media Type",
                options=sorted(df["Media Type"].unique().tolist()),
                default=sorted(df["Media Type"].unique().tolist()),
            )

        filtered_df = df[df["Media Type"].isin(selected_types)].reset_index(drop=True)

        with excel_col:
            excel_bytes = export_dataframe_to_excel_bytes(filtered_df, profile_info)
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"{profile_info.get('username', 'instagram')}_ist_timestamps.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with csv_col:
            csv_bytes = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📄 Download CSV (.csv)",
                data=csv_bytes,
                file_name=f"{profile_info.get('username', 'instagram')}_ist_timestamps.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Post URL": st.column_config.LinkColumn(
                    "Post URL", display_text="Open on Instagram ↗"
                ),
            },
        )


def run_cli_mode() -> None:
    """Allows running the script directly from the terminal to generate an Excel file."""
    parser = argparse.ArgumentParser(
        description="Fetch Instagram post & Reel timestamps in IST (UTC+5:30) and export to Excel."
    )
    parser.add_argument("target", help="Instagram username, @handle, or profile URL")
    parser.add_argument(
        "--login", help="Optional Instagram username for authentication", default=None
    )
    parser.add_argument("--password", help="Optional Instagram password", default=None)
    parser.add_argument(
        "--session-file", help="Optional path to Instaloader session file", default=None
    )
    parser.add_argument(
        "--max-posts",
        type=int,
        help="Optional maximum number of posts to fetch",
        default=None,
    )
    parser.add_argument(
        "--reels-only", action="store_true", help="Only include Reels/Videos"
    )
    parser.add_argument("--output", help="Output Excel filename", default=None)
    args = parser.parse_args()

    username = extract_username(args.target)
    if not username:
        print("Error: Invalid Instagram username or URL.", file=sys.stderr)
        sys.exit(1)

    print(f"[*] Fetching posts for @{username} in IST (UTC+5:30)...")

    def cli_progress(fetched: int, total: int, scanned: int):
        total_str = str(total) if total else "?"
        print(
            f"    -> Extracted {fetched} posts (scanned {scanned}/{total_str})...",
            end="\r",
            flush=True,
        )

    profile_info, df = fetch_profile_posts_ist(
        username=username,
        login_user=args.login,
        login_password=args.password,
        session_file=args.session_file,
        max_posts=args.max_posts,
        only_reels_videos=args.reels_only,
        progress_callback=cli_progress,
    )
    print()

    if profile_info.get("partial_warning"):
        print(f"[!] Warning: {profile_info['partial_warning']}")

    out_path = (
        args.output or f"{profile_info.get('username', username)}_ist_timestamps.xlsx"
    )
    excel_bytes = export_dataframe_to_excel_bytes(df, profile_info)
    with open(out_path, "wb") as f:
        f.write(excel_bytes)

    total_count_val = profile_info.get("total_media_count")
    total_str = str(total_count_val) if total_count_val is not None else "unknown"
    print(
        f"[+] Done! Extracted {len(df)} posts (out of {total_str} profile total) "
        f"for @{profile_info.get('username', username)}."
    )
    print(f"[+] Saved formatted Excel report to: {out_path}")


if __name__ == "__main__":
    # Detect whether the script is being run via `streamlit run` or standard `python` CLI
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        if get_script_run_ctx() is not None:
            run_streamlit_dashboard()
        else:
            run_cli_mode()
    except ImportError:
        run_cli_mode()
