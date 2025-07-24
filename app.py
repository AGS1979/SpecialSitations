import streamlit as st
import base64

# --- Must be the first st.* command ---
st.set_page_config(page_title="Pre-IPO Memo Generator", layout="wide")

# --- Function to encode local image ---
def get_base64_logo(path: str):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# --- Use a known-good remote image for testing ---
# A simple 200x200px placeholder image
# If this remote image works and your logo doesn't, the issue is the logo file itself.
# REMOTE_LOGO_URL = "https://via.placeholder.com/200x50.png?text=Test+Logo"

# --- Use your local logo file ---
LOCAL_LOGO_PATH = "logo.png"
try:
    logo_base64 = get_base64_logo(LOCAL_LOGO_PATH)
    LOGO_URL = f"data:image/png;base64,{logo_base64}"
except FileNotFoundError:
    st.error("Your logo.png file was not found. Please make sure it's in the same directory.")
    LOGO_URL = "" # Fallback to empty if logo not found

# --- Clean, Isolated CSS and HTML for the Header ---
st.markdown(f"""
    <style>
        /* This targets the main container of the Streamlit app */
        .appview-container .main .block-container {{
            padding-top: 2rem; /* Push content down */
        }}
        
        /* This targets the Streamlit Toolbar (Deploy button, hamburger menu) */
        [data-testid="stToolbar"] {{
            right: 2rem;
        }}
    </style>
    """,
    unsafe_allow_html=True
)

# Use st.image for robust, native image rendering
if LOGO_URL:
    st.image(LOGO_URL, width=180)


# --- Your App's Content Starts Here ---
st.title("Pre-IPO Investment Memo Generator")
st.write("Upload an IPO/DRHP PDF to generate a structured investment memo with optional Q&A.")
st.markdown("---")

# --- UI FOR DEMONSTRATION ---
st.header("Upload PDF and Focus")
st.file_uploader("Upload DRHP or IPO PDF", type=['pdf'])
st.text_area("Optional: Add custom notes to guide memo generation")
st.button("Generate Memo")