import streamlit as st
from attend_selenium import get_work_codes, run_scraper, BASE_URL
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from bs4 import BeautifulSoup
import time

st.set_page_config(page_title="NREGA Attendance Scraper", layout="wide")

st.title("NREGA Attendance Scraper")

# Initialize session state
if 'stage' not in st.session_state:
    st.session_state.stage = 'initial'
if 'driver' not in st.session_state:
    st.session_state.driver = None

def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920x1080")
    options.binary_location = "/usr/bin/chromium"
    return webdriver.Chrome(service=ChromeService(), options=options)

@st.cache_data
def get_available_dates():
    driver = init_driver()
    driver.get(BASE_URL)
    time.sleep(2)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()
    attendance_select = soup.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddl_attendance'})
    if not attendance_select:
        return []
    return [opt['value'] for opt in attendance_select.find_all('option') if opt.has_attr('value') and opt['value']]


# Step 1: Get initial inputs
with st.container():
    st.markdown("### Step 1: Select Date and Panchayath")
    available_dates = get_available_dates()
    panchayath_list = [
        "BAGEWADI", "BAGGURU", "BALAKUNDHI", "BEERAHALLI", "B.M. SUGURU", 
        "BYRAPURA", "DESANOORU", "HACHCHOLI", "HALEKOTE", "H. HOSAHALLI", 
        "KARURU", "K. BELAGALLU", "KENCHANAGUDDA", "KONCHIGERI", "K. SUGURU", 
        "KUDUDHURAHAL", "KURUVALLI", "M. SUGURU", "MUDDATANURU", "NADAVI", 
        "RARAVI", "SANAVASAPURA", "SIRIGERI", "TALURU", "UPPARA HOSAHALLI", 
        "UTHTHANURU"
    ]

    if not available_dates:
        st.error("Could not fetch available dates. The website might be down or its structure may have changed.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.attendance_date = st.selectbox("Select Attendance Date", available_dates)
        with col2:
            st.session_state.panchayath_name = st.selectbox("Select Panchayath", panchayath_list)
        
        if st.button("Find Work Codes"):
            if not st.session_state.panchayath_name:
                st.warning("Please select a Panchayath.")
            else:
                st.session_state.stage = 'work_codes_loading'
                st.rerun()

# Step 2: Fetch and display work codes
if st.session_state.stage == 'work_codes_loading':
    with st.spinner("Fetching available work codes... This may take a moment."):
        try:
            st.session_state.driver = init_driver()
            work_codes, page_source, panchayath_url, workcode_idx, muster_no_idx = get_work_codes(
                st.session_state.driver, 
                st.session_state.attendance_date, 
                st.session_state.panchayath_name
            )
            st.session_state.work_codes = work_codes
            st.session_state.page_source = page_source
            st.session_state.panchayath_url = panchayath_url
            st.session_state.workcode_idx = workcode_idx
            st.session_state.muster_no_idx = muster_no_idx
            st.session_state.stage = 'work_codes_loaded'
            st.rerun()
        except Exception as e:
            st.error(f"Failed to get work codes: {e}")
            if st.session_state.driver:
                st.session_state.driver.quit()
            st.session_state.stage = 'initial'
            st.session_state.driver = None


# Step 3: Select work codes and run scraper
if st.session_state.stage == 'work_codes_loaded':
    with st.container():
        st.markdown("### Step 2: Select Work Codes and Scrape")
        if not st.session_state.work_codes:
            st.warning("No work codes found for the selected criteria.")
            if st.button("Go Back"):
                st.session_state.stage = 'initial'
                st.rerun()
        else:
            use_all = st.checkbox("Select all work codes", value=True)
            
            # Use a column to give the multiselect more space
            col1, col2 = st.columns([2, 1])
            with col1:
                if use_all:
                    selected_codes = ['all']
                    st.multiselect(
                        "Available Work Codes", 
                        st.session_state.work_codes, 
                        default=st.session_state.work_codes, 
                        disabled=True,
                        help="All available work codes are selected."
                    )
                else:
                    selected_codes = st.multiselect(
                        "Available Work Codes", 
                        st.session_state.work_codes,
                        help="Select one or more work codes to scrape."
                    )
            
            if st.button("Start Scraping"):
                if not selected_codes:
                    st.warning("Please select at least one work code.")
                else:
                    st.session_state.selected_codes = selected_codes
                    st.session_state.stage = 'scraping'
                    st.rerun()

# Step 4: Run scraper and show results
if st.session_state.stage == 'scraping':
    st.markdown("### Step 3: Download Results")
    progress_bar = st.progress(0)
    status_text = st.empty()

    def status_callback(message, percentage):
        status_text.text(message)
        progress_bar.progress(int(percentage))

    try:
        muster_rolls_excel, muster_images_excel, raw_data_excel = run_scraper(
            st.session_state.driver,
            st.session_state.page_source,
            st.session_state.panchayath_url,
            st.session_state.panchayath_name,
            st.session_state.attendance_date,
            st.session_state.workcode_idx,
            st.session_state.muster_no_idx,
            st.session_state.selected_codes,
            status_callback
        )
        
        st.session_state.muster_rolls_excel = muster_rolls_excel
        st.session_state.muster_images_excel = muster_images_excel
        st.session_state.raw_data_excel = raw_data_excel
        st.session_state.stage = 'results_ready'
        st.rerun()

    except Exception as e:
        st.error(f"An error occurred during scraping: {e}")
        if st.session_state.driver:
            st.session_state.driver.quit()
        st.session_state.stage = 'initial'
        st.session_state.driver = None

if st.session_state.stage == 'results_ready':
    st.success("Scraping complete! You can now download the files.")

    col1, col2, col3 = st.columns(3)
    dl_date = st.session_state.attendance_date.replace('/', '_')
    # Sanitize panchayath name for filename
    dl_panch = st.session_state.panchayath_name.replace('.', '').replace(' ', '_')

    with col1:
        st.download_button("Muster Rolls Excel", st.session_state.muster_rolls_excel, f"muster_rolls_{dl_panch}_{dl_date}.xlsx")
    with col2:
        st.download_button("Muster Images Excel", st.session_state.muster_images_excel, f"muster_images_{dl_panch}_{dl_date}.xlsx")
    with col3:
        st.download_button("Raw Data Excel", st.session_state.raw_data_excel, f"raw_data_{dl_panch}_{dl_date}.xlsx")

    if st.button("Start New Scrape"):
        # Clean up session state for next run
        for key in ['driver', 'work_codes', 'page_source', 'panchayath_url', 'workcode_idx', 'muster_no_idx', 'selected_codes', 'muster_rolls_excel', 'muster_images_excel', 'raw_data_excel']:
            if key in st.session_state:
                del st.session_state[key]
        st.session_state.stage = 'initial'
        st.rerun()

# Footer
st.markdown("---")
# Replace 'Your Name' with the actual name you want to display
st.markdown("<div style='text-align: center; padding: 10px;'>Developed by: Ajay Shankar A</div>", unsafe_allow_html=True)
