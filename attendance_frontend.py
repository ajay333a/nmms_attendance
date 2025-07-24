import streamlit as st
from datetime import date
from attendance_downloader import run_attendance_downloader

st.title('Attendance Downloader')

# Session state for reset and file tracking
if 'submitted' not in st.session_state:
    st.session_state['submitted'] = False
if 'files' not in st.session_state:
    st.session_state['files'] = None
if 'progress_msgs' not in st.session_state:
    st.session_state['progress_msgs'] = []

# User input fields
panchayat_name = st.text_input('Panchayath Name (e.g., BALAKUNDHI)', key='panchayat_name')
panchayat_code = st.text_input('Panchayath Code (last 3 digits, e.g., 016)', key='panchayat_code')
fin_year = st.text_input('Financial Year (e.g., 2024-2025)', value='2024-2025', key='fin_year')
work_code = st.text_input('Work Code', key='work_code')
msr_start = st.number_input('Muster Roll Start Number', min_value=1, step=1, key='msr_start')
msr_end = st.number_input('Muster Roll End Number', min_value=1, step=1, key='msr_end')
attendance_date = st.date_input('Attendance Date', value=date.today(), key='attendance_date')
digest = st.text_input('Digest', key='digest')

# Progress area
progress_area = st.empty()

# Progress callback for backend
progress_msgs = st.session_state.get('progress_msgs', [])
def progress_callback(msg):
    progress_msgs.append(msg)
    st.session_state['progress_msgs'] = progress_msgs
    progress_area.write('\n'.join(progress_msgs))

# Download button and status
download_btn_col, status_col = st.columns([2, 1])
with download_btn_col:
    download_clicked = st.button('Download Attendance Data')

if download_clicked:
    # Input validation
    errors = []
    if not panchayat_name:
        errors.append('Panchayath Name is required.')
    if not panchayat_code:
        errors.append('Panchayath Code is required.')
    if not fin_year:
        errors.append('Financial Year is required.')
    if not work_code:
        errors.append('Work Code is required.')
    if not digest:
        errors.append('Digest is required.')
    if msr_start > msr_end:
        errors.append('Muster Roll Start Number must be less than or equal to End Number.')
    if errors:
        for err in errors:
            st.error(err)
    else:
        st.session_state['submitted'] = True
        st.session_state['progress_msgs'] = []
        att_date_str = attendance_date.strftime('%d/%m/%Y')
        if not panchayat_code.startswith('1505007'):
            panchayat_code_full = '1505007' + panchayat_code
        else:
            panchayat_code_full = panchayat_code
        st.info('Running backend process...')
        try:
            files = run_attendance_downloader(
                panchayat_name, panchayat_code_full, fin_year, work_code, int(msr_start), int(msr_end), att_date_str, digest,
                progress_callback=progress_callback
            )
            st.session_state['files'] = files
        except Exception as e:
            st.error(f'Error during processing: {e}')

# Show progress messages
if st.session_state.get('progress_msgs'):
    progress_area.write('\n'.join(st.session_state['progress_msgs']))

# Show download buttons if files exist
files = st.session_state.get('files')
if files:
    att_date_str = attendance_date.strftime('%d/%m/%Y')
    file_base = f"{work_code}_{att_date_str}".replace('/', '_')
    file_labels = [
        (files[0], f'attendance_data_{file_base}.xlsx', 'Attendance Data Excel'),
        (files[1], f'attendance_images_{file_base}.xlsx', 'Images Excel'),
        (files[2], f'attendance_with_images_{file_base}.xlsx', 'Attendance+Images Excel'),
        # (files[3], f'attendance_with_images_{file_base}.pdf', 'PDF Report'),  # PDF download button commented out
    ]
    for file_obj, fname, label in file_labels:
        if file_obj is not None:
            st.download_button(f'Download {label}', file_obj, file_name=fname, key=fname)
    with status_col:
        st.success('✔️ Parsing complete! Files are ready for download.')
    # Reset button below download buttons
    if st.button('Reset App'):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun() 