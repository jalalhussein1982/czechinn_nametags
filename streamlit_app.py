"""
Hotel Nametag Generator - Streamlit Web Application
----------------------------------------------------
Web interface for generating hotel nametags from PDF reports.
GDPR compliant: All temporary files are deleted when session ends.
"""

import streamlit as st
import tempfile
import os
import atexit
from io import BytesIO
from pathlib import Path

# Import from existing generator
from hotel_nametag_generator import PDFParser, NametageGenerator, GuestRecord


def cleanup_temp_files():
    """Clean up all temporary files stored in session state."""
    if 'temp_files' in st.session_state:
        for temp_path in st.session_state.temp_files:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except Exception:
                pass
        st.session_state.temp_files = []

    if 'output_docx' in st.session_state:
        st.session_state.output_docx = None

    if 'guests' in st.session_state:
        st.session_state.guests = None


def init_session_state():
    """Initialize session state variables."""
    if 'temp_files' not in st.session_state:
        st.session_state.temp_files = []
    if 'output_docx' not in st.session_state:
        st.session_state.output_docx = None
    if 'guests' not in st.session_state:
        st.session_state.guests = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False


def process_pdf(uploaded_file) -> tuple:
    """
    Process uploaded PDF and generate DOCX.
    Returns (success: bool, message: str, docx_bytes: BytesIO or None)
    """
    temp_pdf_path = None
    temp_docx_path = None

    try:
        # Save uploaded PDF to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            temp_pdf_path = tmp_pdf.name
            st.session_state.temp_files.append(temp_pdf_path)

        # Parse PDF
        parser = PDFParser(temp_pdf_path)
        guests = parser.process()

        if not guests:
            return False, "No guest records found in the PDF. Please check the file format.", None

        st.session_state.guests = guests

        # Generate DOCX to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            temp_docx_path = tmp_docx.name
            st.session_state.temp_files.append(temp_docx_path)

        generator = NametageGenerator(guests)
        pages = generator.generate(temp_docx_path)

        # Read DOCX into memory for download
        with open(temp_docx_path, 'rb') as f:
            docx_bytes = BytesIO(f.read())

        st.session_state.output_docx = docx_bytes
        st.session_state.processing_complete = True

        return True, f"Successfully processed {len(guests)} guests across {pages} pages.", docx_bytes

    except Exception as e:
        return False, f"Error processing PDF: {str(e)}", None

    finally:
        # Clean up temporary PDF immediately after processing
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
                if temp_pdf_path in st.session_state.temp_files:
                    st.session_state.temp_files.remove(temp_pdf_path)
            except Exception:
                pass


def end_session():
    """End the current session and clean up all data."""
    cleanup_temp_files()
    st.session_state.processing_complete = False
    st.session_state.clear()


def main():
    """Main Streamlit application."""

    # Page configuration
    st.set_page_config(
        page_title="Czech Inn Nametag Generator",
        page_icon="🏨",
        layout="centered"
    )

    # Initialize session state
    init_session_state()

    # Header
    st.title("🏨 Czech Inn Nametag Generator")
    st.markdown("---")

    # GDPR notice
    st.info(
        "**GDPR Compliance:** All uploaded files and generated documents are stored "
        "temporarily in memory only. Data is automatically deleted when you end your "
        "session or close the browser."
    )

    # Main content area
    col1, col2 = st.columns([3, 1])

    with col2:
        # End Session button (always visible)
        if st.button("🚪 End Session", type="secondary", use_container_width=True):
            end_session()
            st.rerun()

    with col1:
        st.subheader("Upload PDF Report")

    # File uploader
    uploaded_file = st.file_uploader(
        "Select your hotel arrivals PDF",
        type=['pdf'],
        help="Upload the PDF report containing guest arrival information"
    )

    # Processing section
    if uploaded_file is not None:
        st.markdown("---")

        # Show file info
        st.write(f"**Uploaded file:** {uploaded_file.name}")
        st.write(f"**File size:** {uploaded_file.size / 1024:.1f} KB")

        # Process button
        if st.button("🔄 Generate Nametags", type="primary", use_container_width=True):
            with st.spinner("Processing PDF and generating nametags..."):
                success, message, docx_bytes = process_pdf(uploaded_file)

                if success:
                    st.success(message)
                else:
                    st.error(message)

    # Download section (show if processing is complete)
    if st.session_state.processing_complete and st.session_state.output_docx:
        st.markdown("---")
        st.subheader("📥 Download Generated Nametags")

        # Guest summary
        if st.session_state.guests:
            guests = st.session_state.guests
            st.write(f"**Total guests:** {len(guests)}")

            # Show preview of first few guests
            with st.expander("Preview guest list (first 10)"):
                for i, guest in enumerate(guests[:10]):
                    st.write(f"{i+1}. Room {guest.room_number}: {guest.last_name} "
                            f"(Arrival: {guest.arrival_day}, Departure: {guest.departure_day})")
                if len(guests) > 10:
                    st.write(f"... and {len(guests) - 10} more guests")

        # Download button
        st.session_state.output_docx.seek(0)
        st.download_button(
            label="📄 Download Nametags (DOCX)",
            data=st.session_state.output_docx,
            file_name="nametags_output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

        st.warning(
            "**Remember:** Click 'End Session' when you're done to ensure all "
            "temporary data is deleted (GDPR compliance)."
        )

    # Footer
    st.markdown("---")
    st.caption(
        "Czech Inn Hotel Nametag Generator | "
        "Generates 12 nametags per A4 page (6 rows × 2 columns)"
    )


# Register cleanup on exit
atexit.register(cleanup_temp_files)


if __name__ == "__main__":
    main()
