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

    if 'nametag_counts' in st.session_state:
        st.session_state.nametag_counts = {}

    if 'pending_review' in st.session_state:
        st.session_state.pending_review = False

    if 'multi_guest_rooms' in st.session_state:
        st.session_state.multi_guest_rooms = {}


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
    if 'pending_review' not in st.session_state:
        st.session_state.pending_review = False
    if 'nametag_counts' not in st.session_state:
        st.session_state.nametag_counts = {}  # {room_key: count}
    if 'multi_guest_rooms' not in st.session_state:
        st.session_state.multi_guest_rooms = {}  # {room_key: {info about room}}


def parse_pdf(uploaded_file) -> tuple:
    """
    Parse uploaded PDF and extract guest records.
    Returns (success: bool, message: str, guests: list or None)
    """
    temp_pdf_path = None

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

        # Group guests by room to find rooms with multiple guests
        # Key: (room_number, last_name, arrival_day, departure_day) to uniquely identify a room booking
        room_groups = {}
        for guest in guests:
            if guest.number_of_guests > 1:
                room_key = f"{guest.room_number}|{guest.last_name}|{guest.arrival_day}|{guest.departure_day}"
                if room_key not in room_groups:
                    room_groups[room_key] = {
                        'room_number': guest.room_number,
                        'last_name': guest.last_name,
                        'arrival_day': guest.arrival_day,
                        'departure_day': guest.departure_day,
                        'number_of_guests': guest.number_of_guests,
                        'guest_records': []
                    }
                room_groups[room_key]['guest_records'].append(guest)

        if room_groups:
            st.session_state.multi_guest_rooms = room_groups
            # Initialize nametag counts to number_of_guests (default = max)
            for room_key, room_info in room_groups.items():
                st.session_state.nametag_counts[room_key] = room_info['number_of_guests']
            st.session_state.pending_review = True
            return True, f"Found {len(guests)} guest records. {len(room_groups)} room(s) have multiple guests and need review.", guests
        else:
            # No multi-guest rooms, proceed directly
            st.session_state.pending_review = False
            return True, f"Found {len(guests)} guest records. No multi-guest rooms to review.", guests

    except Exception as e:
        return False, f"Error parsing PDF: {str(e)}", None

    finally:
        # Clean up temporary PDF immediately after processing
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            try:
                os.remove(temp_pdf_path)
                if temp_pdf_path in st.session_state.temp_files:
                    st.session_state.temp_files.remove(temp_pdf_path)
            except Exception:
                pass


def generate_docx(guests: list) -> tuple:
    """
    Generate DOCX from guest records, selecting based on nametag_counts per room.
    Returns (success: bool, message: str, docx_bytes: BytesIO or None)
    """
    temp_docx_path = None

    try:
        # Build the final guest list based on user selections
        final_guests = []
        processed_rooms = set()

        for guest in guests:
            if guest.number_of_guests == 1:
                # Single guest, always include
                final_guests.append(guest)
            else:
                # Multi-guest room - use the room key to get user's selection
                room_key = f"{guest.room_number}|{guest.last_name}|{guest.arrival_day}|{guest.departure_day}"

                if room_key not in processed_rooms:
                    # Get the count user selected for this room
                    count = st.session_state.nametag_counts.get(room_key, guest.number_of_guests)
                    # Get all guest records for this room
                    room_guests = st.session_state.multi_guest_rooms.get(room_key, {}).get('guest_records', [guest])
                    # Take only the selected number of nametags
                    final_guests.extend(room_guests[:count])
                    processed_rooms.add(room_key)

        # Generate DOCX to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
            temp_docx_path = tmp_docx.name
            st.session_state.temp_files.append(temp_docx_path)

        generator = NametageGenerator(final_guests)
        pages = generator.generate(temp_docx_path)

        # Read DOCX into memory for download
        with open(temp_docx_path, 'rb') as f:
            docx_bytes = BytesIO(f.read())

        st.session_state.output_docx = docx_bytes
        st.session_state.processing_complete = True
        st.session_state.pending_review = False

        return True, f"Successfully generated {len(final_guests)} nametags across {pages} pages.", docx_bytes

    except Exception as e:
        return False, f"Error generating DOCX: {str(e)}", None


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

        # Step 1: Parse PDF button (only show if not already parsed)
        if not st.session_state.guests and not st.session_state.pending_review:
            if st.button("📋 Parse PDF", type="primary", use_container_width=True):
                with st.spinner("Parsing PDF..."):
                    success, message, guests = parse_pdf(uploaded_file)

                    if success:
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)

    # Step 2: Multi-guest review section
    if st.session_state.pending_review and st.session_state.multi_guest_rooms:
        st.markdown("---")
        st.subheader("🔍 Review Multi-Guest Rooms")
        st.info(
            "The following rooms have multiple guests. For each room, "
            "select how many nametags you want to print (e.g., private rooms may only need 1)."
        )

        for room_key, room_info in st.session_state.multi_guest_rooms.items():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(
                    f"**Room {room_info['room_number']}** - {room_info['last_name']} | "
                    f"Guests: {room_info['number_of_guests']} | "
                    f"Arrival: {room_info['arrival_day']} → Departure: {room_info['departure_day']}"
                )
            with col2:
                count = st.number_input(
                    f"Nametags",
                    min_value=1,
                    max_value=room_info['number_of_guests'],
                    value=st.session_state.nametag_counts.get(room_key, room_info['number_of_guests']),
                    key=f"count_{room_key}",
                    label_visibility="collapsed"
                )
                st.session_state.nametag_counts[room_key] = count

        st.markdown("---")

        # Summary before generation
        total_single = len([g for g in st.session_state.guests if g.number_of_guests == 1])
        total_multi = sum(st.session_state.nametag_counts.values())
        total_nametags = total_single + total_multi

        st.write(f"**Total nametags to generate:** {total_nametags}")
        st.caption(f"({total_single} from single-guest rooms + {total_multi} from multi-guest rooms)")

        if st.button("🔄 Generate Nametags", type="primary", use_container_width=True):
            with st.spinner("Generating nametags..."):
                success, message, docx_bytes = generate_docx(st.session_state.guests)

                if success:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)

    # If no multi-guest reservations, generate directly after parsing
    if st.session_state.guests and not st.session_state.pending_review and not st.session_state.processing_complete:
        with st.spinner("Generating nametags..."):
            success, message, docx_bytes = generate_docx(st.session_state.guests)

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
            # Calculate actual nametag count
            total_single = len([g for g in guests if g.number_of_guests == 1])
            total_multi = sum(st.session_state.nametag_counts.values())
            total_nametags = total_single + total_multi
            num_rooms = len(st.session_state.multi_guest_rooms) + total_single

            st.write(f"**Total rooms:** {num_rooms}")
            st.write(f"**Total nametags generated:** {total_nametags}")

            # Show preview of rooms
            with st.expander("Preview room list"):
                # Show single-guest rooms
                single_guests = [g for g in guests if g.number_of_guests == 1]
                for i, guest in enumerate(single_guests[:5]):
                    st.write(f"Room {guest.room_number}: {guest.last_name} "
                            f"(Arrival: {guest.arrival_day}, Departure: {guest.departure_day}) "
                            f"- 1 nametag")

                # Show multi-guest rooms
                for room_key, room_info in list(st.session_state.multi_guest_rooms.items())[:5]:
                    count = st.session_state.nametag_counts.get(room_key, room_info['number_of_guests'])
                    st.write(f"Room {room_info['room_number']}: {room_info['last_name']} "
                            f"(Arrival: {room_info['arrival_day']}, Departure: {room_info['departure_day']}) "
                            f"- {count}/{room_info['number_of_guests']} nametag(s)")

                if num_rooms > 10:
                    st.write(f"... and {num_rooms - 10} more rooms")

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
