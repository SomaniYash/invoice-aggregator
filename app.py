import streamlit as st
import io

# Import your two programs as modules (once you share them)
# import program1
# import program2

st.set_page_config(page_title="File Processor", layout="wide")
st.title("🗂️ File Processor")

col1, col2 = st.columns(2)

# ── Program 1 ──────────────────────────────────────────
with col1:
    st.subheader("Program 1 – [Name Here]")
    file1 = st.file_uploader("Drag & drop or click to upload", key="prog1", type=["csv","xlsx","txt"])
    
    if file1 is not None:
        st.success(f"✅ Received: {file1.name}")
        
        if st.button("Run Program 1"):
            with st.spinner("Processing..."):
                # result = program1.run(file1)   ← will wire this up
                result = b"placeholder output"
            
            st.download_button(
                label="⬇️ Download Output",
                data=result,
                file_name="output_program1.csv",
                mime="text/csv"
            )

# ── Program 2 ──────────────────────────────────────────
with col2:
    st.subheader("Program 2 – [Name Here]")
    file2 = st.file_uploader("Drag & drop or click to upload", key="prog2", type=["csv","xlsx","txt"])
    
    if file2 is not None:
        st.success(f"✅ Received: {file2.name}")
        
        if st.button("Run Program 2"):
            with st.spinner("Processing..."):
                # result = program2.run(file2)   ← will wire this up
                result = b"placeholder output"
            
            st.download_button(
                label="⬇️ Download Output",
                data=result,
                file_name="output_program2.csv",
                mime="text/csv"
            )
