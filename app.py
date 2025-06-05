import streamlit as st
import pandas as pd
import time
import io
from email_verifier import EmailVerifier
from excel_processor import ExcelProcessor

def main():
    st.set_page_config(
        page_title="Bulk Email Verification Tool",
        page_icon="📧",
        layout="wide"
    )
    
    st.title("📧 Bulk Email Verification Tool")
    st.markdown("Upload an Excel file to verify email addresses using SMTP validation")
    
    # Initialize session state
    if 'verification_results' not in st.session_state:
        st.session_state.verification_results = None
    if 'original_df' not in st.session_state:
        st.session_state.original_df = None
    if 'email_column' not in st.session_state:
        st.session_state.email_column = None
    
    # File upload section
    st.header("1. Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx, .xls)",
        type=['xlsx', 'xls'],
        help="Upload an Excel file containing email addresses"
    )
    
    if uploaded_file is not None:
        try:
            # Process the uploaded file
            processor = ExcelProcessor()
            df = processor.read_excel(uploaded_file)
            
            st.success(f"✅ File uploaded successfully! Found {len(df)} rows")
            
            # Display file preview
            st.subheader("File Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Email column detection/selection
            st.header("2. Select Email Column")
            
            # Auto-detect email columns
            email_columns = processor.detect_email_columns(df)
            
            if email_columns:
                st.info(f"🔍 Auto-detected potential email columns: {', '.join(email_columns)}")
                default_column = email_columns[0]
            else:
                st.warning("⚠️ No email columns auto-detected. Please select manually.")
                default_column = df.columns[0] if len(df.columns) > 0 else None
            
            # Column selection
            email_column = st.selectbox(
                "Select the column containing email addresses:",
                options=df.columns.tolist(),
                index=df.columns.tolist().index(default_column) if default_column in df.columns else 0
            )
            
            # Show sample emails from selected column
            if email_column:
                st.session_state.email_column = email_column
                st.session_state.original_df = df
                
                sample_emails = df[email_column].dropna().head(5).tolist()
                st.write("Sample emails from selected column:")
                for email in sample_emails:
                    st.write(f"• {email}")
                
                # Email verification section
                st.header("3. Email Verification")
                
                # Get valid emails for verification
                valid_emails = processor.extract_valid_emails(df, email_column)
                total_emails = len(valid_emails)
                
                if total_emails > 0:
                    st.info(f"📊 Found {total_emails} valid email addresses to verify")
                    
                    # Verification settings
                    col1, col2 = st.columns(2)
                    with col1:
                        delay_between_checks = st.slider(
                            "Delay between checks (seconds)",
                            min_value=0.5,
                            max_value=5.0,
                            value=1.0,
                            step=0.5,
                            help="Delay to avoid overwhelming SMTP servers"
                        )
                    
                    with col2:
                        timeout_seconds = st.slider(
                            "SMTP timeout (seconds)",
                            min_value=5,
                            max_value=30,
                            value=10,
                            help="Timeout for SMTP connections"
                        )
                    
                    # Start verification button
                    if st.button("🚀 Start Email Verification", type="primary"):
                        verify_emails(df, email_column, delay_between_checks, timeout_seconds)
                    
                    # Show results if available
                    if st.session_state.verification_results is not None:
                        show_verification_results()
                
                else:
                    st.error("❌ No valid email addresses found in the selected column")
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")


def verify_emails(df, email_column, delay, timeout):
    """Perform email verification with progress tracking"""
    processor = ExcelProcessor()
    verifier = EmailVerifier(timeout=timeout)
    
    # Get emails to verify
    valid_emails = processor.extract_valid_emails(df, email_column)
    total_emails = len(valid_emails)
    
    if total_emails == 0:
        st.error("No valid emails to verify")
        return
    
    # Create progress indicators
    progress_bar = st.progress(0)
    status_text = st.empty()
    results_placeholder = st.empty()
    
    # Initialize results tracking
    verified_count = 0
    valid_count = 0
    invalid_count = 0
    error_count = 0
    
    # Create a copy of the dataframe for results
    result_df = df.copy()
    result_df['Email_Verification_Status'] = 'Not Checked'
    result_df['Verification_Details'] = ''
    
    try:
        for index, row in df.iterrows():
            email = row[email_column]
            
            if pd.isna(email) or not processor.is_valid_email_format(str(email)):
                result_df.at[index, 'Email_Verification_Status'] = 'Invalid Format'
                result_df.at[index, 'Verification_Details'] = 'Invalid email format'
                continue
            
            # Update status
            verified_count += 1
            status_text.text(f"Verifying {verified_count}/{total_emails}: {email}")
            
            # Verify email
            is_valid, details = verifier.verify_email(str(email))
            
            if is_valid:
                result_df.at[index, 'Email_Verification_Status'] = 'Valid'
                valid_count += 1
            elif details == 'error':
                result_df.at[index, 'Email_Verification_Status'] = 'Error'
                error_count += 1
            else:
                result_df.at[index, 'Email_Verification_Status'] = 'Invalid'
                invalid_count += 1
            
            result_df.at[index, 'Verification_Details'] = details
            
            # Update progress
            progress = verified_count / total_emails
            progress_bar.progress(progress)
            
            # Show intermediate results
            with results_placeholder.container():
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("✅ Valid", valid_count)
                col2.metric("❌ Invalid", invalid_count)
                col3.metric("⚠️ Errors", error_count)
                col4.metric("📊 Progress", f"{verified_count}/{total_emails}")
            
            # Rate limiting
            if verified_count < total_emails:
                time.sleep(delay)
        
        # Store results
        st.session_state.verification_results = result_df
        
        # Final status
        status_text.text("✅ Verification completed!")
        progress_bar.progress(1.0)
        
        st.success(f"🎉 Email verification completed! Valid: {valid_count}, Invalid: {invalid_count}, Errors: {error_count}")
        
    except Exception as e:
        st.error(f"❌ Error during verification: {str(e)}")


def show_verification_results():
    """Display verification results and download option"""
    st.header("4. Verification Results")
    
    if st.session_state.verification_results is not None:
        result_df = st.session_state.verification_results
        
        # Summary statistics
        col1, col2, col3, col4 = st.columns(4)
        
        valid_count = len(result_df[result_df['Email_Verification_Status'] == 'Valid'])
        invalid_count = len(result_df[result_df['Email_Verification_Status'] == 'Invalid'])
        error_count = len(result_df[result_df['Email_Verification_Status'] == 'Error'])
        total_count = len(result_df)
        
        col1.metric("Total Emails", total_count)
        col2.metric("✅ Valid", valid_count)
        col3.metric("❌ Invalid", invalid_count)
        col4.metric("⚠️ Errors", error_count)
        
        # Results table
        st.subheader("Detailed Results")
        st.dataframe(result_df, use_container_width=True)
        
        # Download section
        st.subheader("Download Results")
        
        # Prepare download
        processor = ExcelProcessor()
        excel_buffer = processor.dataframe_to_excel(result_df)
        
        st.download_button(
            label="📥 Download Verified Excel File",
            data=excel_buffer,
            file_name=f"email_verification_results_{int(time.time())}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        # Option to start over
        if st.button("🔄 Start Over"):
            st.session_state.verification_results = None
            st.session_state.original_df = None
            st.session_state.email_column = None
            st.rerun()


if __name__ == "__main__":
    main()
