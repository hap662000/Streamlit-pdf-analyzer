import streamlit as st
import time

# Page configuration
st.set_page_config(
    page_title="AutoTake Login",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# st.set_page_config(
#     page_title="Email PDF Analyzer",
#     page_icon=":email:",
#     layout="wide",
#     initial_sidebar_state="collapsed"
# )

# Custom CSS for dark theme
st.markdown("""
<style>
    .main {
        background-color: #121212;
    }
    .login-container {
        max-width: 350px;
        margin: 0 auto;
        padding: 2rem;
        border-radius: 10px;
        background-color: #1e1e1e;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }
    .input-container {
        max-width: 280px;
        margin: 0 auto;
    }
    .stButton > button {
        width: 280px;
        margin: 0 auto;
        display: block;
        background-color: #4361ee;
        color: white;
        border: none;
        padding: 0.5rem 0;
        font-weight: 500;
        border-radius: 5px;
        transition: all 0.3s;
    }
    .stButton > button:hover {
        background-color: #3a56d4;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.4);
    }
    .title {
        color: #ffffff;
        font-weight: 700;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .subtitle {
        color: #b0b0b0;
        font-size: 0.9rem;
        text-align: center;
        margin-bottom: 2rem;
    }
    .login-footer {
        text-align: center;
        font-size: 0.8rem;
        color: #777;
        margin-top: 2rem;
    }
    /* Fix label colors for dark theme */
    label {
        color: #e0e0e0 !important;
    }
    /* Remove streamlit elements margins that might cause empty space */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0 !important;
        max-width: 100% !important;
    }
    /* Dark input fields */
    .stTextInput > div > div > input {
        background-color: #2d2d2d !important;
        color: #e0e0e0 !important;
        border: 1px solid #444 !important;
    }
    /* Center form elements */
    .stTextInput {
        max-width: 280px;
        margin: 0 auto 1rem auto;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for authentication
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

# Login Page
def login_page():
    with st.container():
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        
        # Title and subtitle
        st.markdown('<h1 class="title">Welcome to AutoTake</h1>', unsafe_allow_html=True)
        st.markdown('<p class="subtitle">Sign in to access your dashboard</p>', unsafe_allow_html=True)
        
        # Login form
        username = st.text_input("Username", placeholder="Enter your username")
        password = st.text_input("Password", type="password", placeholder="Enter your password")
        
        # Login button
        st.markdown("<div style='height: 1rem'></div>", unsafe_allow_html=True)
        if st.button("Sign In"):
            if username == "admin" and password == "admin@autotake":
                # Success animation
                with st.spinner("Signing in..."):
                    time.sleep(1)
                st.success("Login successful!")
                
                # Redirect simulation (without balloons)
                st.markdown("Redirecting to dashboard...")
                progress = st.progress(0)
                for i in range(100):
                    progress.progress(i + 1)
                    time.sleep(0.01)
                
                # Set authentication state to True
                st.session_state.authenticated = True
                st.rerun()  # Rerun the app to show the dashboard
            else:
                st.error("Invalid credentials.")
        
        # Footer
        st.markdown('<p class="login-footer">© 2025 AutoTake. All rights reserved.</p>', unsafe_allow_html=True)
        
        # Close the container
        st.markdown('</div>', unsafe_allow_html=True)

# Main App Logic
if not st.session_state.authenticated:
    login_page()
else:
    # Import the main app logic here
    from app import main  # Assuming your main app logic is in a file named `app.py`
    main()