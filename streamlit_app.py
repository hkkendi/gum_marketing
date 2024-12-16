import streamlit as st
import socket
import ipaddress
from streamlit.web.server.websocket_headers import _get_websocket_headers
from streamlit.runtime.scriptrunner import get_script_run_ctx

def get_client_ip():
    """Get client's IP address."""
    try:
        ctx = get_script_run_ctx()
        if ctx is None:
            return None
        
        # Get headers from the context
        headers = _get_websocket_headers()
        
        # Try to get IP from X-Forwarded-For header first
        ip = headers.get("X-Forwarded-For", "").split(",")[0].strip()
        if not ip:
            # Fallback to other headers
            ip = headers.get("X-Real-IP", "")
        if not ip:
            # Final fallback
            ip = headers.get("Remote-IP", "")
            
        return ip
    except Exception:
        return None

def is_allowed_ip(ip):
    """Check if IP is in allowed range (192.168.xxx)."""
    try:
        if ip is None:
            return False
        
        ip_obj = ipaddress.ip_address(ip)
        
        # Check if IP starts with 192.168
        return str(ip_obj).startswith("192.168.")
    except ValueError:
        return False

# IP checking middleware
def check_ip_access():
    client_ip = get_client_ip()
    
    if not is_allowed_ip(client_ip):
        st.error("Access Denied: Your IP address is not authorized to access this application.")
        st.write(f"Your IP: {client_ip}")
        st.stop()

# Main app code
def main():
    # Check IP access first
    check_ip_access()
    
    # Your original app code goes here
    st.title("IP-Restricted Streamlit App")
    st.write("Welcome! You're accessing from an authorized IP address.")
    
    # Add your app's functionality here
    
if __name__ == "__main__":
    main()