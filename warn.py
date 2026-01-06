import win32ts
import win32con

def send_message_to_all_sessions(title, message, timeout=30):
    """Send message to all active sessions"""
    try:
        # Get all sessions
        sessions = win32ts.WTSEnumerateSessions()
        
        for session in sessions:
            session_id = int(session['SessionId'])
            state = session['State']
            
            # Only send to active sessions
            if state == win32ts.WTSActive:
                try:
                    # Get username for logging
                    username = win32ts.WTSQuerySessionInformation(
                        win32ts.WTS_CURRENT_SERVER_HANDLE,
                        session_id,
                        win32ts.WTSUserName
                    )
                    
                    if username and username.strip():
                        print(f"Sending message to user: {username} (Session {session_id})")
                        
                        # Send message with OK button and timeout
                        response = win32ts.WTSSendMessage(
                            win32ts.WTS_CURRENT_SERVER_HANDLE,  # Server handle
                            session_id,                          # Session ID
                            title,                              # Message title
                            message,                            # Message text
                            win32con.MB_OK | win32con.MB_ICONINFORMATION,  # Style
                            timeout,                            # Timeout in seconds
                            True                                # Wait for response
                        )
                        
                        print(f"Message sent to {username}, response: {response}")
                        
                except Exception as e:
                    print(f"Error sending message to session {session_id}: {e}")
                    
    except Exception as e:
        print(f"Error enumerating sessions: {e}")

# Example usage
send_message_to_all_sessions(
    "System Notification",
    "The server will restart in 15 minutes. Please save your work.",
    60
)
