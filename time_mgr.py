import win32serviceutil
import win32service
import win32event
import servicemanager
import win32ts
import win32security
import win32api
import win32con
import yaml
import requests
import time
import os
import datetime
from pathlib import Path

class TimeMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'TimeMonitorService'
    _svc_display_name_ = 'Time Monitor Service'
    _svc_description_ = 'Monitors user session time and logs off non-admin users outside allowed hours'
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_alive = True
        self.config_file = os.path.join(os.path.dirname(__file__), 'time_config.yaml')
        
    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
        
    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        self.main()
        
    def load_time_config(self):
        """Load time configuration from remote URL, caching locally"""
        config_url = "http://r-pi/time-mgr.config.yaml"

        # Try to fetch from remote URL first
        try:
            response = requests.get(config_url, timeout=10)
            response.raise_for_status()  # Raise exception for bad status codes

            # Save the remote config to local file as cache
            with open(self.config_file, 'w') as f:
                f.write(response.text)

            # Parse and return the config
            return yaml.safe_load(response.text)

        except (requests.RequestException, yaml.YAMLError) as e:
            servicemanager.LogErrorMsg(f"Error loading remote config: {e}, falling back to local cache")

        # Fall back to local file if remote fetch failed
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    return yaml.safe_load(f)
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error loading local config: {e}")

        return None
        
    def is_user_admin(self, session_id):
        """Check if user in session is administrator"""
        try:
            # Get user token for the session
            user_token = win32ts.WTSQueryUserToken(session_id)
            
            # Check if user is in administrators group
            admin_sid = win32security.CreateWellKnownSid(win32security.WinBuiltinAdministratorsSid)
            is_admin = win32security.CheckTokenMembership(user_token, admin_sid)
            
            win32api.CloseHandle(user_token)
            return is_admin
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error checking admin status: {e}")
            return True  # Default to admin to avoid accidental logoffs
            
    def is_time_allowed(self, config):
        """Check if current time is within allowed hours"""
        if not config or 'days' not in config:
            return True
            
        now = datetime.datetime.now()
        day_name = now.strftime('%A').lower()
        
        if day_name not in config['days']:
            return False
            
        time_range = config['days'][day_name]
        if time_range == 'off' or not time_range:
            return False
            
        try:
            start_time, end_time = time_range.split('-')
            start_hour, start_min = map(int, start_time.split(':'))
            end_hour, end_min = map(int, end_time.split(':'))
            
            current_time = now.time()
            start = datetime.time(start_hour, start_min)
            end = datetime.time(end_hour, end_min)
            
            servicemanager.LogInfoMsg(f"Current time: {current_time}, interval: {start} - {end}")
            
            if start <= end:
                return start <= current_time <= end
            else:  # Crosses midnight
                return current_time >= start or current_time <= end
                
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error parsing time range: {e}")
            return True
            
    def logoff_user_session(self, session_id):
        """Log off a user session"""
        try:
            result = win32ts.WTSLogoffSession(win32ts.WTS_CURRENT_SERVER_HANDLE, session_id, False)
            if result:
                servicemanager.LogInfoMsg(f"Successfully logged off session {session_id}")
            else:
                servicemanager.LogErrorMsg(f"Failed to log off session {session_id}")
        except Exception as e:
            servicemanager.LogErrorMsg(f"Error logging off session {session_id}: {e}")
            
    def check_sessions(self):
        """Check all active sessions and log off non-admin users if outside time range"""
        config = self.load_time_config()
        
        if not self.is_time_allowed(config):
            try:
                sessions = win32ts.WTSEnumerateSessions(win32ts.WTS_CURRENT_SERVER_HANDLE, 0, 1)
                
                for session in sessions:
                    session_id = int(session['SessionId'])
                    state = session['State']
                    
                    # Only check active or disconnected sessions
                    if state in [win32ts.WTSActive, win32ts.WTSDisconnected]:
                        try:
                            # Get username for logging
                            username = win32ts.WTSQuerySessionInformation(
                                win32ts.WTS_CURRENT_SERVER_HANDLE,
                                session_id,
                                win32ts.WTSUserName
                            )
                            
                            if username and not self.is_user_admin(session_id):
                                servicemanager.LogInfoMsg(f"Logging off non-admin user: {username} (Session {session_id})")
                                self.logoff_user_session(session_id)
                                
                        except Exception as e:
                            servicemanager.LogErrorMsg(f"Error processing session {session_id}: {e}")
                            
            except Exception as e:
                servicemanager.LogErrorMsg(f"Error enumerating sessions: {e}")
        else:
            servicemanager.LogInfoMsg(f"Allowed time")
    def main(self):
        """Main service loop"""
        while self.is_alive:
            self.check_sessions()
            
            # Wait for 60 seconds or until service stop is requested
            if win32event.WaitForSingleObject(self.hWaitStop, 60000) == win32event.WAIT_OBJECT_0:
                break

if __name__ == '__main__':
    if len(os.sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TimeMonitorService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TimeMonitorService)
