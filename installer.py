import os
import sys
import subprocess

def install_service():
    """Install the Windows service"""
    try:
        # Install the service
        subprocess.run([sys.executable, 'time_monitor_service.py', 'install'], check=True)
        print("Service installed successfully!")
        
        # Start the service
        subprocess.run([sys.executable, 'time_monitor_service.py', 'start'], check=True)
        print("Service started successfully!")
        
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")
        print("Make sure to run this script as Administrator")

if __name__ == '__main__':
    install_service()
