import os
import comtypes.client
import time
import threading
import pythoncom
from pynput.mouse import Listener 
from tkinter import messagebox, Toplevel, Label
import win32gui
import win32con
import subprocess
from camera import CameraHandler
import win32com.client

class PowerPointHandler:
    """Class to manage PowerPoint interactions and overlays."""

    def __init__(self):
        self.overlay_window = None
        self.powerpoint_lock = threading.Lock()
        self.activation_checked = False

    def check_powerpoint_activation(self):
        """Check if PowerPoint is properly activated before proceeding."""
        try:
            # Try to create a PowerPoint instance
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = False
            
            # Try to access Presentations collection
            _ = powerpoint.Presentations.Count
            
            # If we get here, PowerPoint is activated
            powerpoint.Quit()
            return True
        except Exception as e:
            if "not activated" in str(e).lower() or "rejected by callee" in str(e).lower():
                messagebox.showerror(
                    "PowerPoint Activation Required",
                    "Microsoft PowerPoint is not properly activated. Please activate PowerPoint before using this application.\n\n"
                    "Steps to resolve:\n"
                    "1. Close this application\n"
                    "2. Open PowerPoint normally\n"
                    "3. Complete the activation process\n"
                    "4. Restart this application"
                )
                return False
            return True  # Other errors might not be activation-related
        finally:
            try:
                powerpoint.Quit()
            except:
                pass

    @staticmethod
    def safe_com_call(func, max_retries=5, delay=1, *args, **kwargs):
        """Safely call a COM method with retry logic to handle busy application."""
        for attempt in range(max_retries):
            try:
                return func(*args, **kwargs)
            except comtypes.COMError as e:
                if "rejected by callee" in str(e).lower():
                    raise Exception("PowerPoint is not properly activated. Please activate it first.")
                if attempt < max_retries - 1:
                    print(f"COM call failed (attempt {attempt + 1}), retrying...")
                    time.sleep(delay)
                else:
                    print(f"COM call failed after {max_retries} attempts.")
                    raise e

    @staticmethod
    def initialize_listener():
        """Starts a mouse listener to focus PowerPoint on mouse clicks."""
        def on_click(x, y, button, pressed):
            if pressed:
                PowerPointHandler.focus_powerpoint_window()

        listener = Listener(on_click=on_click)
        listener.start()

    @staticmethod
    def focus_powerpoint_window():
        """Focuses on the PowerPoint window when it's running."""
        try:
            hwnd = win32gui.FindWindow(None, None)
            while hwnd:
                window_title = win32gui.GetWindowText(hwnd)
                if "PowerPoint Slide Show" in window_title:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    break
                hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
        except Exception as e:
            print(f"Error focusing PowerPoint window: {e}")

    def close_all_powerpoint_windows(self):
        """Close all PowerPoint-related windows including activation dialogs."""
        def enum_windows_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if any(title in window_title for title in [
                    "PowerPoint", 
                    "Microsoft Office Activation",
                    "Microsoft Office Professional Plus"
                ]):
                    try:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        time.sleep(0.1)  # Give it a moment to close
                    except:
                        pass
        win32gui.EnumWindows(enum_windows_callback, None)

    def force_kill_powerpoint(self):
        """Forcefully kills all running PowerPoint processes."""
        try:
            print("Forcefully killing all PowerPoint processes...")
            # First try to close windows gracefully
            self.close_all_powerpoint_windows()
            time.sleep(0.5)  # Wait a bit for windows to close
            
            # Then force kill any remaining processes
            subprocess.run(
                ["taskkill", "/F", "/IM", "POWERPNT.EXE"],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            time.sleep(1)  # Wait for processes to fully terminate
            print("All PowerPoint processes terminated.")
        except subprocess.CalledProcessError as e:
            if e.returncode != 128:  # 128 means no processes found, which is fine
                print(f"Error forcefully killing PowerPoint processes: {e}")

    def display_camera_overlay(self, root):
        """Displays the camera overlay window with a live camera feed."""
        self.overlay_window = Toplevel(root)
        self.overlay_window.title("Camera Feed Overlay")
        self.overlay_window.attributes('-topmost', True)
        self.overlay_window.geometry("300x200+1000+0")

        camera_label = Label(self.overlay_window)
        camera_label.pack()

        # Pass the root argument to CameraHandler
        camera_handler = CameraHandler(camera_label, root)  
        camera_handler.start_camera()

        print("Camera overlay displayed.")
        print("Camera overlay window geometry:", self.overlay_window.geometry())
        print("Camera overlay window state:", self.overlay_window.state())

    def run_powerpoint(self, ppt_file, root):
        """Runs the PowerPoint presentation with proper thread and COM handling."""
        if not self.activation_checked:
            if not self.check_powerpoint_activation():
                root.deiconify()  # Show the main window again
                return
            self.activation_checked = True

        powerpoint = None
        ppt = None

        # Kill all existing PowerPoint processes before starting
        self.force_kill_powerpoint()

        with self.powerpoint_lock:
            retries = 3  # Number of retries
            for attempt in range(retries):
                try:
                    pythoncom.CoInitialize()

                    # Create PowerPoint Application
                    print("Initializing PowerPoint application...")
                    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                    powerpoint.Visible = 1

                    # Open the PowerPoint file
                    print(f"Opening PowerPoint file: {ppt_file}")
                    formatted_ppt_file = os.path.abspath(ppt_file)
                    ppt = powerpoint.Presentations.Open(formatted_ppt_file)

                    if not ppt:
                        raise Exception("Failed to open PowerPoint file.")

                    # Start the slideshow in windowed mode
                    print("Starting slideshow in windowed mode...")
                    slide_show_settings = ppt.SlideShowSettings
                    slide_show_settings.ShowType = 2  # ppShowTypeWindow
                    slide_show_settings.Run()

                    # Focus PowerPoint window
                    time.sleep(1)
                    self.focus_powerpoint_window()

                    print("Displaying camera overlay...")
                    self.display_camera_overlay(root)

                    # Monitor the slideshow
                    while powerpoint.SlideShowWindows.Count > 0:
                        time.sleep(0.5)

                    break  # Exit the retry loop if successful

                except Exception as e:
                    print(f"Attempt {attempt + 1} failed: {e}")
                    if attempt == retries - 1:
                        messagebox.showerror("Error", str(e))
                        root.deiconify()  # Show the main window again
                    time.sleep(2)  # Wait before retrying

                finally:
                    try:
                        if ppt:
                            ppt.Close()
                        if powerpoint:
                            powerpoint.Quit()
                    except:
                        pass
                    
                    if self.overlay_window:
                        self.overlay_window.destroy()
                    
                    pythoncom.CoUninitialize()
                    self.force_kill_powerpoint()  # Ensure everything is cleaned up
    def cleanup(self):
        """Clean up resources when closing the application."""
        self.force_kill_powerpoint()
        if self.overlay_window:
            self.overlay_window.destroy()