import os
import threading
import ctypes
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Frame, BOTTOM, X
from PIL import Image, ImageTk
from utils import PowerPointHandler  # Import the PowerPointHandler class

class AirTrackerApp:
    """Main application for Air Tracker."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AIR TRACKER")
        self.root.geometry("1350x700")
        self.root.configure(bg="#a91f2d")
        self.powerpoint_handler = PowerPointHandler()
        self.powerpoint_handler.initialize_listener()  # Initialize the mouse listener
        self._initialize_ui()

    def _initialize_ui(self):
        """Initializes the user interface."""
        # Header
        header_label = tk.Label(
            self.root,
            text="AIR TRACKER",
            font=("Georgia", 36, "italic", "bold"),
            bg="#a91f2d",
            fg="white"
        )
        header_label.pack(pady=10)

        # Image Display
        self._display_image()

        # Footer
        footer_frame = Frame(self.root, bg="white", height=50, width=1400)
        footer_frame.pack(side=BOTTOM, fill=X)

        footer_label = tk.Label(
            footer_frame,
            text="Wave goodbye to clicks and embrace gestures.",
            font=("Georgia", 16, "italic"),
            bg="white",
            fg="#a91f2d"
        )
        footer_label.pack(pady=10)

        # Upload Button
        self._create_upload_button()

        # Set the icon for the app
        self._set_app_icon()

    def _display_image(self):
        """Displays the main application image."""
        try:
            # Use a raw string to avoid Unicode escape issues
            image_path = r"C:\Users\mfabid\OneDrive\Desktop\Ai-Track3\hand.png"
            original_image = Image.open(image_path)
            resized_image = original_image.resize((300, 300))
            my_image = ImageTk.PhotoImage(resized_image)

            image_label = Label(self.root, image=my_image, bg="#a91f2d")
            image_label.image = my_image  # Keep a reference to avoid garbage collection
            image_label.pack(pady=20)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load image: {e}")

    def _set_app_icon(self):
        """Sets the application icon."""
        # Use forward slashes to avoid issues with backslashes
        icon_path = "C:/Users/mfabid/OneDrive/Desktop/Ai-Track3/hand.ico"
        if os.name == 'nt':  # Check if running on Windows
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"MyAppID")
            self.root.iconbitmap(icon_path)

    def _create_upload_button(self):
        """Creates the upload button."""
        def on_hover(event):
            upload_button.configure(fg_color="white", text_color="#a91f2d", hover_color="#ff7373")

        def on_leave(event):
            upload_button.configure(fg_color="#a91f2d", text_color="white", hover_color="#ff7373")

        upload_button = ctk.CTkButton(
            self.root,
            text="Upload PowerPoint File",
            font=("Arial", 20, "bold"),
            fg_color="#a91f2d",
            text_color="white",
            hover_color="#ff7373",
            corner_radius=20,
            command=self.run_presentation
        )

        upload_button.bind("<Enter>", on_hover)
        upload_button.bind("<Leave>", on_leave)
        upload_button.pack(pady=20)

    def run_presentation(self):
        """Handles the upload and execution of a PowerPoint presentation."""
        try:
            # Kill any existing PowerPoint processes first
            os.system('taskkill /F /IM POWERPNT.EXE 2>nul')
            
            ppt_file = filedialog.askopenfilename(
                title="Select PowerPoint file",
                filetypes=[("PowerPoint Files", "*.pptx *.ppt")]
            )

            if ppt_file:
                if os.path.exists(ppt_file):
                    self.root.withdraw()
                    # Start PowerPoint in a separate thread
                    threading.Thread(
                        target=self.powerpoint_handler.run_powerpoint,
                        args=(ppt_file, self.root),
                        daemon=True
                    ).start()
                else:
                    messagebox.showerror("File Not Found", f"The file could not be found: {ppt_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            # Try to restore the main window if there's an error
            try:
                self.root.deiconify()
            except:
                pass

    def run(self):
        """Runs the main application loop."""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            # Handle clean exit on Ctrl+C
            self.powerpoint_handler.cleanup()
            self.root.quit()
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            self.powerpoint_handler.cleanup()
            self.root.quit()

if __name__ == "__main__":
    app = AirTrackerApp()
    app.run()