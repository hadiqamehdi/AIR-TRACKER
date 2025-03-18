AIR TRACKER
AIR TRACKER is an advanced application designed to provide a gesture-based control system for PowerPoint presentations and additional interactive features. Utilizing computer vision and hand gesture recognition, it delivers a touch-free and seamless presentation experience.

Key Features
Gesture-Controlled Navigation: Move through slides, zoom in/out, and control video playback using hand gestures.
Virtual Drawing Canvas: Highlight, draw, and erase on a virtual overlay for better presentation engagement.
Live Camera Feed: Display a real-time camera feed during presentations.
PowerPoint Automation: Open, manage, and close PowerPoint presentations directly from the application.
Customizable Tools: Easily switch between different drawing tools, such as a pen, highlighter, and eraser.
Installation Guide
Clone the repository:
bash
Copy
Edit
git clone https://github.com/username/AIR-TRACKER.git  
cd AIR-TRACKER  
Install dependencies:
bash
Copy
Edit
pip install -r requirements.txt  
Ensure a functional webcam for gesture detection.
How to Use
Launch the application:
bash
Copy
Edit
python main.py  
Upload a PowerPoint file via the GUI and start presenting.
Use these gestures for navigation:
Index and Pinky Fingers Up → Toggle virtual canvas.
Thumb Up → Move to the previous slide.
Pinky Up → Advance to the next slide.
Four Fingers Up → Zoom out.
Three Middle Fingers Up → Zoom in.
Thumb and Pinky Up → Exit the application.
Utilize the virtual canvas to draw, highlight, or erase using the available tools.
Project Structure
graphql
Copy
Edit
AIR-TRACKER/  
├── main.py               # Main graphical interface  
├── utils.py              # Functions for PowerPoint automation  
├── new.py                # Gesture detection and camera management  
├── canvas_handler.py     # Manages virtual drawing canvas  
├── camera.py             # Handles real-time camera feed and gesture recognition  
├── requirements.txt      # List of project dependencies  
├── media/                # Contains icons and other media assets  
└── README.md             # Project documentation  
System Requirements
Python 3.8 or later
Windows OS (for PowerPoint integration)
A functional webcam for gesture tracking
Required Dependencies
The application utilizes the following Python libraries:

customtkinter – Modern GUI framework
Pillow – Image processing
opencv-python – Computer vision processing
cvzone – Utility for image and video processing
mediapipe – Hand tracking module
numpy – Numerical operations
keras – Deep learning framework
pyautogui – Automating UI interactions
pynput – Handling keyboard and mouse input
comtypes, pywin32 – PowerPoint automation
Check requirements.txt for precise versions.

Common Issues & Fixes
Ensure that PowerPoint files are valid and accessible.
The application is currently optimized for Windows and may not function correctly on other operating systems.
