import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import scrolledtext
import cv2
import os
import time
import sounddevice as sd
import winreg
import sys
import subprocess
from PIL import Image, ImageTk
import tkinter.filedialog as filedialog
from win10toast import ToastNotifier
import pystray
from pystray import MenuItem as item
import platform
import psutil
import hashlib
from tkinter import simpledialog
import win32clipboard
import win32con
import getpass
import pyaudio
import wave
import ctypes
import webbrowser
import audioop

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
if not is_admin():
    try:
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 0
        )
    except:
        pass
    sys.exit(0)
save_folder = "./"
video_bitrate = 1500000
recording_paused = False
segment_time = 600
last_segment_time = 0
show_fps_info = True
flip_mode = 0
watermark_enabled = True
current_cam_index = 0
tray_icon = None
is_minimized_to_tray = False
video_writer = None
is_recording = False
recording_start_time = 0
audio_stream = None
audio_wave_file = None
temp_video_path = ""
temp_audio_path = ""
final_video_path = ""
p = None
audio_only_recording = False
brightness_value = 50
contrast_value = 50
def init_required_dirs():
    dir_list = ["image", "video", "audio", "ErrorLog"]
    for dir_name in dir_list:
        if not os.path.exists(dir_name):
            try:
                os.makedirs(dir_name)
                print(f"Auto created directory: {dir_name}")
            except:
                pass
init_required_dirs()
root = tk.Tk()
root.geometry("1000x1000+400+200")
root.title("CapVisionPro")
root.eval('tk::PlaceWindow . center')
root.resizable(False, False)
save_prefix = tk.StringVar(value="CapVision_")
capture_format = tk.StringVar(value="bmp")
Option1 = tk.BooleanVar()
Option2 = tk.BooleanVar(value=True)
OptionAudio = tk.BooleanVar(value=False)
OptionCap = tk.BooleanVar(value=False)
mode_var = tk.IntVar(value=1)

def close_current_subwindow(event=None):
    focused = root.focus_get()
    if focused and focused != root:
        focused.destroy()

try:
    root.iconbitmap("Capicon.ico")
except:
    pass

root.bind("<Alt-b>", lambda event: statusbar())
root.bind("<Alt-n>", lambda event: set_password())
root.bind("<Alt-u>", lambda event: verify_on_start())
root.bind("<Alt-p>", lambda event: exitwindows())
root.bind("<Alt-k>", lambda event: AboutKeyboard())
root.bind("<Alt-a>", lambda event: show_about())
root.bind("<Alt-l>", lambda event: LicenseAgreement())
root.bind("<Alt-h>", lambda event: Show_help())
root.bind("<Alt-s>", lambda event: show_system_info())
root.bind("<Alt-i>", lambda event: OpenImageFolder())
root.bind("<Alt-v>", lambda event: OpenVideoFolder())
root.bind("<Alt-g>", lambda event: OpenLogFolder())
root.bind("<Alt-c>", lambda event: CaptureImage())
root.bind("<Alt-r>", lambda event: start_recording_countdown())
root.bind("<Alt-o>", lambda event: StopRecording())
root.bind("<Alt-d>", lambda event: DeleteTools())
root.bind("<Alt-t>", lambda event: Settings())
root.bind("<Alt-m>", lambda event: minimize_to_tray())
root.bind("<Alt-f>", lambda event: ToggleFullscreen())
root.bind("<Alt-w>", close_current_subwindow)
root.bind("<Escape>", close_current_subwindow)
root.bind("<Alt-Shift-S>", lambda e: open_last_screenshot())
root.bind("<Alt-Shift-D>", lambda e: open_last_recording())
root.bind("<Alt-Shift-Z>", lambda e: open_last_audio())
root.bind("<Alt-Shift-I>", lambda event: convert_image_format())
root.bind("<Alt-Shift-V>", lambda event: convert_video_format())
root.bind("<Alt-F>", lambda e: toggle_flip())
root.bind("<Alt-W>", lambda e: toggle_watermark())
root.bind("<Alt-Shift-T>", lambda e: capture_with_countdown())
root.bind("<Alt-P>", lambda e: toggle_pause_recording())
root.bind("<Alt-C>", lambda e: switch_camera())
root.bind("<Alt-R>", lambda e: rotate_selected_image())
root.bind("<Alt-X>", lambda e: set_save_prefix())
root.bind("<Alt-J>", lambda e: capture_as_jpg())
root.bind("<Alt-S>", lambda e: start_audio_only_recording())
root.bind("<Alt-Q>", lambda e: stop_audio_only_recording())
root.bind("<Alt-Shift-A>", lambda e: convert_audio_format())
root.bind("<Alt-Shift-R>", lambda e: run_script_file())
root.bind("<Alt-Shift-C>", lambda e: CaptureAndCopy())
root.bind("<Alt-L>", lambda e: view_log_file())
root.bind("<Alt-Shift-F>", lambda e: toggle_fps_display())
root.bind("<Alt-E>", lambda e: open_adjustment_window())
root.bind("<Alt-Shift-E>", lambda e: set_capture_format())
root.bind("<Alt-Shift-L>", lambda e: set_segment_time())
root.bind("<Alt-Shift-U>", lambda e: Updatelog())
root.bind("<Alt-Shift-O>", lambda e: OpenCode())
root.bind("<Alt-Shift-K>", lambda e: about_script())
root.bind("<Alt-Shift-Y>", lambda e: OpenAudioFolder())
root.bind("<Control-Shift-C>", lambda e: ColourMode())
root.bind("<Control-Shift-G>", lambda e: Graymode())
root.bind("<Control-Shift-I>", lambda e: Get_Image_Information())
root.bind("<Alt-Shift-@>",lambda e:Credits())
root.bind("<Control-Shift-V>", lambda e: ColourVideomode())
root.bind("<Control-Shift-B>", lambda e: GrayVideomode())
def AboutKeyboard():
    Keyboard_shortcuts="""
Alt-p - Exit Program
Alt-k - Shortcuts Help
Alt-a - About CapVisionPro
Alt-l - License Agreement
Alt-h - View Help Document
Alt-s - System Information
Alt-c - Capture Image
Alt-i - Open Image Folder
Alt-v - Open Video Folder
Alt-g - Open Log Folder
Alt-r - Start Recording (with countdown)
Alt-o - Stop Recording
Alt-d - Cleanup Tools
Alt-t - Application Settings
Alt-m - Minimize to Tray
Alt-f - Toggle Fullscreen
Alt-w - Close Subwindow
Alt-b - Show/Hide Status Bar
Alt-n - Set Password
Alt-u - Unlock Software
Alt-L - View Log File
Alt+Shift+S - Open Last Screenshot
Alt+Shift+D - Open Last Recording
Alt+Shift+Z - Open Last Audio
ESC - Close Current Subwindow
Alt+Shift+I - Convert Image Format
Alt+Shift+V - Convert Video Format
Alt+Shift+A - Convert Audio Format
Alt-F - Mirror / Flip
Alt-W - Toggle Watermark
Alt-Shift-T - Countdown Capture
Alt-P - Pause/Resume Recording
Alt-C - Switch Camera
Alt-R - Rotate Image
Alt-X - Set File Prefix
Alt-J - Capture High Quality JPG
Alt+S - Start Audio Only Recording
Alt+Q - Stop Audio Only Recording
Alt+Shift+R - Run Script File
Alt+Shift+C - Capture & Copy to Clipboard
Alt+Shift+F - Toggle FPS Display
Alt-E - Image Adjustment
Alt+Shift+E - Capture Format Settings
Alt+Shift+L - Auto Split Recording
Alt+Shift+U - Update Log
Alt+Shift+O - Open Source Code
Alt+Shift+K - About CVScript
Alt+Shift-Y - Open Audio Folder
Ctrl+Shift+C - Open Image in Colour Mode
Ctrl+Shift+G - Open Image in Gray Mode
Ctrl+Shift+I - Get Image Information
Alt+Shift+@ - Credits
Ctrl+Shift+V - Open Video in Colour Mode
Ctrl+Shift+B - Open Video in Gray Mode
"""
    messagebox.showinfo("Keyboard Shortcut",Keyboard_shortcuts)
def exitwindows():
    global cap, tray_icon
    if tray_icon:
        try:
            tray_icon.stop()
        except:
            pass
    if 'cap' in globals() and cap is not None:
        cap.release()
    clean_temp_files()
    root.destroy()
    sys.exit()

menu_bar = tk.Menu(root)
root.config(menu=menu_bar)
file_menu = tk.Menu(menu_bar, tearoff=0)

def OpenLogFolder():
    file_FolderLog="ErrorLog"
    if not os.path.exists(file_FolderLog):
        messagebox.showinfo("Directory Status","Error log directory not found.\nAutomatically creating directory in the application root path.")
        os.makedirs(file_FolderLog)
    else:
        try:
            subprocess.Popen(['explorer',file_FolderLog])
        except Exception:
            messagebox.showerror("Error","Failed to open the log directory.")

def OpenImageFolder():
    file_folderimage="image"
    if not os.path.exists(file_folderimage):
        messagebox.showinfo("Directory Status","Image directory not found.\nAutomatically creating directory in the application root path.")
        os.makedirs(file_folderimage)
    else:
        try:
            subprocess.Popen(['explorer',file_folderimage])
        except Exception:
            messagebox.showerror("Error","Failed to open the image directory.")

def OpenVideoFolder():
    file_folderVideo="video"
    if not os.path.exists(file_folderVideo):
        messagebox.showinfo("Directory Status","Video directory not found.\nAutomatically creating directory in the application root path.")
        os.makedirs(file_folderVideo)
    else:
        try:
            subprocess.Popen(['explorer',file_folderVideo])
        except Exception:
            messagebox.showerror("Error","Failed to open the video directory.")

def OpenAudioFolder():
    file_folderAudio="audio"
    if not os.path.exists(file_folderAudio):
        messagebox.showinfo("Directory Status","Audio directory not found.\nAutomatically creating directory in the application root path.")
        os.makedirs(file_folderAudio)
    else:
        try:
            subprocess.Popen(['explorer',file_folderAudio])
        except Exception:
            messagebox.showerror("Error","Failed to open the audio directory.")

def minimize_to_tray():
    global tray_icon, is_minimized_to_tray
    if is_minimized_to_tray:
        return
    root.withdraw()
    toaster = ToastNotifier()
    toaster.show_toast(title="CapVisionPro",msg="Application minimized to system tray.",icon_path="Capicon.ico",duration=3,threaded=True)
    try:
        icon_image = Image.open("Capicon.ico")
    except:
        icon_image = Image.new('RGB', (64, 64), color='blue')
    def restore_window(icon, item):
        global is_minimized_to_tray
        root.deiconify()
        icon.stop()
        is_minimized_to_tray = False
    def quit_app(icon, item):
        icon.stop()
        exitwindows()
    menu = (
        item('Restore', restore_window),
        item('Quit', quit_app),
    )
    tray_icon = pystray.Icon("CapVisionPro", icon_image, "CapVisionPro", menu)
    is_minimized_to_tray = True
    tray_icon.run()
def setup_tray_hotkey():
    def on_minimize(event):
        if root.state() == 'iconic':
            root.after(100, minimize_to_tray)
    root.bind("<Unmap>", on_minimize)

def show_system_info():
    info_win = tk.Toplevel(root)
    info_win.title("System Information")
    info_win.geometry("580x400")
    info_win.resizable(False, False)
    try:
        info_win.iconbitmap("Capicon.ico")
    except:
        pass
    info_win.bind("<Escape>", lambda e: info_win.destroy())
    title_label = tk.Label(
        info_win,
        text="CapVisionPro • System Information",
        font=("Segoe UI", 13, "bold")
    )
    title_label.pack(pady=10)
    frame = tk.Frame(info_win)
    frame.pack(padx=20, pady=5, fill="both", expand=True)
    sys_info = f"""
    Operating System     : {platform.system()} {platform.release()} ({platform.version()})
    Architecture         : {platform.machine()}
    Processor            : {platform.processor()}
    CPU Cores            : {psutil.cpu_count(logical=True)}
    Total RAM            : {round(psutil.virtual_memory().total / (1024**3), 2)} GB
    Available RAM        : {round(psutil.virtual_memory().available / (1024**3), 2)} GB
    Python Version       : {sys.version.split()[0]}
    Camera Resolution    : {int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))} × {int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))}
    Camera FPS           : {int(cap.get(cv2.CAP_PROP_FPS))}
    OpenCV Version       : {cv2.__version__}
    Application          : CapVisionPro
    Build                : Stable 1.99.01 Insider
    Window               : Main active
    """
    info_text = tk.Label(
        frame,
        text=sys_info,
        font=("Consolas", 10),
        justify="left",
        anchor="w"
    )
    info_text.pack(anchor="w")
    close_btn = tk.Button(
        info_win,
        text="Close",
        command=info_win.destroy,
        width=12
    )
    close_btn.pack(pady=10)
def show_about():
    win = tk.Toplevel(root)
    win.title("About CapVisionPro")
    win.geometry("400x350")
    win.resizable(False, False)
    try:
        win.iconbitmap("Capicon.ico")
    except:
        pass
    win.bind("<Escape>", lambda e: win.destroy())
    title_label = tk.Label(win, text="CapVisionPro", font=("Arial",12,"bold"), fg="#4682B4")
    title_label.pack(pady=20)

    notice_label = tk.Label(
        win,
        text="Modified versions must be open-sourced.",
        font=("Segoe UI",11),
        bg="#FFFF00",
        fg="#E02424",
        padx=10,
        pady=5
    )
    notice_label.pack(pady=10)

    version_label = tk.Label(win, text="Version 1.99.01 Insider", font=("Segoe UI",11))
    version_label.pack(pady=5)

    desc_label1 = tk.Label(
        win,
        text="A powerful camera capture and recording tool\nBuilt with Python, OpenCV and Tkinter",
        font=("Segoe UI",11),
        justify="center"
    )
    desc_label1.pack(pady=10)

    copyright_label = tk.Label(
        win,
        text="Copyright © 2025 Meter Corporation\nAll rights reserved.",
        font=("Segoe UI",11),
        justify="center"
    )
    copyright_label.pack(pady=10)

    ok_btn = tk.Button(win, text="OK", command=win.destroy, font=("Segoe UI",11), width=10)
    ok_btn.pack(pady=10)
def open_last_screenshot():
    image_dir = "image"
    if not os.path.exists(image_dir):
        messagebox.showwarning("Warning", "Screenshot folder does not exist!")
        return
    image_files = []
    for f in os.listdir(image_dir):
        if f.lower().endswith((".png", ".jpg", ".bmp")):
            full_path = os.path.join(image_dir, f)
            image_files.append((os.path.getmtime(full_path), full_path))
    if not image_files:
        messagebox.showinfo("Info", "No screenshot files found!")
        return
    image_files.sort(reverse=True)
    latest_image = image_files[0][1]
    try:
        os.startfile(latest_image)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open file: {str(e)}")

def open_last_recording():
    video_dir = "video"
    if not os.path.exists(video_dir):
        messagebox.showwarning("Warning", "Recording folder does not exist!")
        return
    video_files = []
    for f in os.listdir(video_dir):
        if f.lower().endswith((".mp4", ".avi", ".mov", ".mkv")):
            full_path = os.path.join(video_dir, f)
            video_files.append((os.path.getmtime(full_path), full_path))
    if not video_files:
        messagebox.showinfo("Info", "No recording files found!")
        return
    video_files.sort(reverse=True)
    latest_video = video_files[0][1]
    try:
        os.startfile(latest_video)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open file: {str(e)}")

def open_last_audio():
    audio_dir = "audio"
    if not os.path.exists(audio_dir):
        messagebox.showwarning("Warning", "Audio folder does not exist!")
        return
    audio_files = []
    for f in os.listdir(audio_dir):
        if f.lower().endswith((".mp3", ".wav", ".aac", ".flac")):
            full_path = os.path.join(audio_dir, f)
            audio_files.append((os.path.getmtime(full_path), full_path))
    if not audio_files:
        messagebox.showinfo("Info", "No audio files found!")
        return
    audio_files.sort(reverse=True)
    latest_audio = audio_files[0][1]
    try:
        os.startfile(latest_audio)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open file: {str(e)}")

def Updatelog():
    Text="""CapVisionPro - Update Log
Version 1.99.01 Insider
- Initial official release
- Camera preview and real-time display
- Video recording support
- Screenshot capture (JPG / PNG / BMP)
- Timestamp watermark
- Image flip / mirror function
- Independent audio recording
- Config import & export system
- Built-in CVScript interpreter
- MIT open source license
- Log and diagnostics functions
- New UI layout and performance optimization
- Added more CVScript commands
- Improved camera compatibility
- Enhanced recording stability
- Fixed known minor bugs
- Updated open source copyright information
- Overall user experience improvement
- Added brightness & contrast adjustment
- Added log file viewer
- Fixed shortcut key conflicts
- Optimized pause recording logic
- Added auto-segment recording
- Added capture format selection
- Changed main script extension to .pyw (no console window on startup)
- Kept administrator privilege support (prevents "access denied" errors)
- Removed console-based copyright prints
- Moved all legal & credit info into the GUI (via Help menu)
- Added dedicated Credits entry with consistent keyboard shortcut
"""
    messagebox.showinfo("Update Log",Text)

file_menu.add_command(label="Minimize to Tray", command=lambda:minimize_to_tray(), accelerator="Alt-m")
file_menu.add_command(label="System Info", command=show_system_info, accelerator="Alt-s")
file_menu.add_command(label="Close Subwindow", command=close_current_subwindow, accelerator="Alt-w / ESC")
file_menu.add_separator()
file_menu.add_command(label="Exit Program", command=lambda: exitwindows(), accelerator="Alt-p")
menu_bar.add_cascade(label="File", menu=file_menu)
def ColourMode():
    path = filedialog.askopenfilename(title="Select Image to Open",filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")])
    if not path:
        return
    read_image=cv2.imread(path,1)
    cv2.imshow(f"Image - {path}",read_image)
    cv2.waitKey(0)
    try:
        cv2.destroyWindow(f"Image - {path}")
    except cv2.error:
        pass
def Graymode():
    path=filedialog.askopenfilename(title="Select Image to Open",filetypes=[("Image files","*.png *.jpg *.jpeg *.bmp")])
    if not path:
        return
    read_image = cv2.imread(path, 0)
    win_name=f"Gray Image - {path}"
    cv2.imshow(win_name,read_image)
    cv2.waitKey(0)
    try:
        cv2.destroyWindow(win_name)
    except cv2.error:
        pass
def Get_Image_Information():
    path=filedialog.askopenfilename(title="Select Image to Open",filetypes=[("Image files","*.png *.jpg *.jpeg *.bmp")])
    if not path:
        return
    image_colour=cv2.imread(path,1)
    try:
        if image_colour is None:
            messagebox.showerror("Error", "Failed to read image file!")
            return
        color_info = f"""
Get Image Color Information:
shape = {image_colour.shape}  (height × width × channels)
size = {image_colour.size}  (total pixels)
dtype = {image_colour.dtype}
"""
        print(color_info)
        image_gray = cv2.imread(path, 0)
        gray_info=f"""Get Image Gray Information:
shape = {image_gray.shape}  (height × width)
size = {image_gray.size}  (total pixels)
dtype = {image_gray.dtype}
"""
        print(gray_info)
        base_name = os.path.splitext(os.path.basename(path))[0]
        txt_path = os.path.join(os.path.dirname(path), f"{base_name}_info.txt")
        full_info=f"""
=============================================
CapVisionPro Image Information Report
Version: 1.99.01 Insider
=============================================
File Path: {path}
File Name: {os.path.basename(path)}
=============================================
{color_info}
=============================================
{gray_info}
=============================================
Generated at: {time.ctime()}
=============================================
"""
        f=open(txt_path,"w",encoding="UTF-8")
        f.write(full_info)
        f.close()
        messagebox.showinfo("Image Information Saved",f"Image information has been successfully saved to:\n{txt_path}")
        write_log(f"Image info saved: {txt_path}")
    except Exception as e:
        error_msg = f"Failed to get image info: {str(e)}"
        messagebox.showerror("Error", error_msg)
        write_log(error_msg)
        return
def ColourVideomode():
    path = filedialog.askopenfilename(title="Select Video to Open", filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv")])
    if not path:
        return
    
    video_cap = cv2.VideoCapture(path)
    if not video_cap.isOpened():
        messagebox.showerror("Error", "Failed to open video file!")
        return
    
    fps = int(video_cap.get(cv2.CAP_PROP_FPS))
    delay = max(1, int(1000 / fps))
    paused = False
    win_name = f"Video - Color Mode: {os.path.basename(path)}"
    
    while True:
        if not paused:
            ret, frame = video_cap.read()
            if not ret:
                break
            cv2.imshow(win_name, frame)
        
        key = cv2.waitKey(delay) & 0xFF
        if key == ord('q'):
            break
        elif key == ord(' '):
            paused = not paused
    
    video_cap.release()
    try:
        cv2.destroyWindow(win_name)
    except cv2.error:
        pass
def GrayVideomode():
    path = filedialog.askopenfilename(title="Select Video to Open", filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv")])
    if not path:
        return
    
    video_cap = cv2.VideoCapture(path)
    if not video_cap.isOpened():
        messagebox.showerror("Error", "Failed to open video file!")
        return
    
    fps = int(video_cap.get(cv2.CAP_PROP_FPS))
    delay = max(1, int(1000 / fps))
    paused = False
    win_name = f"Video - Gray Mode: {os.path.basename(path)}"
    
    while True:
        if not paused:
            ret, frame = video_cap.read()
            if not ret:
                break
            gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            cv2.imshow(win_name, gray_frame)
        
        key = cv2.waitKey(delay) & 0xFF
        if key == ord('q'):
            break
        elif key == ord(' '):
            paused = not paused
    
    video_cap.release()
    try:
        cv2.destroyWindow(win_name)
    except cv2.error:
        pass
def Open_Audio():
    path = filedialog.askopenfilename(
        title="Select Audio File to Open",
        filetypes=[("Audio files", "*.mp3;*.mp2;*.aac;*.m4a;*.flac;*.wav;*.ogg;*.opus;*.alac;*.ape;*.wma;*.ac3;*.dts;*.aiff;*.au;*.amr;*.mid;*.midi;*.ra;*.tak;*.caf")]
    )
    if not path:
        return

    audio_win = tk.Toplevel(root)
    audio_win.title("Audio Player Control")
    audio_win.geometry("500x280")
    audio_win.resizable(False, False)
    try:
        audio_win.iconbitmap("Capicon.ico")
    except:
        pass
    audio_win.bind("<Escape>", lambda e: audio_win.destroy())

    tk.Label(audio_win, text="File:", font=("Arial", 10, "bold")).place(x=10, y=10)
    path_label = tk.Label(audio_win, text=os.path.basename(path), wraplength=480, anchor="w")
    path_label.place(x=50, y=10, width=440)

    tk.Label(audio_win, text="Volume:", font=("Arial", 10)).place(x=10, y=50)
    vol_var = tk.IntVar(value=100)
    vol_scale = tk.Scale(audio_win, from_=0, to=200, variable=vol_var, orient=tk.HORIZONTAL)
    vol_scale.place(x=80, y=40, width=380)

    tk.Label(audio_win, text="Speed:", font=("Arial", 10)).place(x=10, y=100)
    speed_var = tk.DoubleVar(value=1.0)
    speed_combo = ttk.Combobox(audio_win, textvariable=speed_var,
                               values=[0.5, 0.75, 1.0, 1.25, 1.5, 2.0],
                               state="readonly", width=8)
    speed_combo.place(x=80, y=100)

    loop_var = tk.BooleanVar()
    tk.Checkbutton(audio_win, text="Loop", variable=loop_var).place(x=200, y=100)

    topmost_var = tk.BooleanVar()
    tk.Checkbutton(audio_win, text="Always on Top", variable=topmost_var).place(x=280, y=100)

    def start_ffplay():
        volume = vol_var.get()
        speed = speed_var.get()

        args = [
            "ffplay",
            path,
            "-volume", str(volume),
            "-autoexit"
        ]

        if speed != 1.0:
            args.extend(["-filter:a", f"atempo={speed}"])

        if loop_var.get():
            args.extend(["-loop", "-1"])

        if topmost_var.get():
            args.append("-alwaysontop")

        subprocess.Popen(
            args,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

    tk.Button(audio_win, text="Start Play", bg="#87E8DE", command=start_ffplay,
              font=("Arial", 11, "bold")).place(x=100, y=180, width=140, height=40)

    tk.Button(audio_win, text="Close", command=audio_win.destroy,
              font=("Arial", 11)).place(x=260, y=180, width=140, height=40)
def BinaryNormal():
    path=filedialog.askopenfilename(title="Select Image to Open",filetypes=[("Image files","*.png *.jpg *.jpeg *.bmp")])
    if not path:
        return
    img = cv2.imread(path)
    img = cv2.resize(img, (640, 480))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    threshold_val = 127  
    max_val = 255
    _, binary_normal = cv2.threshold(gray,threshold_val,max_val,cv2.THRESH_BINARY)
    cv2.imshow(f"Binary Normal-{path}", binary_normal)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
def BinaryInvert():
    path=filedialog.askopenfilename(title="Select Image to Open",filetypes=[("Image files","*.png *.jpg *.jpeg *.bmp")])
    if not path:
        return
    img = cv2.imread(path)
    img = cv2.resize(img, (640, 480))
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    threshold_val = 127  
    max_val = 255
    _, binary_invert = cv2.threshold(gray,threshold_val,max_val,cv2.THRESH_BINARY_INV )
    cv2.imshow(f"Binary Invert-{path}", binary_invert)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
def VideoBinaryNormal():
    path = filedialog.askopenfilename(title="Select Video to Open", filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv")])
    if not path:
        return
    video_cap=cv2.VideoCapture(path)
    if not video_cap.isOpened():
        messagebox.showerror("Error", "Failed to open video file!")
        return
    fps = int(video_cap.get(cv2.CAP_PROP_FPS))
    delay = max(1, int(1000 / fps))
    paused = False
    win_name = f"Video - Binary Mode: {os.path.basename(path)}"
    
    thresh = 127
    max_val = 255
    
    while True:
        if not paused:
            ret, frame = video_cap.read()
            if not ret:
                break
            gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            _, binary_frame = cv2.threshold(gray_frame, thresh, max_val, cv2.THRESH_BINARY)
            cv2.imshow(win_name, binary_frame)
        
        key = cv2.waitKey(delay) & 0xFF
        if key == ord('q'):
            break
        elif key == ord(' '):
            paused = not paused
    
    video_cap.release()
    try:
        cv2.destroyWindow(win_name)
    except cv2.error:
        pass
def VideoBinaryInvert():
    path = filedialog.askopenfilename(title="Select Video to Open", filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv")])
    if not path:
        return
    video_cap=cv2.VideoCapture(path)
    if not video_cap.isOpened():
        messagebox.showerror("Error", "Failed to open video file!")
        return
    fps = int(video_cap.get(cv2.CAP_PROP_FPS))
    delay = max(1, int(1000 / fps))
    paused = False
    win_name = f"Video - Invert Binary Mode: {os.path.basename(path)}"
    
    thresh = 127
    max_val = 255
    
    while True:
        if not paused:
            ret, frame = video_cap.read()
            if not ret:
                break
            gray_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            _, binary_frame = cv2.threshold(gray_frame, thresh, max_val, cv2.THRESH_BINARY_INV)
            cv2.imshow(win_name, binary_frame)
        
        key = cv2.waitKey(delay) & 0xFF
        if key == ord('q'):
            break
        elif key == ord(' '):
            paused = not paused
    
    video_cap.release()
    try:
        cv2.destroyWindow(win_name)
    except cv2.error:
        pass
OpenFile_menu=tk.Menu(menu_bar,tearoff=0)
menu_bar.add_cascade(label="Open", menu=OpenFile_menu)
OpenFile_menu.add_command(label="Open Last Screenshot", command=open_last_screenshot, accelerator="Alt+Shift+S")
OpenFile_menu.add_command(label="Open Last Recording", command=open_last_recording, accelerator="Alt+Shift+D")
OpenFile_menu.add_command(label="Open Last Audio", command=open_last_audio, accelerator="Alt+Shift+Z")
OpenFile_menu.add_separator()
OpenFile_menu.add_command(label="Open Image Folder", command=OpenImageFolder, accelerator="Alt+I")
OpenFile_menu.add_command(label="Open Video Folder", command=OpenVideoFolder, accelerator="Alt+V")
OpenFile_menu.add_command(label="Open Audio Folder", command=OpenAudioFolder, accelerator="Alt+Shift+Y")
OpenFile_menu.add_command(label="Open Log Folder", command=OpenLogFolder, accelerator="Alt+G")
OpenFile_menu.add_separator()
OpenExistingFile=tk.Menu(OpenFile_menu,tearoff=0)
OpenExistingFile.add_command(label="Color Mode", command=ColourMode, accelerator="Ctrl+Shift+C")
OpenExistingFile.add_command(label="Gray Mode", command=Graymode, accelerator="Ctrl+Shift+G")
OpenExistingFile.add_command(label="Binary Normal",command=BinaryNormal)
OpenExistingFile.add_command(label="Binary Invert",command=BinaryInvert)
OpenFile_menu.add_cascade(label="Open Image (View)", menu=OpenExistingFile)
OpenFile_menu.add_separator()
OpenVideoFile=tk.Menu(OpenFile_menu,tearoff=0)
OpenVideoFile.add_command(label="Color Mode",accelerator="Ctrl+Shift+V",command=ColourVideomode)
OpenVideoFile.add_command(label="Gray Mode",accelerator="Ctrl+Shift+B",command=GrayVideomode)
OpenVideoFile.add_command(label="Binary Normal",command=VideoBinaryNormal)
OpenVideoFile.add_command(label="Binary Invert",command=VideoBinaryInvert)
OpenFile_menu.add_cascade(label="Open Video (View)", menu=OpenVideoFile)
OpenFile_menu.add_separator()
OpenFile_menu.add_command(label="Open Audio",accelerator="Ctrl+Shift+Alt+A",command=Open_Audio)
OpenFile_menu.add_separator()
OpenFile_menu.add_command(label="Get Image Information", command=Get_Image_Information, accelerator="Ctrl+Shift+I")
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="About", command=lambda:show_about(), accelerator="Alt-a")

def Show_help():
    file_path = "Caphelp.html"
    if not os.path.exists(file_path):
        messagebox.showerror("Help Document Missing","The help document could not be loaded. The file may be missing, corrupted, or inaccessible.")
        return
    try:
        os.startfile(file_path)
    except Exception:
        messagebox.showerror("Open Failed", "Failed to open the help document.")
PASSWORD_FILE = "password.bin"

def set_password():
    pwd = simpledialog.askstring("Set Password", "Set your new password:", show="*")
    if not pwd:
        messagebox.showwarning("Warning", "Password cannot be empty!")
        return
    encrypted = hashlib.sha256(pwd.encode()).digest()
    with open(PASSWORD_FILE, "wb") as f:
        f.write(encrypted)
    messagebox.showinfo("Success", "Password has been set successfully!\nFile: password.bin")

def check_password(input_pwd):
    if not input_pwd:
        return False
    if not os.path.exists(PASSWORD_FILE):
        return True
    with open(PASSWORD_FILE, "rb") as f:
        stored = f.read()
    encrypted_input = hashlib.sha256(input_pwd.encode()).digest()
    return stored == encrypted_input
def verify_on_start():
    if not os.path.exists(PASSWORD_FILE):
        return
    root.withdraw()
    pwd_win = tk.Toplevel(root)
    pwd_win.title("CapVisionPro - Unlock")
    pwd_win.geometry("320x180")
    pwd_win.resizable(False, False)
    pwd_win.attributes("-topmost", True)
    pwd_win.protocol("WM_DELETE_WINDOW", lambda: sys.exit())
    try:
        pwd_win.iconbitmap("Capicon.ico")
    except:
        pass
    tk.Label(pwd_win, text="Enter Password to Unlock", font=("Segoe UI", 12, "bold")).pack(pady=20)
    pwd_entry = tk.Entry(pwd_win, show="*", font=("Segoe UI", 11), width=25)
    pwd_entry.pack(pady=5)
    pwd_entry.focus_set()
    tip_label = tk.Label(pwd_win, text="", font=("Segoe UI", 9), fg="red")
    tip_label.pack(pady=5)
    attempt_left = 3

    def check_pwd(event=None):
        nonlocal attempt_left
        input_pwd = pwd_entry.get().strip()
        if not input_pwd:
            tip_label.config(text="Password cannot be empty!")
            return
        if check_password(input_pwd):
            pwd_win.destroy()
            root.deiconify()
            return
        attempt_left -= 1
        if attempt_left > 0:
            tip_label.config(text=f"Wrong password! {attempt_left} chance(s) left")
            pwd_entry.delete(0, tk.END)
        else:
            tip_label.config(text="Too many failed attempts! Exiting...")
            pwd_entry.config(state=tk.DISABLED)
            pwd_win.after(1000, sys.exit)

    tk.Button(pwd_win, text="Unlock", command=check_pwd, width=10, font=("Segoe UI", 10)).pack(pady=8)
    pwd_entry.bind("<Return>", check_pwd)
    pwd_win.grab_set()
    pwd_win.wait_window()
help_menu.add_command(label="View Help", command=lambda:Show_help(), accelerator="Alt-h")
help_menu.add_command(label="License Agreement", command=lambda: LicenseAgreement(), accelerator="Alt-l")
help_menu.add_separator()

def LicenseAgreement():
    win = tk.Toplevel(root)
    win.title("Open Source License Agreement")
    win.geometry("600x500")
    win.resizable(False, False)
    win.bind("<Escape>", lambda e: win.destroy())
    try:
        win.iconbitmap("Capicon.ico")
    except:
        pass
    text = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("Arial", 9))
    text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
    License_Text = """
Open Source Software License Agreement (Version 1.1)
This Open Source Software License Agreement (this “Agreement”) is entered into as of the Effective Date between the Copyright Owner and any individual or Legal Entity obtaining, using, modifying, distributing, or otherwise exercising rights under this Agreement (“Licensee”). This Agreement governs the use, copying, modification, distribution, and maintenance of the Software and defines the rights and obligations of all parties.

By accessing, downloading, installing, using, modifying, or distributing the Software, Licensee agrees to be fully bound by all terms of this Agreement. If Licensee does not agree to any term herein, Licensee must immediately cease all access, use, modification, and distribution of the Software.

---

## 1. Definitions
1.1 **Copyright Owner**
The individual(s), entity(ies), or authorized representatives holding exclusive copyright and related intellectual property rights in the Software, including original authors, contributors, and authorized agents.

1.2 **Software**
The open‑source software package, including all Source Code, Object Code, Documentation, configuration files, scripts, libraries, modules, media resources, and related components. The Software includes updates, patches, and bug fixes unless such materials are licensed separately.

1.3 **Source Code**
The preferred human‑readable form of the Software for modification, including code files, header files, build scripts, makefiles, configuration templates, and related definition files. Source Code does not include compiled binaries, obfuscated code, encrypted code, or content not reasonably modifiable.

1.4 **Object Code**
Any compiled, assembled, or interpreted executable form of the Software, including binary files, bytecode, dynamically linked libraries, and runtime executable files.

1.5 **Documentation**
All written or electronic materials accompanying the Software, including user manuals, developer guides, technical specifications, API references, tutorials, and help documents.

1.6 **Derivative Works**
Any work based on or derived from the Software, including:
(a) modifications, enhancements, or deletions to the Software;
(b) translations, adaptations, or ports to other platforms, architectures, or languages;
(c) works that incorporate a substantial portion of the Software;
(d) combinations of the Software with other software that constitute a derivative work under applicable copyright law.

Mere interaction through public APIs without modification or inclusion of the Software’s Source or Object Code is **not** a Derivative Work.

1.7 **Legal Entity**
An entity and its parent, subsidiaries, and affiliates under common control. “Control” means direct or indirect ownership of 50% or more voting equity, or the power to direct management and policies.

1.8 **Effective Date**
The date Licensee first accesses, downloads, installs, or uses the Software. For Legal Entities, the Effective Date is when an authorized representative first uses the Software on its behalf.

1.9 **Redistribution**
Making the Software or Derivative Works available to third parties by any means, including commercial distribution, non‑commercial sharing, sublicensing, renting, leasing, or network deployment.

1.10 **Patent Claims**
All patent claims owned or controlled by the Copyright Owner that are necessarily infringed by the use or distribution of the Software as provided.

---

## 2. Grant of Rights
Subject to Licensee’s compliance with this Agreement, the Copyright Owner grants Licensee a **worldwide, non‑exclusive, royalty‑free, perpetual license** to:
(a) Use, run, and execute the Software for personal, educational, internal business, or commercial purposes;
(b) Study, modify, and create Derivative Works of the Software;
(c) Reproduce and distribute the Software, original or modified, to third parties;
(d) Distribute Object Code, provided that corresponding Source Code is made available as required.

## 3. Conditions on Redistribution
3.1 All redistributions of the Software or Derivative Works must include a copy of this Agreement and the original copyright notice.
3.2 If distributing Object Code only, Licensee must provide corresponding Source Code in a machine‑readable, accessible form.
3.3 Licensee may not remove, alter, or obscure copyright, patent, trademark, or attribution notices.
3.4 Licensee may not use the name, trademarks, service marks, or logos of the Copyright Owner without separate written permission.

## 4. Attribution and Copyright
Licensee must clearly retain all original copyright statements, author attributions, license texts, and legal notices in all copies and Derivative Works. Any modified versions must prominently state that changes have been made and carry prominent attribution to the original Copyright Owner.

## 5. Prohibited Uses
Licensee may not use the Software:
(a) For illegal, infringing, abusive, or unauthorized surveillance activities;
(b) To violate privacy rights, data protection laws, or third‑party property rights;
(c) In military, weapons, surveillance, or harmful systems where failure could cause harm;
(d) To reverse engineer, decompile, or bypass security measures except as permitted by law;
(e) To sublicense the Software under terms that conflict with this Agreement.
(f) The Software is licensed for NON-COMMERCIAL USE ONLY.
    Any commercial use, including but not limited to selling, renting,
    licensing, or using the Software in business services, is STRICTLY PROHIBITED.

(g) You may not use the Software for any commercial purpose,
    whether direct or indirect, without prior written permission from the Copyright Owner.
## 6. Intellectual Property
All intellectual property rights in the original Software remain the exclusive property of the Copyright Owner. Licensee obtains only the rights expressly granted herein. Nothing in this Agreement transfers ownership of copyright, patents, trademarks, or trade secrets.

## 7. Disclaimer of Warranty
THE SOFTWARE IS PROVIDED “AS IS” WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, NONINFRINGEMENT, OR QUALITY OF PERFORMANCE. THE COPYRIGHT OWNER DOES NOT WARRANT THAT THE SOFTWARE IS ERROR‑FREE, SECURE, UNINTERRUPTED, OR COMPATIBLE WITH ALL DEVICES.

## 8. Limitation of Liability
IN NO EVENT SHALL THE COPYRIGHT OWNER BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

## 9. Termination
This Agreement automatically terminates if Licensee breaches any material term. Upon termination, Licensee must cease all use, modification, and distribution of the Software and destroy all copies in its possession.

## 10. General Provisions
This Agreement shall be governed by the laws of the jurisdiction where the Copyright Owner resides, without regard to conflict of law rules. If any provision is held invalid or unenforceable, the remaining provisions shall remain in full force.

---

# Important Notices for Camera Software
## 1. Privacy & Legal Compliance
- This camera software is designed **for lawful personal, educational, or testing purposes only**.
- You must obtain **explicit, voluntary consent** from any person whose image, voice, or likeness is captured or recorded.
- Unauthorized surveillance, secret recording, peeping, or invasion of privacy is **strictly prohibited** and may violate local laws, criminal statutes, and data protection regulations (such as GDPR, CCPA, domestic privacy laws, etc.).
- You are solely responsible for compliance with all applicable laws related to recording, storage, transmission, and disclosure of audiovisual data.

## 2. Security Risks
- Camera feeds may contain sensitive personal information. Use encryption and secure storage where applicable.
- Avoid exposing camera streams to public networks or untrusted devices to prevent unauthorized access, hijacking, or leakage.
- The Copyright Owner is not responsible for any security incidents caused by improper configuration, network environment, or user negligence.

## 3. Hardware Compatibility
- This software may depend on system drivers, camera firmware, and platform permissions.
- You must grant necessary permissions (camera access, microphone access, storage) in accordance with your operating system settings.
- The software may not support all devices, and no guarantee is provided for specific hardware performance.

## 4. Prohibited Usage
- Do not use this software for illegal monitoring, harassment, intimidation, or unauthorized data collection.
- Do not use this software in restricted areas, private spaces, or areas where recording is forbidden by law or regulation.
- Do not use this software to infringe the portrait rights, privacy rights, or intellectual property rights of others.

---

# Copyright & Intellectual Property Statement
1. **Original Copyright Ownership**
All original code, structure, logic, interface design, documentation, and related content of this camera software are the exclusive property of the Copyright Owner. Copyright is protected under the Berne Convention and applicable national copyright laws.

2. **Attribution Requirement**
When reusing, modifying, or distributing this software, you must:
- Keep all original copyright headers, author names, and license statements;
- Clearly indicate the source of the original work;
- Not falsely claim authorship or ownership of the original software.

3. **Trademark & Brand Restrictions**
You may not use the project name, developer name, logo, or trademarks of the Copyright Owner for promotion, endorsement, or public representation without prior written permission.

4. **Contributor Rights**
Any voluntary contributions (code patches, suggestions, documentation improvements) shall be licensed under the same open‑source license. Contributors grant the Copyright Owner a perpetual, worldwide, non‑exclusive license to use their contributions.

5. **Commercial Use**
Commercial use, redistribution, or integration into commercial products is permitted **only if fully compliant with this license**. You may not rebrand or sublicense the software in a way that misleads users or evades attribution obligations.

6. **Piracy & Unauthorized Modification**
Unauthorized removal of copyright information, obfuscation of source code, or distribution of cracked/modified versions constitutes copyright infringement and may lead to civil or criminal liability.

"""
    text.insert(tk.END, License_Text)
    text.config(state=tk.DISABLED)
    def close():
        win.destroy()
    btn = tk.Button(win, text="Agree & Close", command=close, width=18, bg="#6be629")
    btn.pack(pady=10)
def OpenCode():
    OpencodeAsk=messagebox.askyesno("CapVisionPro","Continue anyway?")
    if OpencodeAsk == True:
        Urllink="https://feishu.doubao.com/drive/file/LVu4bcgqEo5m4GxlZVjcwYcNnhf"
        try:
            webbrowser.open(Urllink)
            messagebox.showinfo("Password","key:MeterCapVisionPro_XSDHJ")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open link: {str(e)}")
        return
    elif OpencodeAsk == False:
        return
def view_log_file():
    log_dir = "ErrorLog"
    log_file = os.path.join(log_dir, f"CapVision_Log_{time.strftime('%Y%m%d')}.txt")
    
    if not os.path.exists(log_dir):
        messagebox.showwarning("Warning", "Log folder does not exist!")
        return
        
    if not os.path.exists(log_file):
        messagebox.showwarning("Warning", "No log file found for today!")
        return
        
    try:
        os.startfile(log_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open log file: {str(e)}")
def Credits():
    info="""This software uses the following open-source libraries:

Python
Copyright (c) 2001-2023 Python Software Foundation.
Copyright (c) 2000 BeOpen.com.
Copyright (c) 1995-2001 Corporation for National Research Initiatives.
Copyright (c) 1991-1995 Stichting Mathematisch Centrum, Amsterdam.

OpenCV (Open Source Computer Vision Library)
Copyright (C) 2000-2025, Intel Corporation, all rights reserved.
Copyright (C) 2009-2025, Willow Garage, all rights reserved.
Copyright (C) 2015-2025, OpenCV Foundation, all rights reserved.
Copyright (C) 2015-2025, Itseez, Inc., all rights reserved.
Third party copyrights are property of their respective owners.

Tkinter
Copyright (c) 1989-1994 The Regents of the University of California.
Copyright (c) 1994-2023 Sun Microsystems, Inc.
Copyright (c) 1994-2023 Digital Equipment Corporation.
Copyright (c) 1994-2023 Ajuba Solutions.
Copyright (c) 1994-2023 ActiveState Corporation.
Copyright (c) 2000-2023 The Tcl Core Team.
Copyright (c) 2000-2023 The Tk Core Team.
Copyright (c) 2001-2023 Python Software Foundation.

FFmpeg, PyAudio
Copyright (c) 2000-2003 Fabrice Bellard
Copyright (c) 2003-2025 The FFmpeg developers.
Copyright (c) 2006-2014 Hubert Pham"""
    messagebox.showinfo("Credits",info)
help_menu.add_command(label="Shortcuts", command=lambda: AboutKeyboard(), accelerator="Alt-k")
help_menu.add_command(label="Update Log", command=Updatelog, accelerator="Alt+Shift+U")
help_menu.add_separator()
help_menu.add_command(label="Open Source Code", command=OpenCode, accelerator="Alt+Shift+O")
help_menu.add_command(label="View Log File", command=view_log_file, accelerator="Alt+L")
help_menu.add_command(label="Credits",command=Credits,accelerator="Alt+Shift+@")
menu_bar.add_cascade(label="Help", menu=help_menu)

Tool_menu=tk.Menu(menu_bar,tearoff=0)

def toggle_flip():
        global flip_mode
        flip_mode = (flip_mode + 1) % 3
        names = ["Normal", "Mirror", "Flip+Mirror"]
        status_bar.config(text=f"Mode: {names[flip_mode]}")

def DeleteTools():
    Tools = tk.Toplevel(root)
    Tools.title("Cleanup Tools")
    Tools.geometry("400x350")
    Tools.bind("<Escape>", lambda event: Tools.destroy())
    try:
        Tools.iconbitmap("Capicon.ico")
    except:
        pass
    Tools.resizable(False, False)
    Label = tk.Label(Tools, text="Clean up files in one click", font=("Arial", 12))
    Label.pack(pady=15)

    def clear_log_file():
        result = messagebox.askyesno("Confirm", "Are you sure you want to clear ALL files in the ErrorLog folder?")
        if not result:
            return
        if os.path.exists("ErrorLog"):
            for file in os.listdir("ErrorLog"):
                file_path = os.path.join("ErrorLog", file)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception:
                    continue
        messagebox.showinfo("Success", "All log files have been cleared successfully.")

    def clear_Image_file():
        result = messagebox.askyesno("Confirm", "Are you sure you want to clear ALL files in the image folder?")
        if not result:
            return
        if os.path.exists("image"):
            for file in os.listdir("image"):
                file_path = os.path.join("image", file)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception:
                    continue
        messagebox.showinfo("Success", "All image files have been cleared successfully.")

    def clear_Video_file():
        result = messagebox.askyesno("Confirm", "Are you sure you want to clear ALL files in the video folder?")
        if not result:
            return
        if os.path.exists("video"):
            for file in os.listdir("video"):
                file_path = os.path.join("video", file)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception:
                    continue
        messagebox.showinfo("Success", "All video files have been cleared successfully.")

    def clear_Audio_file():
        result = messagebox.askyesno("Confirm", "Are you sure you want to clear ALL files in the audio folder?")
        if not result:
            return
        if os.path.exists("audio"):
            for file in os.listdir("audio"):
                file_path = os.path.join("audio", file)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception:
                    continue
        messagebox.showinfo("Success", "All audio files have been cleared successfully.")

    tk.Button(Tools, text="Clear Log Files", font=("Arial", 11), width=20, bg="#00BFFF", command=clear_log_file).pack(pady=6)
    tk.Button(Tools, text="Clear Image Files", font=("Arial", 11), width=20, bg="#FFF8DC", command=clear_Image_file).pack(pady=6)
    tk.Button(Tools, text="Clear Video Files", font=("Arial", 11), width=20, bg="#8332AB", command=clear_Video_file).pack(pady=6)
    tk.Button(Tools, text="Clear Audio Files", font=("Arial", 11), width=20, bg="#00FF00", command=clear_Audio_file).pack(pady=6)
def toggle_watermark():
    global watermark_enabled
    watermark_enabled = not watermark_enabled
    status_bar.config(text=f"Watermark: {'ON' if watermark_enabled else 'OFF'}")

def capture_with_countdown():
    cwin = tk.Toplevel(root)
    cwin.title("Countdown Capture")
    cwin.geometry("200x100")
    cwin.resizable(False,False)
    cwin.attributes("-topmost",True)
    try:
        cwin.iconbitmap("Capicon.ico")
    except:
        pass
    lbl = tk.Label(cwin, font=("Arial",36,"bold"))
    lbl.pack(expand=True)
    
    def countdown(n=3):
        if n > 0:
            lbl.config(text=str(n))
            root.after(1000, countdown, n-1)
        else:
            cwin.destroy()
            CaptureImage()
    countdown()
def switch_camera():
    global cap, current_cam_index
    current_cam_index += 1
    if current_cam_index > 5:
        current_cam_index = 0
    try:
        if cap is not None:
            cap.release()
        
        cap = cv2.VideoCapture(current_cam_index, cv2.CAP_DSHOW)
        if not cap.isOpened():
            status_bar.config(text=f"Camera {current_cam_index} not found")
            return
            
        cap.set(cv2.CAP_PROP_FPS, 30)
        status_bar.config(text=f"Switched to Camera {current_cam_index}")
    except Exception:
        status_bar.config(text="Camera switch failed")
def rotate_selected_image():
    path = filedialog.askopenfilename(filetypes=[("Images","*.png *.jpg *.jpeg *.bmp *.tiff")])
    if not path:
        return
    try:
        img = cv2.imread(path)
        if img is None:
            messagebox.showerror("Error", "Failed to load image!")
            return
            
        img = cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
        cv2.imwrite(path, img)
        messagebox.showinfo("Success", "Rotated 90° clockwise")
    except Exception as e:
        messagebox.showerror("Error", f"Image rotation failed: {str(e)}")
def write_log(msg):
    os.makedirs("ErrorLog",exist_ok=True)
    log_file = os.path.join("ErrorLog", f"CapVision_Log_{time.strftime('%Y%m%d')}.txt")
    log_line = f"[{time.ctime()}] {msg}\n"
    try:
        with open(log_file,"a",encoding="utf-8") as f:
            f.write(log_line)
        print(log_line.strip())
    except:
        pass

CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 44100
WAVE_OUTPUT_FILENAME = "temp_audio.wav"

def start_audio_recording(is_audio_only=False):
    global p, audio_stream, audio_wave_file, temp_audio_path, audio_only_recording
    save_dir = "audio" if is_audio_only else "video"
    os.makedirs(save_dir, exist_ok=True)
    temp_audio_path = os.path.join(save_dir, WAVE_OUTPUT_FILENAME)
    audio_only_recording = is_audio_only
    p = pyaudio.PyAudio()
    try:
        audio_stream = p.open(format=FORMAT,
                            channels=CHANNELS,
                            rate=RATE,
                            input=True,
                            frames_per_buffer=CHUNK)
        audio_wave_file = wave.open(temp_audio_path, 'wb')
        audio_wave_file.setnchannels(CHANNELS)
        audio_wave_file.setsampwidth(p.get_sample_size(FORMAT))
        audio_wave_file.setframerate(RATE)
    except Exception as e:
        write_log(f"Audio recording init failed: {str(e)}")
        messagebox.showerror("Audio Error", f"Failed to start audio recording: {str(e)}")
        return
    def write_audio_frames():
        if (is_recording or audio_only_recording) and not recording_paused and audio_stream and audio_stream.is_active():
            try:
                data = audio_stream.read(CHUNK, exception_on_overflow=False)
                audio_wave_file.writeframes(data)
            except:
                pass
            root.after(10, write_audio_frames)
    write_audio_frames()
    write_log("Audio recording started" if not is_audio_only else "Audio-only recording started")

def stop_audio_recording():
    global p, audio_stream, audio_wave_file, audio_only_recording
    try:
        if audio_stream:
            audio_stream.stop_stream()
            audio_stream.close()
        if audio_wave_file:
            audio_wave_file.close()
        if p:
            p.terminate()
    except:
        pass
    audio_only_recording = False
    write_log("Audio recording stopped")

def start_audio_only_recording():
    global status_bar
    if audio_only_recording:
        status_bar.config(text="Already recording audio!")
        return
    start_audio_recording(is_audio_only=True)
    status_bar.config(text="Audio Only Recording in progress")
    write_log("Start audio-only recording")

def stop_audio_only_recording():
    global status_bar, temp_audio_path
    if not audio_only_recording:
        status_bar.config(text="Not recording audio!")
        return
    stop_audio_recording()
    status_bar.config(text="Converting audio to MP3...")
    root.update()
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    final_audio_name = f"{save_prefix.get()}{timestamp}.mp3"
    final_audio_path = os.path.join("audio", final_audio_name)
    ffmpeg_path = "ffmpeg.exe"
    if not os.path.exists(ffmpeg_path):
        messagebox.showerror("FFmpeg Error", "ffmpeg.exe not found!\nPlease put it in the software root directory.")
        status_bar.config(text="Audio saved as WAV (FFmpeg missing)")
        return
    cmd = [
        ffmpeg_path,
        "-y",
        "-i", temp_audio_path,
        "-b:a", "192k",
        final_audio_path
    ]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if os.path.exists(temp_audio_path):
            os.remove(temp_audio_path)
        if os.path.exists(final_audio_path):
            status_bar.config(text=f"🎙 Audio Saved: {final_audio_name}")
            write_log(f"Audio-only recording saved: {final_audio_path}")
            toaster = ToastNotifier()
            toaster.show_toast(
                title="CapVisionPro",
                msg=f"Audio saved: {final_audio_name}",
                icon_path="Capicon.ico",
                duration=3,
                threaded=True
            )
        else:
            status_bar.config(text="Audio recording failed!")
    except Exception as e:
        write_log(f"Audio convert error: {str(e)}")
        status_bar.config(text="Audio convert failed!")
def ffmpeg_merge_audio_video(video_path, audio_path, output_path):
    ffmpeg_path = "ffmpeg.exe"
    if not os.path.exists(ffmpeg_path):
        messagebox.showerror("FFmpeg Error", "ffmpeg.exe not found!\nPlease put ffmpeg.exe in the software root directory.")
        write_log("FFmpeg merge failed: ffmpeg.exe not found")
        return False
    cmd = [
        ffmpeg_path,
        "-y",
        "-i", video_path,
        "-i", audio_path,
        "-c:v", "copy",
        "-c:a", "aac",
        "-strict", "experimental",
        output_path
    ]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if os.path.exists(output_path):
            write_log(f"FFmpeg merge success: {output_path}")
            return True
        else:
            write_log("FFmpeg merge failed: no output file")
            return False
    except Exception as e:
        write_log(f"FFmpeg merge error: {str(e)}")
        messagebox.showerror("FFmpeg Error", f"Merge failed: {str(e)}")
        return False
def clean_temp_files():
    global temp_video_path, temp_audio_path
    for temp_path in [temp_video_path, temp_audio_path]:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                write_log(f"Clean temp file: {temp_path}")
            except:
                write_log(f"Clean temp file failed: {temp_path}")
    temp_video_path = ""
    temp_audio_path = ""

def set_save_prefix():
    p = simpledialog.askstring("Set Prefix", "Enter file name prefix:", initialvalue=save_prefix.get())
    if p:
        save_prefix.set(p)
        status_bar.config(text=f"File prefix set to: {p}")

def set_brightness(val):
    global brightness_value
    brightness_value = int(val)
    try:
        cap.set(cv2.CAP_PROP_BRIGHTNESS, brightness_value)
        status_bar.config(text=f"Brightness set to: {brightness_value}%")
    except:
        pass

def set_contrast(val):
    global contrast_value
    contrast_value = int(val)
    try:
        cap.set(cv2.CAP_PROP_CONTRAST, contrast_value)
        status_bar.config(text=f"Contrast set to: {contrast_value}%")
    except:
        pass

def open_adjustment_window():
    adj_win = tk.Toplevel(root)
    adj_win.title("Image Adjustment")
    adj_win.geometry("350x180")
    adj_win.resizable(False, False)
    adj_win.bind("<Escape>", lambda e: adj_win.destroy())
    try:
        adj_win.iconbitmap("Capicon.ico")
    except:
        pass
    tk.Label(adj_win, text="Brightness", font=("Arial", 10)).pack(pady=5, anchor="w", padx=20)
    brightness_scale = tk.Scale(adj_win, from_=0, to=100, orient=tk.HORIZONTAL, length=300, command=set_brightness)
    brightness_scale.set(brightness_value)
    brightness_scale.pack(padx=20)
    tk.Label(adj_win, text="Contrast", font=("Arial", 10)).pack(pady=5, anchor="w", padx=20)
    contrast_scale = tk.Scale(adj_win, from_=0, to=100, orient=tk.HORIZONTAL, length=300, command=set_contrast)
    contrast_scale.set(contrast_value)
    contrast_scale.pack(padx=20)

def set_capture_format():
    format_win = tk.Toplevel(root)
    format_win.title("Capture Format Settings")
    format_win.geometry("300x150")
    format_win.resizable(False, False)
    format_win.bind("<Escape>", lambda e: format_win.destroy())
    try:
        format_win.iconbitmap("Capicon.ico")
    except:
        pass
    tk.Label(format_win, text="Select Default Capture Format", font=("Arial", 11, "bold")).pack(pady=15)
    format_list = ["bmp", "png", "jpg"]
    format_combo = ttk.Combobox(format_win, textvariable=capture_format, values=format_list, state="readonly", width=15)
    format_combo.pack(pady=5)
    def save_format():
        status_bar.config(text=f"Default capture format set to: {capture_format.get()}")
        format_win.destroy()
    tk.Button(format_win, text="Save", command=save_format, width=10).pack(pady=10)

def capture_as_jpg():
    if not Option2.get():
        return
    os.makedirs("image",exist_ok=True)
    ret,f = cap.read()
    if not ret:
        return
    ts = time.strftime("%Y%m%d_%H%M%S")
    fn = f"image/{save_prefix.get()}{ts}.jpg"
    cv2.imwrite(fn,f,[cv2.IMWRITE_JPEG_QUALITY,95])
    status_bar.config(text=f"Saved JPG: {os.path.basename(fn)}")
    write_log("Captured JPG")
    toaster = ToastNotifier()
    toaster.show_toast(title="CapVisionPro",msg=f"Screenshot saved: {os.path.basename(fn)}",icon_path="Capicon.ico",duration=3,threaded=True)

def toggle_pause_recording():
    global recording_paused
    if not is_recording and not audio_only_recording:
        status_bar.config(text="Not recording (video/audio)")
        return
    recording_paused = not recording_paused
    if is_recording:
        status_text = "Recording (Audio+Video) Paused" if recording_paused else "Recording (Audio+Video) Resumed"
    else:
        status_text = "Audio Only Paused" if recording_paused else "Audio Only Resumed"
    status_bar.config(text=status_text)
    write_log(f"Recording pause: {recording_paused}")

Tool_menu.add_command(label="Cleanup Tools", command=lambda:DeleteTools(), accelerator="Alt-d")
Tool_menu.add_separator()
Tool_menu.add_command(label="Set Password", command=set_password, accelerator="Alt-n")
Tool_menu.add_command(label="Unlock Software", command=verify_on_start, accelerator="Alt-u")
Tool_menu.add_separator()

def setup_settings_functions():
    Check_var1 = tk.BooleanVar(value=False)
    auto_cam_var = tk.BooleanVar(value=False)
    auto_min_var = tk.BooleanVar(value=False)
    custom_path_var = tk.BooleanVar(value=False)
    def set_startup_by_task(enable):
        exe_path = os.path.abspath(sys.argv[0])
        task_name = "CapVisionPro_AutoStart"
        if sys.argv[0].endswith(".py"):
            messagebox.showwarning("Warning", "Only available in EXE version.")
            return
        try:
            if enable:
                cmd = f'schtasks /create /tn "{task_name}" /tr "{exe_path}" /sc onlogon /ru Interactive /f'
            else:
                cmd = f'schtasks /delete /tn "{task_name}" /f'
            subprocess.Popen(cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass
    def select_save_directory():
        path = filedialog.askdirectory(title="Select Save Folder")
        if path:
            with open("save_path.txt", "w", encoding="utf-8") as f:
                f.write(path)
            messagebox.showinfo("Success", f"Path set:\n{path}")
    def save_all_settings():
        set_startup_by_task(Check_var1.get())
        with open("config.txt", "w", encoding="utf-8") as f:
            f.write(f"{auto_cam_var.get()}\n")
            f.write(f"{auto_min_var.get()}\n")
            f.write(f"{custom_path_var.get()}\n")
        messagebox.showinfo("Success", "All settings saved!")
    return Check_var1, auto_cam_var, auto_min_var, custom_path_var, select_save_directory, save_all_settings

def Settings():
    SettingsWin = tk.Toplevel(root)
    SettingsWin.geometry("400x380")
    SettingsWin.title("Application Settings")
    SettingsWin.resizable(False, False)
    SettingsWin.bind("<Escape>", lambda e: SettingsWin.destroy())
    try:
        SettingsWin.iconbitmap("Capicon.ico")
    except:
        pass
    Check_var1, auto_cam_var, auto_min_var, custom_path_var, select_save_directory, save_all_settings = setup_settings_functions()
    lbl_title = tk.Label(SettingsWin, text="CapVisionPro Settings", font=("Arial",12,"bold"))
    lbl_title.pack(pady=15)
    tk.Checkbutton(SettingsWin, text="Auto-run on Windows startup", variable=Check_var1, font=("Arial",11)).pack(anchor="w",padx=30,pady=5)
    tk.Checkbutton(SettingsWin, text="Auto-open camera on start", variable=auto_cam_var, font=("Arial",11)).pack(anchor="w",padx=30,pady=5)
    tk.Checkbutton(SettingsWin, text="Start minimized", variable=auto_min_var, font=("Arial",11)).pack(anchor="w",padx=30,pady=5)
    tk.Checkbutton(SettingsWin, text="Custom save directory", variable=custom_path_var, font=("Arial",11)).pack(anchor="w",padx=30,pady=5)
    tk.Button(SettingsWin, text="Choose Save Folder", command=select_save_directory, width=18).pack(pady=8)
    tk.Button(SettingsWin, text="Save Settings", command=save_all_settings, width=15).pack(pady=12)

def toggle_fps_display():
    global show_fps_info
    show_fps_info = not show_fps_info
    status_bar.config(text=f"FPS Display: {'ON' if show_fps_info else 'OFF'}")

def set_segment_time():
    val = simpledialog.askinteger("Auto Split", "Auto split time (seconds):", initialvalue=600)
    if val and val > 0:
        global segment_time
        segment_time = val
        status_bar.config(text=f"Auto split set to {val}s")
        write_log(f"Auto segment time set to {val}s")

def draw_audio_bar(frame, level=0):
    h, w = frame.shape[:2]
    bar_w = 22
    bar_h = int(160 * min(level, 1.0))
    x = w - 42
    y = h - 30 - bar_h
    cv2.rectangle(frame, (x, y), (x + bar_w, y + bar_h), (0, 255, 0), -1)
    cv2.rectangle(frame, (x, y-1), (x + bar_w, h-30), (200,200,200), 1)

def get_audio_level():
    try:
        if audio_stream and audio_stream.is_active():
            data = audio_stream.read(1024, exception_on_overflow=False)
            return min(audioop.rms(data, 2) / 32768, 1.0)
    except:
        pass
    return 0.0

def countdown_animation(count=3):
    for i in range(count, 0, -1):
        status_bar.config(text=f"Start recording after {i}...")
        root.update()
        time.sleep(1)
    status_bar.config(text="Recording...")

def enhance_frame_info(frame):
    if show_fps_info:
        fps = int(cap.get(cv2.CAP_PROP_FPS))
        w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        info = f"{w}x{h} | {fps} FPS"
        if is_recording:
            elap = int(time.time() - recording_start_time)
            info += f" | {elap//60:02d}:{elap%60:02d}"
        cv2.putText(frame, info, (10, 70), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0,255,255), 2)
    if (is_recording or audio_only_recording) and not recording_paused:
        draw_audio_bar(frame, get_audio_level())
    if recording_paused:
        h, w = frame.shape[:2]
        cv2.putText(frame, "PAUSED", (w//2 - 100, h//2), cv2.FONT_HERSHEY_SIMPLEX, 2, (0,0,255), 3)
    return frame

def check_auto_segment():
    if not is_recording or recording_paused:
        return
    if time.time() - recording_start_time >= segment_time:
        StopRecording()
        root.after(300, StartRecording)
        status_bar.config(text="Auto segment saved")
        write_log("Auto segment recording completed")
Tool_menu.add_command(label="Image Adjustment", command=open_adjustment_window, accelerator="Alt+E")
Tool_menu.add_command(label="Capture Format Settings", command=set_capture_format, accelerator="Alt+Shift+E")
Tool_menu.add_separator()
Tool_menu.add_command(label="Show FPS & Resolution", command=toggle_fps_display, accelerator="Alt+Shift+F")
Tool_menu.add_command(label="Auto Split Recording", command=set_segment_time, accelerator="Alt+Shift+L")
Tool_menu.add_separator()
Tool_menu.add_command(label="Application Settings", command=lambda:Settings(), accelerator="Alt-t")
Tool_menu.add_separator()
Tool_menu.add_command(label="Mirror / Flip", command=toggle_flip, accelerator="Alt+F")
Tool_menu.add_command(label="Toggle Watermark", command=toggle_watermark, accelerator="Alt+W")
Tool_menu.add_command(label="Countdown Capture", command=capture_with_countdown, accelerator="Alt+Shift+T")
Tool_menu.add_command(label="Pause/Resume Recording", command=toggle_pause_recording, accelerator="Alt+P")
Tool_menu.add_command(label="Switch Camera", command=switch_camera, accelerator="Alt+C")
Tool_menu.add_command(label="Rotate Image", command=rotate_selected_image, accelerator="Alt+R")
Tool_menu.add_separator()
Tool_menu.add_command(label="Set File Prefix", command=set_save_prefix, accelerator="Alt-X")
Tool_menu.add_command(label="Capture High Quality JPG", command=capture_as_jpg, accelerator="Alt+J")
Tool_menu.add_separator()
menu_bar.add_cascade(label="Tools",menu=Tool_menu)

Capture_menu = tk.Menu(menu_bar, tearoff=0)

def copy_screenshot_to_clipboard(image_path):
    try:
        img = Image.open(image_path)
        img.save("temp_clipboard.png", format="PNG")
        bmp_img = Image.open("temp_clipboard.png")
        bmp_img.save("temp_clipboard.bmp", format="BMP")
        with open("temp_clipboard.bmp", "rb") as f:
            bmp_data = f.read()[14:]
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, bmp_data)
    except Exception as e:
        if status_bar is not None:
            status_bar.config(text="Failed to copy image to clipboard")
        write_log(f"Clipboard copy failed: {str(e)}")
    finally:
        try:
            win32clipboard.CloseClipboard()
        except:
            pass
        if os.path.exists("temp_clipboard.png"):
            os.remove("temp_clipboard.png")
        if os.path.exists("temp_clipboard.bmp"):
            os.remove("temp_clipboard.bmp")

def CaptureImage():
    global status_bar, cap
    if not Option2.get():
        if status_bar is not None:
            status_bar.config(text="Capture disabled: Status Bar is off")
        return
    save_dir = "image"
    os.makedirs(save_dir, exist_ok=True)
    ret, frame = cap.read()
    if ret:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        format_ext = capture_format.get()
        photo_name = f"{save_prefix.get()}{timestamp}.{format_ext}"
        photo_path = os.path.join(save_dir, photo_name)
        if format_ext == "jpg":
            cv2.imwrite(photo_path, frame, [cv2.IMWRITE_JPEG_QUALITY, 95])
        else:
            cv2.imwrite(photo_path, frame)
        if status_bar is not None:
            status_bar.config(text=f"Capture completed: {photo_name}, Ready!")
        copy_screenshot_to_clipboard(photo_path)
        write_log(f"Capture completed: {photo_path}")
        toaster = ToastNotifier()
        toaster.show_toast(title="CapVisionPro",msg=f"Screenshot saved: {os.path.basename(photo_name)}",icon_path="Capicon.ico",duration=3,threaded=True)
    else:
        if status_bar is not None:
            status_bar.config(text="Capture failed: Camera not available")
        write_log("Capture failed: Camera not available")
def CaptureAndCopy():
    CaptureImage()
    image_dir = "image"
    files = sorted([f for f in os.listdir(image_dir)], key=lambda x: os.path.getmtime(os.path.join(image_dir, x)), reverse=True)
    if files:
        latest = os.path.join(image_dir, files[0])
        copy_screenshot_to_clipboard(latest)
        status_bar.config(text="Captured & Copied to Clipboard!")
    else:
        status_bar.config(text="Capture failed")

def StartRecording():
    global video_writer, is_recording, status_bar, cap, recording_start_time
    global temp_video_path, final_video_path
    Capture_menu.entryconfig(3, state="disabled")
    Capture_menu.entryconfig(4, state="normal")
    Capture_menu.entryconfig(6, state="disabled")
    Capture_menu.entryconfig(7, state="disabled")
    if not Option2.get():
        if status_bar is not None:
            status_bar.config(text="Recording disabled: Status Bar is off")
        return
    save_dir = "video"
    os.makedirs(save_dir, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    temp_video_name = f"temp_recording_{timestamp}.mp4"
    temp_video_path = os.path.join(save_dir, temp_video_name)
    final_video_name = f"{save_prefix.get()}{timestamp}.mp4"
    final_video_path = os.path.join(save_dir, final_video_name)
    fps = 30.0
    width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    width = width if width % 2 == 0 else width - 1
    height = height if height % 2 == 0 else height - 1
    fourcc = cv2.VideoWriter_fourcc(*'mp4v')
    video_writer = cv2.VideoWriter(temp_video_path, fourcc, fps, (width, height))
    if not video_writer.isOpened():
        if status_bar is not None:
            status_bar.config(text="Recording failed: Video writer init error")
        Capture_menu.entryconfig(3, state="normal")
        Capture_menu.entryconfig(4, state="disabled")
        Capture_menu.entryconfig(6, state="normal")
        Capture_menu.entryconfig(7, state="normal")
        write_log("Recording failed: Video writer init error")
        return
    start_audio_recording(is_audio_only=False)
    is_recording = True
    recording_paused = False
    recording_start_time = time.time()
    update_recording_timer()
    if status_bar is not None:
        status_bar.config(text="Recording (Audio+Video) in progress")
    write_log(f"Start recording: {final_video_path}")

def update_recording_timer():
    global recording_start_time
    if is_recording:
        if recording_paused:
            status_bar.config(text="⏸ Recording (Audio+Video) Paused")
            root.after(1000, update_recording_timer)
            return
        elapsed = int(time.time() - recording_start_time)
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        status_text = f"Recording (Audio+Video): {h:02d}:{m:02d}:{s:02d}"
        if status_bar is not None:
            status_bar.config(text=status_text)
        root.after(1000, update_recording_timer)
    else:
        if not audio_only_recording and status_bar is not None:
            status_bar.config(text="Ready")

def StopRecording():
    global video_writer, is_recording, status_bar
    Capture_menu.entryconfig(3, state="normal")
    Capture_menu.entryconfig(4, state="disabled")
    Capture_menu.entryconfig(6, state="normal")
    Capture_menu.entryconfig(7, state="normal")
    if not Option2.get() or not is_recording:
        return
    is_recording = False
    if video_writer is not None:
        video_writer.release()
        video_writer = None
    stop_audio_recording()
    if status_bar is not None:
        status_bar.config(text="Merging audio & video...")
        root.update()
    merge_success = ffmpeg_merge_audio_video(temp_video_path, temp_audio_path, final_video_path)
    clean_temp_files()
    if status_bar is not None:
        if merge_success:
            status_bar.config(text=f"Recording saved: {os.path.basename(final_video_path)}")
            write_log(f"Recording saved: {final_video_path}")
            toaster = ToastNotifier()
            toaster.show_toast(title="CapVisionPro",msg=f"Recording saved: {os.path.basename(final_video_path)}",icon_path="Capicon.ico",duration=3,threaded=True)
        else:
            status_bar.config(text="Recording saved (no audio)")
            write_log("Recording saved without audio")

def start_recording_countdown():
    countdown = tk.Toplevel(root)
    countdown.title("Countdown Recording")
    countdown.geometry("200x100")
    countdown.resizable(False, False)
    countdown.attributes("-topmost", True)
    label = tk.Label(countdown, font=("Arial", 32))
    label.pack(expand=True)
    for i in range(3, 0, -1):
        label.config(text=str(i))
        root.update()
        root.after(1000)
    countdown.destroy()
    StartRecording()

Capture_menu.add_command(label="Capture Image", command=lambda: CaptureImage(), accelerator="Alt-c")
Capture_menu.add_command(label="Capture & Copy to Clipboard", command=CaptureAndCopy, accelerator="Alt+Shift+C")
Capture_menu.add_separator()
Capture_menu.add_command(label="Start Recording", command=start_recording_countdown, accelerator="Alt-r")
Capture_menu.add_command(label="Stop Recording", command=lambda: StopRecording(), state="disabled", accelerator="Alt-o")
Capture_menu.add_separator()
Capture_menu.add_command(label="Start Audio Only Rec", command=start_audio_only_recording, accelerator="Alt+S")
Capture_menu.add_command(label="Stop Audio Only Rec", command=stop_audio_only_recording, accelerator="Alt+Q")
menu_bar.add_cascade(label="Capture", menu=Capture_menu)

Option_menu = tk.Menu(menu_bar, tearoff=0)

def ToggleFullscreen():
    is_full = Option1.get()
    root.attributes('-fullscreen', is_full)
    root.attributes('-topmost', is_full)
    status_bar.config(text=f"Fullscreen: {'ON' if is_full else 'OFF'}")

status_bar = None

def statusbar():
    global status_bar
    if status_bar is not None:
        status_bar.destroy()
        status_bar = None
    show_status = Option2.get()
    if show_status:
        status_bar = tk.Label(
            root,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

Option_menu.add_checkbutton(label="Fullscreen Mode", variable=Option1, command=lambda: ToggleFullscreen(), accelerator="Alt-f")
Option_menu.add_separator()
Option_menu.add_checkbutton(label="Show Status Bar", variable=Option2, command=lambda: statusbar(), accelerator="Alt-b")
menu_bar.add_cascade(label="Options", menu=Option_menu)
statusbar()

Equipment_menu = tk.Menu(menu_bar, tearoff=0)

def AudioDriver():
    audio_devices = []
    try:
        devices = sd.query_devices()
        input_devices = [d['name'] for d in devices if d['max_input_channels'] > 0]
        audio_devices = input_devices if input_devices else ["No audio device detected"]
    except Exception as e:
        audio_devices = [f"Audio Error: {str(e)}"]
    return audio_devices[0]

def CapDriver():
    camera_devices = []
    index = 0
    max_check = 5
    while index < max_check:
        cap_test = cv2.VideoCapture(index, cv2.CAP_DSHOW)
        if cap_test.isOpened():
            camera_devices.append(f"Camera {index}")
            cap_test.release()
            index += 1
        else:
            cap_test.release()
            break
    return camera_devices if camera_devices else ["No camera detected"]

Equipment_menu.add_checkbutton(label="Audio: "+AudioDriver(), variable=OptionAudio, accelerator="Alt+Shift+A")
Equipment_menu.add_checkbutton(label="Camera: "+CapDriver()[0], variable=OptionCap, accelerator="Alt+Shift+V")
menu_bar.add_cascade(label="Equipment", menu=Equipment_menu)

preview_canvas = tk.Canvas(root)
preview_canvas.pack(fill=tk.BOTH, expand=True)
cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
cap.set(cv2.CAP_PROP_FPS, 30)
cap.set(cv2.CAP_PROP_BRIGHTNESS, brightness_value)
cap.set(cv2.CAP_PROP_CONTRAST, contrast_value)

def update_camera():
    global is_recording, video_writer, cap
    if not cap.isOpened():
        messagebox.showerror("Camera Error",
        "Camera failed to start.\nPossible causes:\n• Device disconnected\n• Driver missing/outdated\n• Camera occupied by another program")
        root.after_cancel(update_camera)
        return
    ret, frame = cap.read()
    if ret:
        if flip_mode !=0:
            frame = cv2.flip(frame, flip_mode)
        if watermark_enabled:
            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            cv2.putText(frame, ts, (10,30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0,255,0), 2)
        frame = enhance_frame_info(frame)
        win_w = max(10, preview_canvas.winfo_width())
        win_h = max(10, preview_canvas.winfo_height())
        frame = cv2.resize(frame, (win_w, win_h))
        h, w = frame.shape[:2]
        aspect = w / h
        new_w = win_w
        new_h = int(new_w / aspect)
        if new_h > win_h:
            new_h = win_h
            new_w = int(new_h * aspect)
        frame = cv2.resize(frame, (new_w, new_h))
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame_rgb)
        imgtk = ImageTk.PhotoImage(image=img)
        preview_canvas.imgtk = imgtk
        preview_canvas.create_image(0, 0, image=imgtk, anchor=tk.NW)
        if is_recording and video_writer is not None and video_writer.isOpened() and not recording_paused:
            try:
                w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                w = w if w % 2 == 0 else w - 1
                h = h if h % 2 == 0 else h - 1
                f = cv2.resize(frame, (w, h))
                video_writer.write(f)
            except Exception as e:
                write_log(f"Video write error: {str(e)}")
    check_auto_segment()
    root.after(30, update_camera)

update_camera()

def on_close():
    global is_recording, video_writer, audio_stream, audio_only_recording, tray_icon
    if is_recording or audio_only_recording:
        res = messagebox.askyesno("Confirm Exit", "Recording is in progress.\nAre you sure you want to exit and stop recording?")
        if not res:
            return
        is_recording = False
        audio_only_recording = False
        if video_writer is not None:
            video_writer.release()
        if audio_stream:
            try:
                audio_stream.stop_stream()
                audio_stream.close()
            except:
                pass
    if tray_icon:
        try:
            tray_icon.stop()
        except:
            pass
    if cap is not None:
        cap.release()
    clean_temp_files()
    root.destroy()

convert_menu = tk.Menu(menu_bar, tearoff=0)
def convert_image_format():
    path = filedialog.askopenfilename(
        title="Select Image",
        filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")])
    if not path:
        return
    save_path = filedialog.asksaveasfilename(
        title="Save As",
        defaultextension=".png",
        filetypes=[("PNG", ".png"), ("JPG", ".jpg"), ("BMP", ".bmp")])
    if not save_path:
        return
    try:
        img = cv2.imread(path)
        cv2.imwrite(save_path, img)
        messagebox.showinfo("Success", "Image converted!\n" + save_path)
        write_log(f"Image converted: {path} -> {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed: {str(e)}")
        write_log(f"Image convert failed: {str(e)}")

def convert_video_format():
    path = filedialog.askopenfilename(
        title="Select Video",
        filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv")])
    if not path:
        return
    save_path = filedialog.asksaveasfilename(
        title="Save As",
        defaultextension=".mp4",
        filetypes=[("MP4", ".mp4"), ("AVI", ".avi"), ("MOV", ".mov")])
    if not save_path:
        return
    try:
        cap_vid = cv2.VideoCapture(path)
        fps = int(cap_vid.get(cv2.CAP_PROP_FPS)) or 30
        width = int(cap_vid.get(cv2.CAP_PROP_FRAME_WIDTH))
        height = int(cap_vid.get(cv2.CAP_PROP_FRAME_HEIGHT))
        width = width if width % 2 == 0 else width - 1
        height = height if height % 2 == 0 else height - 1
        ext = save_path.split(".")[-1].lower()
        if ext == "mp4":
            fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        elif ext == "avi":
            fourcc = cv2.VideoWriter_fourcc(*'XVID')
        elif ext == "mov":
            fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        else:
            fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        out = cv2.VideoWriter(save_path, fourcc, fps, (width, height))
        while True:
            ret, frame = cap_vid.read()
            if not ret:
                break
            out.write(frame)
        cap_vid.release()
        out.release()
        messagebox.showinfo("Success", "Video converted!\n" + save_path)
        write_log(f"Video converted: {path} -> {save_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Conversion failed: {str(e)}")
        write_log(f"Video convert failed: {str(e)}")

def convert_audio_format():
    path = filedialog.askopenfilename(
        title="Select Audio",
        filetypes=[("Audio files", "*.wav *.mp3 *.flac *.aac *.ogg *.m4a")])
    if not path:
        return
    save_path = filedialog.asksaveasfilename(
        title="Save As",
        defaultextension=".mp3",
        filetypes=[("MP3", ".mp3"), ("WAV", ".wav"), ("AAC", ".aac"), ("FLAC", ".flac"), ("OGG", ".ogg"), ("M4A", ".m4a")])
    if not save_path:
        return
    ffmpeg_path = "ffmpeg.exe"
    if not os.path.exists(ffmpeg_path):
        messagebox.showerror("FFmpeg Error", "ffmpeg.exe not found!\nPlease put it in the software root directory.")
        return
    cmd = [
        ffmpeg_path,
        "-y",
        "-i", path,
        "-b:a", "192k",
        save_path
    ]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
        if os.path.exists(save_path):
            messagebox.showinfo("Success", "Audio converted!\n" + save_path)
            write_log(f"Audio converted: {path} -> {save_path}")
        else:
            messagebox.showerror("Error", "Audio conversion failed")
    except Exception as e:
        messagebox.showerror("Error", f"Audio conversion failed: {str(e)}")
        write_log(f"Audio convert failed: {str(e)}")

convert_menu.add_command(label="Convert Image", command=convert_image_format, accelerator="Alt+Shift+I")
convert_menu.add_separator()
convert_menu.add_command(label="Convert Video", command=convert_video_format, accelerator="Alt+Shift+V")
convert_menu.add_separator()
convert_menu.add_command(label="Convert Audio", command=convert_audio_format, accelerator="Alt+Shift+A")
menu_bar.add_cascade(label="Convert", menu=convert_menu)

def run_script_file():
    path = filedialog.askopenfilename(title="Add Script File", filetypes=[("CapVision Script", "*.cvscript *.script"), ("All Files", "*.*")])
    if not path:
        return
    with open(path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        cmd = line.split()[0].upper()
        args = line[len(cmd):].strip()
        if cmd == "CAPTURE":
            CaptureImage()
        elif cmd == "CAPTURE_JPG":
            capture_as_jpg()
        elif cmd == "RECORD_START":
            StartRecording()
        elif cmd == "RECORD_STOP":
            StopRecording()
        elif cmd == "RECORD_PAUSE":
            global recording_paused
            recording_paused = True
        elif cmd == "RECORD_RESUME":
            recording_paused = False
        elif cmd == "AUDIO_START":
            start_audio_only_recording()
        elif cmd == "AUDIO_STOP":
            stop_audio_only_recording()
        elif cmd == "WATERMARK_ON":
            global watermark_enabled
            watermark_enabled = True
        elif cmd == "WATERMARK_OFF":
            watermark_enabled = False
        elif cmd == "FLIP_ON":
            global flip_mode
            flip_mode = 1
        elif cmd == "FLIP_OFF":
            flip_mode = 0
        elif cmd == "FULLSCREEN_ON":
            Option1.set(True)
            ToggleFullscreen()
        elif cmd == "FULLSCREEN_OFF":
            Option1.set(False)
            ToggleFullscreen()
        elif cmd == "DELAY":
            try:
                ms = int(args)
                time.sleep(ms / 1000)
            except:
                pass
        elif cmd == "LOG":
            write_log(args.strip('"'))
        elif cmd == "SHOWMSG":
            messagebox.showinfo("Script", args.strip('"'))
        elif cmd == "OPEN_IMAGE_FOLDER":
            OpenImageFolder()
        elif cmd == "OPEN_VIDEO_FOLDER":
            OpenVideoFolder()
        elif cmd == "CLEAN_TEMP":
            clean_temp_files()
        elif cmd == "EXIT":
            root.destroy()
        elif cmd == "MINIMIZE":
            root.iconify()
        elif cmd == "MAXIMIZE":
            root.state("zoomed")
        elif cmd == "NORMAL_SIZE":
            root.state("normal")
        elif cmd == "ALWAYS_ON_TOP":
            root.attributes("-topmost", True)
        elif cmd == "ALWAYS_OFF_TOP":
            root.attributes("-topmost", False)
        elif cmd == "STATUS_ON":
            Option2.set(True)
            statusbar()
        elif cmd == "STATUS_OFF":
            Option2.set(False)
            statusbar()
        elif cmd == "BRIGHTNESS":
            try:
                val = int(args)
                val = max(0, min(100, val))
                brightness_value = val
                cap.set(cv2.CAP_PROP_BRIGHTNESS, brightness_value)
            except:
                pass
        elif cmd == "CONTRAST":
            try:
                val = int(args)
                val = max(0, min(100, val))
                contrast_value = val
                cap.set(cv2.CAP_PROP_CONTRAST, contrast_value)
            except:
                pass
        elif cmd == "QUALITY_LOW":
            set_quality_low()
        elif cmd == "QUALITY_MID":
            set_quality_medium()
        elif cmd == "QUALITY_HIGH":
            set_quality_high()
        elif cmd == "RES_720P":
            set_resolution_720p()
        elif cmd == "RES_1080P":
            set_resolution_1080p()
    messagebox.showinfo("Script", "Execution completed")
    write_log(f"Script executed: {path}")

def about_script():
    info = """CVScript — Built-in interpreted scripting language for CapVisionPro
Supported commands:
CAPTURE            Take a screenshot
CAPTURE_JPG        Take a high quality JPG screenshot
RECORD_START       Start video recording
RECORD_STOP        Stop video recording
RECORD_PAUSE       Pause recording
RECORD_RESUME      Resume recording
AUDIO_START        Start audio recording
AUDIO_STOP         Stop audio recording
WATERMARK_ON       Enable watermark
WATERMARK_OFF      Disable watermark
FLIP_ON            Enable horizontal flip
FLIP_OFF           Disable horizontal flip
FULLSCREEN_ON      Enter fullscreen
FULLSCREEN_OFF     Exit fullscreen
STATUS_ON          Show status bar
STATUS_OFF         Hide status bar
DELAY ms           Delay in milliseconds
LOG "message"      Write to log
SHOWMSG "text"     Show message box
OPEN_IMAGE_FOLDER  Open image output folder
OPEN_VIDEO_FOLDER  Open video output folder
CLEAN_TEMP         Clean temporary files
BRIGHTNESS val     Set camera brightness (0-100)
CONTRAST val       Set camera contrast (0-100)
QUALITY_LOW        Set recording quality to low
QUALITY_MID        Set recording quality to medium
QUALITY_HIGH       Set recording quality to high
RES_720P           Set resolution to 1280x720
RES_1080P          Set resolution to 1920x1080
ALWAYS_ON_TOP      Keep window on top
ALWAYS_OFF_TOP     Disable always on top
MINIMIZE           Minimize window
MAXIMIZE           Maximize window
NORMAL_SIZE        Restore window size
EXIT               Exit the program
Script extensions: .cvscript .script
"""
    messagebox.showinfo("About CVScript", info)
def set_output_folder():
    global save_folder
    path = filedialog.askdirectory(title="Select Output Folder")
    if path:
        save_folder = path
        write_log(f"Output folder set to: {path}")
        messagebox.showinfo("Success", f"Output folder set to:\n{path}")

def set_resolution_720p():
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
    write_log("Resolution set to 1280x720")

def set_resolution_1080p():
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
    write_log("Resolution set to 1920x1080")

def set_resolution_default():
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
    write_log("Resolution set to default 640x480")

def set_quality_low():
    global video_bitrate
    video_bitrate = 500_000
    write_log("Recording quality: Low")

def set_quality_medium():
    global video_bitrate
    video_bitrate = 1500_000
    write_log("Recording quality: Medium")

def set_quality_high():
    global video_bitrate
    video_bitrate = 4000_000
    write_log("Recording quality: High")

def set_brightness(val):
    global brightness_value
    brightness_value = int(val)
    cap.set(cv2.CAP_PROP_BRIGHTNESS, brightness_value)

def set_contrast(val):
    global contrast_value
    contrast_value = int(val)
    cap.set(cv2.CAP_PROP_CONTRAST, contrast_value)

settings_menu = tk.Menu(menu_bar, tearoff=0)

settings_menu.add_command(label="Set Output Folder", command=set_output_folder, accelerator="Alt+O")

settings_menu.add_separator()
settings_menu.add_command(label="720p Resolution", command=set_resolution_720p, accelerator="Alt+7")
settings_menu.add_command(label="1080p Resolution", command=set_resolution_1080p, accelerator="Alt+1")
settings_menu.add_command(label="Default Resolution", command=set_resolution_default, accelerator="Alt+0")

settings_menu.add_separator()
settings_menu.add_command(label="Low Quality", command=set_quality_low, accelerator="Alt+1")
settings_menu.add_command(label="Medium Quality", command=set_quality_medium, accelerator="Alt+2")
settings_menu.add_command(label="High Quality", command=set_quality_high, accelerator="Alt+3")

menu_bar.add_cascade(label="Settings", menu=settings_menu)
script_menu = tk.Menu(menu_bar, tearoff=0)
script_menu.add_command(label="Run Script File", command=run_script_file, accelerator="Alt+Shift+R")
script_menu.add_separator()
script_menu.add_command(label="About CVScript", command=about_script, accelerator="Alt+Shift+K")
menu_bar.add_cascade(label="Script", menu=script_menu)

root.protocol("WM_DELETE_WINDOW", on_close)
if __name__ == "__main__":
    verify_on_start()
    write_log("CapVisionPro started successfully")
    try:
        root.mainloop()
    except Exception as e:
        error_log = f"\n{'='*50}\n SOFTWARE ERROR OCCURRED!\n{'='*50}\nError: {e}\n\nPlease send this error log to the developer!\n{'='*50}\n"
        write_log(f"Fatal error: {str(e)}")
        os.system("pause")