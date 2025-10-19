import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import time
import threading
import queue
import pynput
from pynput import mouse, keyboard
from datetime import datetime
import os
import sys
from collections import deque
import copy
import ctypes

class NaMouseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NaMouse - Automation Tool")
        self.root.geometry("900x700")
        
        # Get screen dimensions for boundary checking
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        # Get actual screen height including taskbar
        # This ensures we can click on taskbar items
        user32 = ctypes.windll.user32
        self.actual_screen_height = user32.GetSystemMetrics(1)  # SM_CYSCREEN
        self.actual_screen_width = user32.GetSystemMetrics(0)   # SM_CXSCREEN
        
        # Set application theme
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.is_recording = False
        self.is_playing = False
        self.is_paused = False
        self.recorded_events = []
        self.start_time = None
        self.playback_thread = None
        self.playback_stop_event = threading.Event()
        
        # Settings variables
        self.playback_speed = tk.DoubleVar(value=1.0)
        self.repeat_count = tk.IntVar(value=1)
        self.repeat_interval = tk.DoubleVar(value=0.0)
        self.current_file = None
        self.mouse_smoothing = tk.BooleanVar(value=False)  # Disabled by default for stability
        self.ignore_minimal_movements = tk.BooleanVar(value=True)
        self.minimal_movement_threshold = tk.IntVar(value=3)
        self.force_position = tk.BooleanVar(value=True)  # NEW: Force exact positioning
        
        # Hotkeys
        self.record_hotkey = tk.StringVar(value="F9")
        self.stop_hotkey = tk.StringVar(value="F10")
        self.play_hotkey = tk.StringVar(value="F11")
        self.pause_hotkey = tk.StringVar(value="F12")
        
        # Filter options
        self.record_mouse_moves = tk.BooleanVar(value=True)
        self.record_mouse_clicks = tk.BooleanVar(value=True)
        self.record_keyboard = tk.BooleanVar(value=True)
        self.record_scroll = tk.BooleanVar(value=True)
        
        # Performance options
        self.use_high_precision = tk.BooleanVar(value=False)  # Disabled by default for stability
        self.last_mouse_pos = None
        self.last_event_time = 0
        
        # Controllers
        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()
        self.mouse_listener = None
        self.keyboard_listener = None
        self.hotkey_listener = None
        
        # Statistics
        self.total_events = tk.StringVar(value="0")
        self.recording_duration = tk.StringVar(value="0.00s")
        
        # Recording state
        self.recording_start_time = None
        
        self.setup_ui()
        self.setup_global_hotkeys()
        
    def setup_ui(self):
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        
        # Menu Bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self.new_script, accelerator="Ctrl+N")
        file_menu.add_command(label="Open", command=self.open_script, accelerator="Ctrl+O")
        file_menu.add_command(label="Save", command=self.save_script, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As", command=self.save_script_as, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="Export as Python", command=self.export_as_python)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)
        
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Clear All", command=self.clear_script)
        edit_menu.add_command(label="Delete Selected", command=self.delete_selected)
        edit_menu.add_command(label="Optimize Script", command=self.optimize_script)
        
        # Main Control Frame
        control_frame = ttk.Frame(self.root, padding="10")
        control_frame.pack(fill=tk.X)
        
        # Control Buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.pack()
        
        button_style = {'width': 12}
        
        self.record_btn = ttk.Button(button_frame, text="⏺ Record", 
                                     command=self.start_recording, **button_style)
        self.record_btn.grid(row=0, column=0, padx=5, pady=5)
        
        self.stop_btn = ttk.Button(button_frame, text="⏹ Stop", 
                                   command=self.stop_action, **button_style,
                                   state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=5, pady=5)
        
        self.play_btn = ttk.Button(button_frame, text="▶ Play", 
                                   command=self.start_playback, **button_style)
        self.play_btn.grid(row=0, column=2, padx=5, pady=5)
        
        self.pause_btn = ttk.Button(button_frame, text="⏸ Pause", 
                                    command=self.pause_playback, **button_style,
                                    state=tk.DISABLED)
        self.pause_btn.grid(row=0, column=3, padx=5, pady=5)
        
        # Status Frame
        status_frame = ttk.Frame(control_frame)
        status_frame.pack(pady=10)
        
        self.status_label = ttk.Label(status_frame, text="Ready", 
                                      font=("Arial", 11, "bold"))
        self.status_label.pack()
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var,
                                           length=300, mode='determinate')
        self.progress_bar.pack(pady=5)
        
        # Statistics frame
        stats_frame = ttk.Frame(status_frame)
        stats_frame.pack()
        
        ttk.Label(stats_frame, text="Events:").grid(row=0, column=0, padx=5)
        ttk.Label(stats_frame, textvariable=self.total_events).grid(row=0, column=1, padx=5)
        ttk.Label(stats_frame, text="Duration:").grid(row=0, column=2, padx=5)
        ttk.Label(stats_frame, textvariable=self.recording_duration).grid(row=0, column=3, padx=5)
        
        # Notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Settings Tab
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="Settings")
        
        # Scrollable settings
        canvas = tk.Canvas(settings_frame, bg='white')
        scrollbar = ttk.Scrollbar(settings_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Playback Settings
        playback_group = ttk.LabelFrame(scrollable_frame, text="Playback Settings", padding="10")
        playback_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(playback_group, text="Playback Speed:").grid(row=0, column=0, sticky=tk.W, pady=5)
        speed_frame = ttk.Frame(playback_group)
        speed_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        speed_scale = ttk.Scale(speed_frame, from_=0.1, to=5.0, variable=self.playback_speed,
                               orient=tk.HORIZONTAL, length=200)
        speed_scale.pack(side=tk.LEFT)
        
        self.speed_label = ttk.Label(speed_frame, text="1.0x")
        self.speed_label.pack(side=tk.LEFT, padx=5)
        
        def update_speed_label(*args):
            self.speed_label.config(text=f"{self.playback_speed.get():.1f}x")
        
        self.playback_speed.trace('w', update_speed_label)
        
        ttk.Label(playback_group, text="Repeat Count:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(playback_group, from_=1, to=9999, textvariable=self.repeat_count,
                   width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Label(playback_group, text="(0 = infinite)").grid(row=1, column=2, sticky=tk.W, pady=5)
        
        ttk.Label(playback_group, text="Repeat Interval (s):").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(playback_group, from_=0, to=999, textvariable=self.repeat_interval,
                   increment=0.1, format="%.1f", width=10).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Performance Settings
        performance_group = ttk.LabelFrame(scrollable_frame, text="Performance Settings", padding="10")
        performance_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Checkbutton(performance_group, text="High Precision Mode (May cause issues)",
                       variable=self.use_high_precision).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(performance_group, text="Mouse Movement Smoothing (Experimental)",
                       variable=self.mouse_smoothing).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(performance_group, text="Force Exact Position (For Taskbar)",
                       variable=self.force_position).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(performance_group, text="Ignore Minimal Movements",
                       variable=self.ignore_minimal_movements).pack(anchor=tk.W, pady=2)
        
        threshold_frame = ttk.Frame(performance_group)
        threshold_frame.pack(anchor=tk.W, pady=2)
        ttk.Label(threshold_frame, text="Movement Threshold (pixels):").pack(side=tk.LEFT)
        ttk.Spinbox(threshold_frame, from_=1, to=50, textvariable=self.minimal_movement_threshold,
                   width=10).pack(side=tk.LEFT, padx=5)
        
        # Recording Filter
        filter_group = ttk.LabelFrame(scrollable_frame, text="Recording Filter", padding="10")
        filter_group.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Checkbutton(filter_group, text="Record Mouse Movements",
                       variable=self.record_mouse_moves).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(filter_group, text="Record Mouse Clicks",
                       variable=self.record_mouse_clicks).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(filter_group, text="Record Mouse Scroll",
                       variable=self.record_scroll).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(filter_group, text="Record Keyboard Input",
                       variable=self.record_keyboard).pack(anchor=tk.W, pady=2)
        
        # Hotkeys Settings
        hotkey_group = ttk.LabelFrame(scrollable_frame, text="Hotkeys", padding="10")
        hotkey_group.pack(fill=tk.X, padx=10, pady=5)
        
        hotkeys = [
            ("Record:", self.record_hotkey),
            ("Stop:", self.stop_hotkey),
            ("Play:", self.play_hotkey),
            ("Pause:", self.pause_hotkey)
        ]
        
        for i, (label, var) in enumerate(hotkeys):
            ttk.Label(hotkey_group, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
            ttk.Entry(hotkey_group, textvariable=var, width=15).grid(row=i, column=1, sticky=tk.W, pady=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Script Tab
        script_frame = ttk.Frame(notebook)
        notebook.add(script_frame, text="Script")
        
        # Script toolbar
        script_toolbar = ttk.Frame(script_frame)
        script_toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(script_toolbar, text="Clear All", command=self.clear_script).pack(side=tk.LEFT, padx=2)
        ttk.Button(script_toolbar, text="Delete Selected", command=self.delete_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(script_toolbar, text="Insert Delay", command=self.insert_delay).pack(side=tk.LEFT, padx=2)
        ttk.Button(script_toolbar, text="Optimize", command=self.optimize_script).pack(side=tk.LEFT, padx=2)
        
                # Treeview for script display
        tree_frame = ttk.Frame(script_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = ("Index", "Type", "Action", "Details", "Time")
        self.script_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", height=15)
        
        column_widths = {"Index": 50, "Type": 100, "Action": 100, "Details": 300, "Time": 100}
        for col in columns:
            self.script_tree.heading(col, text=col)
            self.script_tree.column(col, width=column_widths.get(col, 100))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.script_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.script_tree.xview)
        self.script_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.script_tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Info Tab
        info_frame = ttk.Frame(notebook)
        notebook.add(info_frame, text="Help")
        
        info_text = tk.Text(info_frame, wrap=tk.WORD, height=20, width=70, font=("Consolas", 10))
        info_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        info_content = """NaMouse - Automation Tool 
Developed by Nader Mahbub Khan
FEATURES:
• Record mouse movements, clicks, and keyboard input
• Adjustable playback speed (0.1x - 5.0x)
• Repeat actions with custom intervals
• Save and load automation scripts
• Pause/resume during playback
• Script optimization
• Export as Python code
• FIXED: Full taskbar support with forced positioning

HOTKEYS (Default):
• F9  - Start Recording
• F10 - Stop Recording/Playback
• F11 - Start Playback
• F12 - Pause/Resume Playback

HOW TO USE:
1. Click Record (F9) to start recording your actions
2. Perform the mouse/keyboard actions you want to automate
3. Click Stop (F10) to stop recording
4. Click Play (F11) to replay the recorded actions
5. Adjust speed and repeat settings as needed
"""
        info_text.insert(1.0, info_content)
        info_text.config(state=tk.DISABLED)
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-n>', lambda e: self.new_script())
        self.root.bind('<Control-o>', lambda e: self.open_script())
        self.root.bind('<Control-s>', lambda e: self.save_script())
        self.root.bind('<Control-Shift-S>', lambda e: self.save_script_as())
    
    def setup_global_hotkeys(self):
        """Setup keyboard listener for global hotkeys"""
        def on_press(key):
            try:
                key_name = None
                if hasattr(key, 'name'):
                    key_name = key.name.upper()
                elif hasattr(key, 'char') and key.char:
                    key_name = key.char.upper()
                
                if not key_name:
                    return
                
                # Check hotkeys
                if key_name == self.record_hotkey.get().upper():
                    if not self.is_recording and not self.is_playing:
                        self.root.after(0, self.start_recording)
                elif key_name == self.stop_hotkey.get().upper():
                    self.root.after(0, self.stop_action)
                elif key_name == self.play_hotkey.get().upper():
                    if not self.is_recording and not self.is_playing:
                        self.root.after(0, self.start_playback)
                elif key_name == self.pause_hotkey.get().upper():
                    if self.is_playing:
                        self.root.after(0, self.pause_playback)
                        
            except Exception:
                pass
        
        self.hotkey_listener = keyboard.Listener(on_press=on_press)
        self.hotkey_listener.start()
    
    def set_mouse_position_forced(self, x, y):
        """Force mouse to exact position with multiple attempts"""
        target_x = int(x)
        target_y = int(y)
        
        if self.force_position.get():
            # Multiple attempts to ensure position is set
            for attempt in range(3):
                self.mouse_controller.position = (target_x, target_y)
                time.sleep(0.005)  # Small delay between attempts
                
                # Verify position
                current_pos = self.mouse_controller.position
                if abs(current_pos[0] - target_x) <= 1 and abs(current_pos[1] - target_y) <= 1:
                    break
        else:
            self.mouse_controller.position = (target_x, target_y)
    
    def validate_mouse_position(self, x, y):
        """Ensure mouse position is within screen boundaries including taskbar"""
        # Allow full screen height including taskbar
        x = max(0, min(x, self.actual_screen_width - 1))
        y = max(0, min(y, self.actual_screen_height - 1))
        return x, y
    
    def validate_mouse_position_for_playback(self, x, y):
        """Special validation for playback that preserves exact positions"""
        # During playback, we want to preserve the exact recorded position
        # No modification - just ensure it's an integer
        return int(x), int(y)
    
    def start_recording(self):
        """Start recording with improved event handling"""
        if self.is_playing:
            messagebox.showwarning("Warning", "Cannot record while playing!")
            return
        
        self.is_recording = True
        self.recorded_events = []
        self.recording_start_time = time.time()
        self.last_mouse_pos = None
        self.last_event_time = 0
        
        self.record_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.play_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Recording...", foreground="red")
        
        # Start recording timer update
        self.update_recording_time()
        
        # Start listeners
        try:
            if self.record_mouse_clicks.get() or self.record_mouse_moves.get() or self.record_scroll.get():
                mouse_callbacks = {}
                if self.record_mouse_moves.get():
                    mouse_callbacks['on_move'] = self.on_mouse_move
                if self.record_mouse_clicks.get():
                    mouse_callbacks['on_click'] = self.on_mouse_click
                if self.record_scroll.get():
                    mouse_callbacks['on_scroll'] = self.on_mouse_scroll
                
                self.mouse_listener = mouse.Listener(**mouse_callbacks)
                self.mouse_listener.start()
            
            if self.record_keyboard.get():
                self.keyboard_listener = keyboard.Listener(
                    on_press=self.on_key_press,
                    on_release=self.on_key_release
                )
                self.keyboard_listener.start()
        except Exception as e:
            self.stop_recording()
            messagebox.showerror("Error", f"Failed to start recording: {str(e)}")
    
    def update_recording_time(self):
        """Update recording duration display"""
        if self.is_recording:
            duration = time.time() - self.recording_start_time
            self.recording_duration.set(f"{duration:.2f}s")
            self.total_events.set(str(len(self.recorded_events)))
            self.root.after(100, self.update_recording_time)
    
    def stop_action(self):
        """Stop recording or playback"""
        if self.is_recording:
            self.stop_recording()
        elif self.is_playing:
            self.stop_playback()
    
    def stop_recording(self):
        """Stop recording with cleanup"""
        self.is_recording = False
        
        # Stop listeners
        try:
            if self.mouse_listener:
                self.mouse_listener.stop()
                self.mouse_listener = None
            
            if self.keyboard_listener:
                self.keyboard_listener.stop()
                self.keyboard_listener = None
        except:
            pass
        
        # Update UI
        self.record_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.play_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Ready", foreground="black")
        
        self.update_script_display()
        self.total_events.set(str(len(self.recorded_events)))
        
        if self.recorded_events:
            duration = self.recorded_events[-1]['time']
            self.recording_duration.set(f"{duration:.2f}s")
    
    def stop_playback(self):
        """Stop playback immediately"""
        self.is_playing = False
        self.is_paused = False
        self.playback_stop_event.set()
        
        self.record_btn.config(state=tk.NORMAL)
        self.play_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Playback Stopped", foreground="black")
        self.progress_var.set(0)
    
    def start_playback(self):
        """Start playback with enhanced precision"""
        if not self.recorded_events:
            messagebox.showinfo("Info", "No events recorded!")
            return
        
        if self.is_recording:
            messagebox.showwarning("Warning", "Cannot play while recording!")
            return
        
        self.is_playing = True
        self.is_paused = False
        self.playback_stop_event.clear()
        
        self.record_btn.config(state=tk.DISABLED)
        self.play_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Playing...", foreground="green")
        
        # Start playback thread
        self.playback_thread = threading.Thread(target=self.playback_events_stable)
        self.playback_thread.daemon = True
        self.playback_thread.start()
    
    def pause_playback(self):
        """Pause or resume playback"""
        if self.is_playing:
            self.is_paused = not self.is_paused
            if self.is_paused:
                self.pause_btn.config(text="▶ Resume")
                self.status_label.config(text="Paused", foreground="orange")
            else:
                self.pause_btn.config(text="⏸ Pause")
                self.status_label.config(text="Playing...", foreground="green")
    
    def playback_events_stable(self):
        """Stable playback with proper timing and taskbar support"""
        try:
            repeat_count = self.repeat_count.get()
            if repeat_count == 0:
                repeat_count = 9999  # Large number instead of infinity
            
            repeat_interval = self.repeat_interval.get()
            speed = self.playback_speed.get()
            
            total_duration = self.recorded_events[-1]['time'] if self.recorded_events else 0
            
            for repeat in range(repeat_count):
                if self.playback_stop_event.is_set():
                    break
                
                # Wait between repeats
                if repeat > 0 and repeat_interval > 0:
                    wait_start = time.time()
                    while time.time() - wait_start < repeat_interval:
                        if self.playback_stop_event.is_set():
                            break
                        time.sleep(0.1)
                
                # Play events
                start_playback_time = time.time()
                
                for i, event in enumerate(self.recorded_events):
                    if self.playback_stop_event.is_set():
                        break
                    
                    # Handle pause
                    while self.is_paused and not self.playback_stop_event.is_set():
                        time.sleep(0.1)
                    
                    # Calculate timing
                    target_time = event['time'] / speed
                    elapsed = time.time() - start_playback_time
                    wait_time = target_time - elapsed
                    
                    # Wait if needed
                    if wait_time > 0:
                        time.sleep(wait_time)
                    
                    # Update progress
                    if total_duration > 0:
                        progress = (event['time'] / total_duration) * 100
                        self.root.after(0, lambda p=progress: self.progress_var.set(p))
                    
                    # Execute event
                    self.execute_event_safe(event)
            
            self.is_playing = False
            self.root.after(0, self.playback_finished)
            
        except Exception as e:
            print(f"Playback error: {e}")
            self.is_playing = False
            self.root.after(0, self.playback_finished)
    
    def execute_event_safe(self, event):
        """Execute a single event with enhanced taskbar support"""
        try:
            event_type = event['type']
            
            if event_type == 'mouse_move':
                # Use exact position for playback
                x, y = self.validate_mouse_position_for_playback(event['x'], event['y'])
                
                if self.mouse_smoothing.get():
                    # Simple smoothing
                    current_pos = self.mouse_controller.position
                    steps = 3
                    for i in range(1, steps + 1):
                        if self.playback_stop_event.is_set():
                            break
                        interp_x = current_pos[0] + (x - current_pos[0]) * (i / steps)
                        interp_y = current_pos[1] + (y - current_pos[1]) * (i / steps)
                        self.set_mouse_position_forced(int(interp_x), int(interp_y))
                        time.sleep(0.003)
                else:
                    self.set_mouse_position_forced(x, y)
                    
            elif event_type == 'mouse_click':
                # Use exact position for clicks (critical for taskbar)
                x, y = self.validate_mouse_position_for_playback(event['x'], event['y'])
                button = mouse.Button.left if event['button'] == 'left' else mouse.Button.right
                
                # Move to exact position with forced positioning
                self.set_mouse_position_forced(x, y)
                time.sleep(0.03)  # Extended delay for taskbar reliability
                
                # Double-check position before clicking
                self.set_mouse_position_forced(x, y)
                time.sleep(0.01)
                
                # Perform the click
                if event['pressed']:
                    self.mouse_controller.press(button)
                    time.sleep(0.01)  # Small delay after press
                else:
                    self.mouse_controller.release(button)
                    time.sleep(0.01)  # Small delay after release
                    
            elif event_type == 'mouse_scroll':
                x, y = self.validate_mouse_position_for_playback(event['x'], event['y'])
                self.set_mouse_position_forced(x, y)
                time.sleep(0.02)
                self.mouse_controller.scroll(event['dx'], event['dy'])
                
            elif event_type == 'key_press':
                key = event['key']
                if key:
                    try:
                        if len(key) == 1:
                            self.keyboard_controller.press(key)
                        else:
                            key_obj = getattr(keyboard.Key, key, None)
                            if key_obj:
                                self.keyboard_controller.press(key_obj)
                    except:
                        pass
                        
            elif event_type == 'key_release':
                key = event['key']
                if key:
                    try:
                        if len(key) == 1:
                            self.keyboard_controller.release(key)
                        else:
                            key_obj = getattr(keyboard.Key, key, None)
                            if key_obj:
                                self.keyboard_controller.release(key_obj)
                    except:
                        pass
                        
            elif event_type == 'delay':
                time.sleep(event['duration'] / self.playback_speed.get())
                
        except Exception as e:
            print(f"Event execution error: {e}")
    
    def playback_finished(self):
        """Clean up after playback finishes"""
        self.record_btn.config(state=tk.NORMAL)
        self.play_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED, text="⏸ Pause")
        self.status_label.config(text="Ready", foreground="black")
        self.progress_var.set(0)

    def on_mouse_move(self, x, y):
        """Record mouse movement - records exact position"""
        if self.is_recording and not self.is_playing:
            # Record exact position without modification
            # Check if movement is significant
            if self.ignore_minimal_movements.get() and self.last_mouse_pos:
                dx = abs(x - self.last_mouse_pos[0])
                dy = abs(y - self.last_mouse_pos[1])
                if dx < self.minimal_movement_threshold.get() and dy < self.minimal_movement_threshold.get():
                    return
            
            # Limit event frequency
            current_time = time.time()
            if current_time - self.last_event_time < 0.01:  # Max 100 events per second
                return
            
            self.last_event_time = current_time
            self.last_mouse_pos = (x, y)
            
            event = {
                'type': 'mouse_move',
                'time': current_time - self.recording_start_time,
                'x': x,
                'y': y
            }
            self.recorded_events.append(event)
    
    def on_mouse_click(self, x, y, button, pressed):
        """Record mouse click with exact position"""
        if self.is_recording and not self.is_playing:
            # Record exact position - no validation during recording
            event = {
                'type': 'mouse_click',
                'time': time.time() - self.recording_start_time,
                'x': x,
                'y': y,
                'button': 'left' if button == mouse.Button.left else 'right',
                'pressed': pressed
            }
            self.recorded_events.append(event)
    
    def on_mouse_scroll(self, x, y, dx, dy):
        """Record mouse scroll with exact position"""
        if self.is_recording and not self.is_playing:
            event = {
                'type': 'mouse_scroll',
                'time': time.time() - self.recording_start_time,
                'x': x,
                'y': y,
                'dx': dx,
                'dy': dy
            }
            self.recorded_events.append(event)
    
    def on_key_press(self, key):
        """Record key press with filtering"""
        if self.is_recording and not self.is_playing:
            try:
                key_name = None
                if hasattr(key, 'char') and key.char:
                    key_name = key.char
                elif hasattr(key, 'name'):
                    key_name = key.name
                
                if not key_name:
                    return
                
                # Don't record hotkeys
                hotkeys = [self.record_hotkey.get().upper(), self.stop_hotkey.get().upper(),
                          self.play_hotkey.get().upper(), self.pause_hotkey.get().upper()]
                if key_name.upper() in hotkeys:
                    return
                
                event = {
                    'type': 'key_press',
                    'time': time.time() - self.recording_start_time,
                    'key': key_name
                }
                self.recorded_events.append(event)
            except:
                pass
    
    def on_key_release(self, key):
        """Record key release"""
        if self.is_recording and not self.is_playing:
            try:
                key_name = None
                if hasattr(key, 'char') and key.char:
                    key_name = key.char
                elif hasattr(key, 'name'):
                    key_name = key.name
                
                if not key_name:
                    return
                
                # Don't record hotkeys
                hotkeys = [self.record_hotkey.get().upper(), self.stop_hotkey.get().upper(),
                          self.play_hotkey.get().upper(), self.pause_hotkey.get().upper()]
                if key_name.upper() in hotkeys:
                    return
                
                event = {
                    'type': 'key_release',
                    'time': time.time() - self.recording_start_time,
                    'key': key_name
                }
                self.recorded_events.append(event)
            except:
                pass
    
    def update_script_display(self):
        """Update the script display"""
        # Clear existing items
        for item in self.script_tree.get_children():
            self.script_tree.delete(item)
        
        # Add events to tree
        for i, event in enumerate(self.recorded_events):
            event_type = event['type'].replace('_', ' ').title()
            action = ""
            details = ""
            
            if event['type'] == 'mouse_move':
                action = "Move"
                details = f"Position: ({event['x']}, {event['y']})"
            elif event['type'] == 'mouse_click':
                action = "Press" if event['pressed'] else "Release"
                details = f"{event['button'].title()} button at ({event['x']}, {event['y']})"
            elif event['type'] == 'mouse_scroll':
                action = "Scroll"
                details = f"Delta: ({event['dx']}, {event['dy']}) at ({event['x']}, {event['y']})"
            elif event['type'] in ['key_press', 'key_release']:
                action = "Press" if event['type'] == 'key_press' else "Release"
                details = f"Key: {event.get('key', 'Unknown')}"
            elif event['type'] == 'delay':
                action = "Wait"
                details = f"Duration: {event['duration']:.3f}s"
            
            time_str = f"{event['time']:.3f}s"
            
            # Add color coding
            tags = []
            if 'mouse' in event['type']:
                tags.append('mouse')
            elif 'key' in event['type']:
                tags.append('keyboard')
            elif event['type'] == 'delay':
                tags.append('delay')
            
            self.script_tree.insert("", "end", values=(i+1, event_type, action, details, time_str), tags=tags)
        
        # Configure tags
        self.script_tree.tag_configure('mouse', foreground='blue')
        self.script_tree.tag_configure('keyboard', foreground='green')
        self.script_tree.tag_configure('delay', foreground='orange')
    
    def insert_delay(self):
        """Insert a custom delay"""
        selected = self.script_tree.selection()
        if not selected:
            index = len(self.recorded_events)
        else:
            index = self.script_tree.index(selected[0]) + 1
        
        # Create delay dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Insert Delay")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Delay Duration (seconds):").pack(pady=10)
        
        delay_var = tk.DoubleVar(value=1.0)
        ttk.Spinbox(dialog, from_=0.1, to=60, textvariable=delay_var,
                   increment=0.1, format="%.1f", width=10).pack(pady=5)
        
        def insert():
            delay_event = {
                'type': 'delay',
                'time': self.recorded_events[index-1]['time'] if index > 0 and self.recorded_events else 0,
                'duration': delay_var.get()
            }
            
            self.recorded_events.insert(index, delay_event)
            
            # Adjust subsequent event times
            for i in range(index + 1, len(self.recorded_events)):
                if self.recorded_events[i]['type'] != 'delay':
                    self.recorded_events[i]['time'] += delay_var.get()
            
            self.update_script_display()
            dialog.destroy()
        
        ttk.Button(dialog, text="Insert", command=insert).pack(pady=10)
        ttk.Button(dialog, text="Cancel", command=dialog.destroy).pack()
    
    def optimize_script(self):
        """Optimize script by removing redundant events"""
        if not self.recorded_events:
            messagebox.showinfo("Info", "No events to optimize")
            return
        
        original_count = len(self.recorded_events)
        optimized = []
        last_mouse_move = None
        
        for event in self.recorded_events:
            # Skip redundant mouse moves
            if event['type'] == 'mouse_move':
                if last_mouse_move and event['time'] - last_mouse_move['time'] < 0.02:
                    # Update the last mouse move
                    last_mouse_move['x'] = event['x']
                    last_mouse_move['y'] = event['y']
                    last_mouse_move['time'] = event['time']
                else:
                    optimized.append(event)
                    last_mouse_move = event
            else:
                optimized.append(event)
                last_mouse_move = None
        
        self.recorded_events = optimized
        removed = original_count - len(optimized)
        
        self.update_script_display()
        self.total_events.set(str(len(self.recorded_events)))
        
        messagebox.showinfo("Optimization Complete", 
                           f"Removed {removed} redundant events\n"
                           f"Original: {original_count} events\n"
                           f"Optimized: {len(optimized)} events")
    
    def export_as_python(self):
        """Export the script as a standalone Python file"""
        if not self.recorded_events:
            messagebox.showinfo("Info", "No events to export")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export as Python",
            defaultextension=".py",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(self.generate_python_code())
                messagebox.showinfo("Success", f"Exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")
    
    def generate_python_code(self):
        """Generate standalone Python code"""
        code = '''#!/usr/bin/env python3
"""
Generated by NaMouse - Automation Tool
Date: {date}
"""

import time
from pynput import mouse, keyboard

# Initialize controllers
mouse_controller = mouse.Controller()
keyboard_controller = keyboard.Controller()

def set_mouse_position_forced(x, y):
    """Force mouse to exact position"""
    for _ in range(3):
        mouse_controller.position = (int(x), int(y))
        time.sleep(0.005)

def run_automation():
    """Execute the recorded automation"""
    print("Starting automation in 3 seconds...")
    time.sleep(3)
    
    start_time = time.time()
    
'''.format(date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        for event in self.recorded_events:
            if event['type'] == 'mouse_move':
                code += f"    # Move mouse\n"
                code += f"    time.sleep(max(0, {event['time']:.3f} - (time.time() - start_time)))\n"
                code += f"    set_mouse_position_forced({event['x']}, {event['y']})\n\n"
                
            elif event['type'] == 'mouse_click':
                button = 'mouse.Button.left' if event['button'] == 'left' else 'mouse.Button.right'
                action = 'press' if event['pressed'] else 'release'
                code += f"    # {action.title()} {event['button']} button\n"
                code += f"    time.sleep(max(0, {event['time']:.3f} - (time.time() - start_time)))\n"
                code += f"    set_mouse_position_forced({event['x']}, {event['y']})\n"
                code += f"    time.sleep(0.03)\n"
                code += f"    set_mouse_position_forced({event['x']}, {event['y']})\n"
                code += f"    time.sleep(0.01)\n"
                code += f"    mouse_controller.{action}({button})\n"
                code += f"    time.sleep(0.01)\n"
                
            elif event['type'] == 'mouse_scroll':
                code += f"    # Scroll\n"
                code += f"    time.sleep(max(0, {event['time']:.3f} - (time.time() - start_time)))\n"
                code += f"    set_mouse_position_forced({event['x']}, {event['y']})\n"
                code += f"    time.sleep(0.02)\n"
                code += f"    mouse_controller.scroll({event['dx']}, {event['dy']})\n\n"
                
            elif event['type'] == 'key_press':
                key = event.get('key', '')
                if len(key) == 1:
                    key_code = f"'{key}'"
                else:
                    key_code = f"keyboard.Key.{key}"
                code += f"    # Press key: {key}\n"
                code += f"    time.sleep(max(0, {event['time']:.3f} - (time.time() - start_time)))\n"
                code += f"    try:\n"
                code += f"        keyboard_controller.press({key_code})\n"
                code += f"    except: pass\n\n"
                
            elif event['type'] == 'key_release':
                key = event.get('key', '')
                if len(key) == 1:
                    key_code = f"'{key}'"
                else:
                    key_code = f"keyboard.Key.{key}"
                code += f"    # Release key: {key}\n"
                code += f"    time.sleep(max(0, {event['time']:.3f} - (time.time() - start_time)))\n"
                code += f"    try:\n"
                code += f"        keyboard_controller.release({key_code})\n"
                code += f"    except: pass\n\n"
                
            elif event['type'] == 'delay':
                code += f"    # Custom delay\n"
                code += f"    time.sleep({event['duration']:.3f})\n\n"
        
        code += '''    print("Automation completed!")

if __name__ == "__main__":
    try:
        run_automation()
    except KeyboardInterrupt:
        print("\\nAutomation interrupted")
    except Exception as e:
        print(f"Error: {e}")
'''
        return code
    
    def clear_script(self):
        """Clear all recorded events"""
        if self.recorded_events:
            if messagebox.askyesno("Confirm", "Clear all recorded events?"):
                self.recorded_events = []
                self.update_script_display()
                self.total_events.set("0")
                self.recording_duration.set("0.00s")
    
    def delete_selected(self):
        """Delete selected events"""
        selected = self.script_tree.selection()
        if selected:
            indices = sorted([self.script_tree.index(item) for item in selected], reverse=True)
            
            for i in indices:
                if i < len(self.recorded_events):
                    del self.recorded_events[i]
            
            self.update_script_display()
            self.total_events.set(str(len(self.recorded_events)))
    
    def new_script(self):
        """Create a new script"""
        if self.recorded_events:
            if messagebox.askyesno("Confirm", "Create new script? Current events will be lost if not saved."):
                self.recorded_events = []
                self.current_file = None
                self.update_script_display()
                self.root.title("NaMouse - Automation Tool")
                self.total_events.set("0")
                self.recording_duration.set("0.00s")
    
    def open_script(self):
        """Open a saved script file"""
        filename = filedialog.askopenfilename(
            title="Open Script",
            filetypes=[("NaMouse Script", "*.nam"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            try:
                with open(filename, 'r') as f:
                    data = json.load(f)
                    
                # Handle both old and new format
                if isinstance(data, list):
                    self.recorded_events = data
                else:
                    self.recorded_events = data.get('events', [])
                    # Load settings if available
                    if 'settings' in data:
                        settings = data['settings']
                        self.playback_speed.set(settings.get('playback_speed', 1.0))
                        self.repeat_count.set(settings.get('repeat_count', 1))
                        self.repeat_interval.set(settings.get('repeat_interval', 0))
                        self.mouse_smoothing.set(settings.get('mouse_smoothing', False))
                        self.use_high_precision.set(settings.get('use_high_precision', False))
                        self.force_position.set(settings.get('force_position', True))
                
                self.current_file = filename
                self.update_script_display()
                self.root.title(f"NaMouse - {os.path.basename(filename)}")
                
                if self.recorded_events:
                    self.total_events.set(str(len(self.recorded_events)))
                    self.recording_duration.set(f"{self.recorded_events[-1]['time']:.2f}s")
                
                messagebox.showinfo("Success", "Script loaded successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load script: {str(e)}")
    
    def save_script(self):
        """Save the current script"""
        if self.current_file:
            self.save_to_file(self.current_file)
        else:
            self.save_script_as()
    
    def save_script_as(self):
        """Save the script with a new name"""
        filename = filedialog.asksaveasfilename(
            title="Save Script",
            defaultextension=".nam",
            filetypes=[("NaMouse Script", "*.nam"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.save_to_file(filename)
            self.current_file = filename
            self.root.title(f"NaMouse - {os.path.basename(filename)}")
    
    def save_to_file(self, filename):
        """Save script to file"""
        try:
            data = {
                'version': '2.3',
                'events': self.recorded_events,
                'settings': {
                    'playback_speed': self.playback_speed.get(),
                    'repeat_count': self.repeat_count.get(),
                    'repeat_interval': self.repeat_interval.get(),
                    'mouse_smoothing': self.mouse_smoothing.get(),
                    'use_high_precision': self.use_high_precision.get(),
                    'force_position': self.force_position.get()
                },
                'metadata': {
                    'created': datetime.now().isoformat(),
                    'total_events': len(self.recorded_events),
                    'duration': self.recorded_events[-1]['time'] if self.recorded_events else 0,
                    'screen_width': self.actual_screen_width,
                    'screen_height': self.actual_screen_height
                }
            }
            
            with open(filename, 'w') as f:
                json.dump(data, f, indent=2)
            
            messagebox.showinfo("Success", "Script saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save script: {str(e)}")
    
    def on_closing(self):
        """Handle application closing"""
        if self.recorded_events and not self.current_file:
            if messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Do you want to save before closing?"):
                self.save_script_as()
        
        # Clean up listeners
        try:
            if self.mouse_listener:
                self.mouse_listener.stop()
            if self.keyboard_listener:
                self.keyboard_listener.stop()
            if self.hotkey_listener:
                self.hotkey_listener.stop()
        except:
            pass
        
        self.root.destroy()

def main():
    root = tk.Tk()
    app = NaMouseApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
