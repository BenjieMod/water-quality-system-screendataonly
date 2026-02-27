import random
import time
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import datetime
import os
import platform
import asyncio
import subprocess
import sys
from playwright.async_api import Playwright, async_playwright, expect
from PIL import Image, ImageDraw
import pystray

# --- Playwright Path Setup for PyInstaller ---
if getattr(sys, 'frozen', False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(sys._MEIPASS, "playwright")

# --- Tkinter Setup ---
root = tk.Tk()
root.geometry("950x680")
root.update_idletasks()

# --- Tkinter Variables ---
headless_mode = tk.BooleanVar(value=False)
mode_var = tk.StringVar(value="sixteen")
flood_mode_var = tk.StringVar(value="Select Mode")

# --- Shift Definitions ---
shift_options = [
    "Select Shift",
    "STRAIGHT SHIFT (5:00 PM - 7:00 AM)",
    "SECOND SHIFT (8:00 AM - 4:00 PM)",
    "THIRD SHIFT (4:00 PM - 11:00 PM)"
]
shift_var = tk.StringVar(value="Select Shift")

shift_definitions = {
    "SECOND SHIFT (8:00 AM - 4:00 PM)": {
        "rows": 9,
        "start_hour": 8,
        "end_hour": 16
    },
    "STRAIGHT SHIFT (5:00 PM - 7:00 AM)": {
        "rows": 15,
        "start_hour": 17,
        "end_hour": 31
    },
    "THIRD SHIFT (4:00 PM - 11:00 PM)": {
        "rows": 8,
        "start_hour": 16,
        "end_hour": 23
    }
}

# --- Flood Mode Definitions ---
flood_options = ["Select Mode", "No Flood", "Flood"]

# --- THEME DICTIONARIES ---
LIGHT_THEME = {
    "BG_COLOR": "#c0d7e0",
    "FG_COLOR": "#064273",
    "ACCENT_COLOR": "#38b6ff",
    "SUCCESS_COLOR": "#1ab441",
    "WARNING_COLOR": "#ffb300",
    "DANGER_COLOR": "#e74c3c",
    "SECONDARY_COLOR": "#c0d7e0",
    "ENTRY_BG": "#ffffff",
    "BUTTON_HOVER": "#add8e6"
}

NIGHT_THEME = {
    "BG_COLOR": "#181a1b",
    "FG_COLOR": "#e0e0e0",
    "ACCENT_COLOR": "#4a90e2",
    "SUCCESS_COLOR": "#22d279",
    "WARNING_COLOR": "#ffb74d",
    "DANGER_COLOR": "#ff4b5c",
    "SECONDARY_COLOR": "#23272e",
    "ENTRY_BG": "#23272e",
    "BUTTON_HOVER": "#323a45"
}

current_theme = "light"

def apply_theme(theme):
    global BG_COLOR, FG_COLOR, ACCENT_COLOR, SUCCESS_COLOR, WARNING_COLOR, DANGER_COLOR, SECONDARY_COLOR, ENTRY_BG, BUTTON_HOVER
    BG_COLOR = theme["BG_COLOR"]
    FG_COLOR = theme["FG_COLOR"]
    ACCENT_COLOR = theme["ACCENT_COLOR"]
    SUCCESS_COLOR = theme["SUCCESS_COLOR"]
    WARNING_COLOR = theme["WARNING_COLOR"]
    DANGER_COLOR = theme["DANGER_COLOR"]
    SECONDARY_COLOR = theme["SECONDARY_COLOR"]
    ENTRY_BG = theme["ENTRY_BG"]
    BUTTON_HOVER = theme["BUTTON_HOVER"]

apply_theme(LIGHT_THEME)

# --- GLOBALS ---
values = []
scheduled_times = []
auto_exit_time = None
pw_window = None
running_event = threading.Event()
start_button = stop_button = None
status_label = None
notification_labels = []
alarm_enabled = False
alarm_lock = threading.Lock()
toggle_btn = None
tray_icon = None

async_loop = None
async_thread = None

# NEW: Execution control and Failure Tracking
execution_lock = threading.Lock()
currently_executing = set()
consecutive_failures = 0
consecutive_failures_lock = threading.Lock()

# --- VALIDATION FUNCTIONS ---
def validate_time_format(time_str):
    try:
        time_str = time_str.strip()
        parts = time_str.split(":")
        if len(parts) == 2:
            parts.append("00")
        if len(parts) != 3:
            return False
        h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
        return 0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59
    except:
        return False

def auto_correct_time(time_str):
    try:
        time_str = time_str.strip()
        parts = time_str.split(":")
        if len(parts) == 2:
            return f"{time_str}:00"
        return time_str
    except:
        return time_str

def validate_value(value_str):
    try:
        val = float(value_str)
        return val >= 0.0
    except:
        return False

def on_value_change(event, entry, row_idx):
    value = entry.get()
    if value:
        if validate_value(value):
            try:
                entry.configure(style="TEntry")
                if row_idx < len(notification_labels):
                    current_notif = notification_labels[row_idx].cget("text")
                    if "Invalid Value" in current_notif:
                        update_row_notification(row_idx, "Idle", color=FG_COLOR)
            except:
                pass
        else:
            try:
                style = ttk.Style()
                style.configure("Warning.TEntry", fieldbackground="#fff3cd", foreground="#856404")
                entry.configure(style="Warning.TEntry")
            except:
                pass

def on_time_focus_out(event, entry):
    time_val = entry.get()
    if time_val:
        corrected = auto_correct_time(time_val)
        if corrected != time_val:
            entry.delete(0, tk.END)
            entry.insert(0, corrected)
        
        for idx, time_entry in enumerate(scheduled_times):
            if time_entry == entry:
                if idx < len(notification_labels):
                    current_notif = notification_labels[idx].cget("text")
                    if "Invalid Time" in current_notif or "Invalid Value" in current_notif:
                        if validate_time_format(corrected):
                            update_row_notification(idx, "Idle", color=FG_COLOR)
                break

def on_time_change(event, entry, row_idx):
    time_val = entry.get()
    if not time_val:
        return
    if len(time_val) >= 5:
        corrected = auto_correct_time(time_val)
        if corrected != time_val:
            entry.delete(0, tk.END)
            entry.insert(0, corrected)
            time_val = corrected
    if len(time_val) >= 8:
        if validate_time_format(time_val):
            try:
                entry.configure(style="TEntry")
                if row_idx < len(notification_labels):
                    current_notif = notification_labels[row_idx].cget("text")
                    if "Invalid Time" in current_notif:
                        update_row_notification(row_idx, "Idle", color=FG_COLOR)
            except:
                pass
        else:
            try:
                style = ttk.Style()
                style.configure("Warning.TEntry", fieldbackground="#fff3cd", foreground="#856404")
                entry.configure(style="Warning.TEntry")
            except:
                pass

# --- UI Helper Functions ---
def on_enter(e, button, color):
    button.config(bg=BUTTON_HOVER)

def on_leave(e, button, color):
    button.config(bg=color)

def update_row_notification(row, message, color=None):
    if 0 <= row < len(notification_labels):
        def inner():
            if 0 <= row < len(notification_labels):
                label = notification_labels[row]
                label.config(text=message, fg=color or FG_COLOR)
        try:
            root.after(0, inner)
        except Exception:
            pass

def trigger_popup_on_focus(latest, difference, scheduled_time, success=True, message="Values sent successfully"):
    root.after(0, lambda: show_dam_difference_popup(latest, difference, scheduled_time, success, message))

def show_dam_difference_popup(latest_level, difference, scheduled_time, success=True, message="Values sent successfully"):
    try:
        hour = int(scheduled_time.split(":")[0])
        minute = int(scheduled_time.split(":")[1])
        if minute >= 30:
            hour = (hour + 1) % 24
        scheduled_str = f"{hour:02d}:00"
    except Exception:
        scheduled_str = scheduled_time

    popup = tk.Toplevel(root)
    popup.title("Dam Level Update")
    popup.configure(bg=SECONDARY_COLOR)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h = sw // 2, sh // 2
    x, y = (sw - w) // 2, (sh - h) // 2
    popup.geometry(f"{w}x{h}+{x}+{y}")
    popup.transient(root)

    msg1 = f"Dam level at ({scheduled_str}) is"
    msg2 = f"Difference from previous is"
    msg3 = message

    latest_color = ACCENT_COLOR
    diff_color = SUCCESS_COLOR if difference >= 0 else DANGER_COLOR
    msg3_color = SUCCESS_COLOR if success else DANGER_COLOR

    frame = tk.Frame(popup, bg=SECONDARY_COLOR)
    frame.pack(expand=True, fill="both")

    tk.Label(frame, text=msg1, font=("Arial", 24, "bold"), fg=FG_COLOR, bg=SECONDARY_COLOR, justify="center").pack(pady=(40,0))
    tk.Label(frame, text=f"{latest_level:.2f}", font=("Arial", 54, "bold"), fg=latest_color, bg=SECONDARY_COLOR, justify="center").pack(pady=(0,5))
    tk.Label(frame, text=msg2, font=("Arial", 18, "bold"), fg=FG_COLOR, bg=SECONDARY_COLOR, justify="center").pack(pady=(10,0))
    tk.Label(frame, text=f"{difference:+.2f}", font=("Arial", 36, "bold"), fg=diff_color, bg=SECONDARY_COLOR, justify="center").pack(pady=(0,10))
    tk.Label(frame, text=msg3, font=("Arial", 30, "bold"), fg=msg3_color, bg=SECONDARY_COLOR).pack(pady=(0,40))
    tk.Button(frame, text="CLOSE", command=popup.destroy, font=("Arial", 14, "bold"), bg=ACCENT_COLOR, fg="white", activebackground=BUTTON_HOVER).pack()

def show_error_messagebox(title, message):
    root.after(0, lambda: messagebox.showerror(title, message))

# Global variable to store the reminder label
reminder_label = None
reminder_animation_after_id = None

def create_animated_reminder_label(parent_frame):
    global reminder_label
    reminder_label = tk.Label(
        parent_frame, 
        text="", 
        font=("Arial", 14, "bold"), 
        bg=SECONDARY_COLOR, 
        fg=DANGER_COLOR,
        wraplength=400,
        justify="center"
    )
    reminder_label.pack(side="top", pady=(5, 0), fill="x")
    return reminder_label

def animate_reminder_text(text, blink_count=0, max_blinks=6):
    global reminder_label, reminder_animation_after_id
    if not reminder_label or not reminder_label.winfo_exists():
        return
    if reminder_animation_after_id:
        root.after_cancel(reminder_animation_after_id)
    if blink_count < max_blinks:
        if blink_count % 2 == 0:
            reminder_label.config(text=f"🚨 {text} 🚨", fg=DANGER_COLOR, font=("Arial", 14, "bold"))
        else:
            reminder_label.config(text=f"⚠️ {text} ⚠️", fg=WARNING_COLOR, font=("Arial", 13, "normal"))
        reminder_animation_after_id = root.after(500, lambda: animate_reminder_text(text, blink_count + 1, max_blinks))
    else:
        reminder_label.config(text=f"⚠️ {text} ⚠️", fg=DANGER_COLOR, font=("Arial", 13, "bold"))

def show_operator_reminder_in_header():
    global reminder_label
    reminder_text = "CHECK OPERATOR IN OLD RESERVOIR JUST IN CASE"
    if reminder_label and reminder_label.winfo_exists():
        animate_reminder_text(reminder_text)
        print(f"FLOOD MODE SELECTED - REMINDER: {reminder_text}")

def on_flood_mode_change(event=None):
    selected_flood_mode = flood_mode_var.get()
    if selected_flood_mode == "Flood":
        show_operator_reminder_in_header()
    else:
        clear_reminder()

def clear_reminder():
    global reminder_label, reminder_animation_after_id
    if reminder_animation_after_id:
        root.after_cancel(reminder_animation_after_id)
        reminder_animation_after_id = None
    if reminder_label and reminder_label.winfo_exists():
        reminder_label.config(text="", font=("Arial", 10, "normal"))

# --- Value Generation & Adjustment ---
def regenerate_values(num_rows):
    new_values = [round(random.uniform(1.40, 1.90), 2) for _ in range(num_rows)]
    for i, value_entry in enumerate(values[:num_rows]):
        value_entry.delete(0, tk.END)
        value_entry.insert(0, f"{new_values[i]}")
        update_row_notification(i, "Idle", color=FG_COLOR)

def adjust_values(num_rows):
    try:
        current_value = float(values[0].get())
        for i in range(1, num_rows):
            adjustment = random.uniform(0.05, 0.15)
            new_value = max(1.40, current_value - adjustment)
            values[i].delete(0, tk.END)
            values[i].insert(0, f"{new_value:.2f}")
            current_value = new_value
            update_row_notification(i, "Idle", color=FG_COLOR)
        messagebox.showinfo("Success", "Values adjusted successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# --- Alarm Functions ---
def toggle_alarm():
    global alarm_enabled
    with alarm_lock:
        alarm_enabled = not alarm_enabled
    root.after(0, update_alarm_button_ui)

def update_alarm_button_ui():
    with alarm_lock:
        current_state = alarm_enabled
    if toggle_btn:
        if current_state:
            toggle_btn.config(text="ALARM ON", bg=SUCCESS_COLOR)
        else:
            toggle_btn.config(text="ALARM OFF", bg=DANGER_COLOR)

def play_alarm_sound():
    print("DEBUG: play_alarm_sound() called")
    try:
        if platform.system() == "Windows":
            import winsound
            try:
                for _ in range(3):
                    winsound.Beep(2000, 500)
                    time.sleep(0.2)
            except RuntimeError as e:
                print(f"winsound.Beep failed: {e}")
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
        else:
            print("Alarm would sound here (not Windows).")
    except Exception as e:
        print(f"Alarm sound error: {e}")

def play_alarm_sound_async_safe():
    threading.Thread(target=play_alarm_sound, daemon=True).start()

# --- Playwright Automation Functions ---
async def run_playwright_script(value, row_idx, scheduled_time, headless_mode_enabled, flood_mode, alarm_enabled_state):
    # CRITICAL FIX: Global declaration MUST be at the very top of the function
    global consecutive_failures
    
    if flood_mode == "Select Mode":
        update_row_notification(row_idx, "Cannot start: Select Flood Mode", color=WARNING_COLOR)
        trigger_popup_on_focus(0, 0, scheduled_time, success=False, message="Please select Flood Mode")
        return
    if flood_mode == "Flood" and float(value) < 5.0:
        update_row_notification(row_idx, "Blocked: <5 in Flood", color=WARNING_COLOR)
        trigger_popup_on_focus(0, 0, scheduled_time, success=False, message="Blocked: Value <5 in Flood mode")
        return

    latest = 0.0
    difference = 0.0
    browser = None

    try:
        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(headless=headless_mode_enabled)
            context = await browser.new_context()
            page = await context.new_page()
            
            page.set_default_timeout(60000)
            page.set_default_navigation_timeout(90000)

            # --- Navigate ---
            login_url = 'http://192.168.1.152:8082/production/pages/login.jsp'
            update_row_notification(row_idx, "Navigating to login page...", color=FG_COLOR)
            try:
                await page.goto(login_url, timeout=30000)
                await page.wait_for_load_state('domcontentloaded')
            except Exception:
                update_row_notification(row_idx, "Retrying Navigation...", color=WARNING_COLOR)
                await asyncio.sleep(5)
                await page.goto(login_url, timeout=30000)

            # --- Login ---
            update_row_notification(row_idx, "Logging in...", color=FG_COLOR)
            await page.fill('input[name="username"]', 'laboratory')
            await page.fill('input[name="password"]', 'lab123')
            await page.press('input[name="password"]', 'Enter')
            await page.wait_for_load_state('domcontentloaded', timeout=20000)

            # --- Extract Dam Levels ---
            update_row_notification(row_idx, "Checking Dam Levels...", color=FG_COLOR)
            dam_xpath = '/html/body/table/tbody/tr[2]/th/table/tbody/tr/td[2]/table[3]/tbody/tr[2]'
            try:
                dam_element = await page.wait_for_selector(f'xpath={dam_xpath}', timeout=15000)
                current_text = await dam_element.inner_text()
                
                dam_values = []
                for part in current_text.strip().split():
                    try:
                        dam_values.append(float(part))
                    except ValueError:
                        continue

                if dam_values and len(dam_values) >= 2:
                    latest = dam_values[-1]
                    previous = dam_values[-2]
                    difference = latest - previous

                    if difference >= 0.08 and alarm_enabled_state:
                        update_row_notification(row_idx, "ALARM TRIGGERED!", color=DANGER_COLOR)
                        play_alarm_sound_async_safe()
            except Exception as e:
                print(f"Row {row_idx} DAM LEVEL CHECK ERROR: {e}")

            # --- Refresh Until Value Logic ---
            update_row_notification(row_idx, "Waiting for 0.00...", color=FG_COLOR)
            xpath_target_values = '/html/body/table/tbody/tr[2]/th/table/tbody/tr/td[2]/table[3]/tbody/tr[5]/td[position() >= 4 and position() <= 11]'

            start_time = time.time()
            found_slot = False
            timeout_seconds = 900  # 15 minutes

            while time.time() - start_time < timeout_seconds:
                try:
                    elements = await page.locator(f'xpath={xpath_target_values}').all()
                    found_texts = [await element.inner_text() for element in elements]
                    
                    if any(element_text.strip() == "0.00" for element_text in found_texts):
                        found_slot = True
                        update_row_notification(row_idx, "0.00 found.", color=SUCCESS_COLOR)
                        break

                    update_row_notification(row_idx, "0.00 not found, refreshing...", color=FG_COLOR)
                    await page.reload(wait_until="domcontentloaded", timeout=15000)
                    await asyncio.sleep(15)
                except Exception:
                    try:
                        await page.reload(wait_until="domcontentloaded", timeout=15000)
                    except:
                        pass
                    await asyncio.sleep(15)
            
            if not found_slot:
                raise TimeoutError("Target value '0.00' not found within 15 minutes.")

            # --- Form Submission ---
            success_flag = False
            max_attempts = 3
            
            for attempt in range(max_attempts):
                try:
                    update_row_notification(row_idx, f"Submitting attempt {attempt + 1}...", color=FG_COLOR)
                    
                    turbidity_link = await page.wait_for_selector('text=Turbidity', state='visible', timeout=15000)
                    await turbidity_link.click()
                    await page.wait_for_load_state('domcontentloaded')

                    form_input = await page.wait_for_selector('input[name="tvalue"]', state='visible', timeout=15000)
                    await form_input.click()
                    await form_input.fill(str(value))
                    await form_input.press('Enter')
                    await page.wait_for_load_state('domcontentloaded')

                    checkbox = await page.wait_for_selector('#checkbox', state='visible', timeout=15000)
                    await checkbox.click()

                    submit_button = await page.wait_for_selector('#button', state='visible', timeout=15000)
                    await submit_button.click()
                    await page.wait_for_load_state('domcontentloaded')
                    await asyncio.sleep(2)

                    # Verify submission
                    submit_button_xpath = "/html/body/table/tbody/tr[2]/th/table/tbody/tr/td[1]/table[1]/tbody/tr[2]/td"
                    
                    try:
                        # If button is gone, success
                        await expect(page.locator(f'xpath={submit_button_xpath}')).to_be_hidden(timeout=10000)
                        success_flag = True
                        break
                    except:
                        pass
                        
                    if attempt < max_attempts - 1:
                        await page.reload(wait_until='domcontentloaded')
                        await asyncio.sleep(3)

                except Exception as e:
                    print(f"Submission attempt failed: {e}")
                    if attempt < max_attempts - 1:
                        await asyncio.sleep(3)

            # --- Final Result Handling ---
            if success_flag:
                # SUCCESS - Reset failure counter
                with consecutive_failures_lock:
                    consecutive_failures = 0
                
                trigger_popup_on_focus(latest, difference, scheduled_time, success=True, message=f"Value {float(value):.2f} sent successfully")
                update_row_notification(row_idx, "✓ Success", color=SUCCESS_COLOR)
            else:
                # FAILURE - Logic handled in exception block usually, but if loop finishes without flag:
                raise Exception("Failed to verify submission after 3 attempts.")

    except Exception as e:
        # FAILURE HANDLING
        error_msg = str(e)
        print(f"Row {row_idx} Error: {error_msg}")
        
        with consecutive_failures_lock:
            consecutive_failures += 1
            current_failures = consecutive_failures
        
        update_row_notification(row_idx, "✖ Fail", color=DANGER_COLOR)
        
        # Check consecutive failures
        if current_failures >= 2:
            # CRITICAL WARNING
            play_alarm_sound_async_safe()
            show_error_messagebox("⚠️ CRITICAL: Multiple Failures", 
                                f"Row {row_idx+1} failed!\n"
                                f"⚠️ {current_failures} CONSECUTIVE SUBMISSIONS FAILED!\n"
                                f"Error: {error_msg}")
        else:
            show_error_messagebox("Automation Failed", f"Row {row_idx+1} failed: {error_msg}")
            trigger_popup_on_focus(0, 0, scheduled_time, success=False, message="Submission Failed")

    finally:
        if browser:
            try:
                await browser.close()
            except:
                pass
        
        # Release executing lock for this row
        def release_after_delay():
            time.sleep(5)
            currently_executing.discard(row_idx)
            try:
                execution_lock.release()
            except:
                pass
        threading.Thread(target=release_after_delay, daemon=True).start()

# --- Asyncio Loop Thread Management ---
def start_async_loop_thread():
    global async_loop, async_thread
    async_loop = asyncio.new_event_loop()
    async_thread = threading.Thread(target=run_async_loop, daemon=True)
    async_thread.start()

def run_async_loop():
    asyncio.set_event_loop(async_loop)
    async_loop.run_forever()

def stop_async_loop_thread():
    global async_loop, async_thread
    if async_loop and async_loop.is_running():
        async_loop.call_soon_threadsafe(async_loop.stop)
        async_thread.join(timeout=5)
    async_loop = None
    async_thread = None

# --- Automation Control ---
def set_status(state):
    if status_label:
        if state == "running":
            status_label.config(text="▶ RUNNING", fg=SUCCESS_COLOR)
        elif state == "stopped":
            status_label.config(text="◼ STOPPED", fg=DANGER_COLOR)

def start_automation_thread(num_rows):
    if running_event.is_set():
        return
    if shift_var.get() == "Select Shift":
        messagebox.showwarning("Cannot Start", "Please select a valid Shift before starting automation.")
        return
    if flood_mode_var.get() == "Select Mode":
        messagebox.showwarning("Cannot Start", "Please select a valid Flood Mode before starting automation.")
        return

    running_event.set()
    set_status("running")
    current_shift_key = shift_var.get()
    active_num_rows = shift_definitions[current_shift_key]["rows"]

    if not async_loop or not async_loop.is_running():
        start_async_loop_thread()

    schedule_thread = threading.Thread(target=check_schedule, args=(active_num_rows,), daemon=True)
    schedule_thread.start()

def stop_automation():
    running_event.clear()
    set_status("stopped")

def check_schedule(num_rows):
    global async_loop
    while running_event.is_set():
        now = datetime.datetime.now().strftime("%H:%M:%S")
        for i in range(num_rows):
            if not running_event.is_set():
                return
            
            try:
                if scheduled_times and i < len(scheduled_times) and values and i < len(values):
                    scheduled_time_str = scheduled_times[i].get()
                    current_value_str = values[i].get()

                    if not scheduled_time_str or not current_value_str:
                        continue

                    scheduled_time_str = auto_correct_time(scheduled_time_str)
                    
                    if now == scheduled_time_str:
                        if not validate_time_format(scheduled_time_str):
                            continue
                        
                        if i in currently_executing:
                            continue
                        
                        # Acquire lock non-blocking
                        if not execution_lock.acquire(blocking=False):
                            continue
                        
                        currently_executing.add(i)
                        
                        if async_loop and async_loop.is_running():
                            with alarm_lock:
                                current_alarm_state = alarm_enabled
                            asyncio.run_coroutine_threadsafe(
                                run_playwright_script(
                                    current_value_str, i, scheduled_time_str,
                                    headless_mode.get(), flood_mode_var.get(), current_alarm_state
                                ),
                                async_loop
                            )
                        else:
                            currently_executing.discard(i)
                            execution_lock.release()

            except Exception as e:
                print(f"Schedule Loop Error: {e}")
                
        time.sleep(1)

# --- Tray Icon Functions ---
def create_tray_icon_image():
    img = Image.new('RGB', (64, 64), color=(255, 255, 255))
    d = ImageDraw.Draw(img)
    d.ellipse((16, 16, 48, 48), fill=(56, 182, 255))
    return img

def on_tray_icon_clicked(icon, item):
    root.after(0, show_window)

def show_window():
    root.deiconify()
    global tray_icon
    try:
        if tray_icon:
            tray_icon.stop()
            tray_icon = None
    except Exception:
        pass

def hide_to_tray():
    root.withdraw()
    global tray_icon
    if tray_icon is None:
        try:
            import pystray
        except ImportError:
            messagebox.showerror("Error", "pystray not installed.")
            root.deiconify()
            return

        tray_icon = pystray.Icon("SweetDreams", create_tray_icon_image(), "Sweet Dreams", menu=pystray.Menu(
            pystray.MenuItem("Show", on_tray_icon_clicked),
            pystray.MenuItem("Exit", lambda icon, item: exit_app())
        ))
        threading.Thread(target=tray_icon.run, daemon=True).start()

# --- Application Exit ---
def exit_app():
    global pw_window, tray_icon
    running_event.clear()
    stop_async_loop_thread()
    try:
        if pw_window and pw_window.winfo_exists():
            pw_window.destroy()
    except Exception:
        pass
    try:
        if tray_icon:
            tray_icon.stop()
            tray_icon = None
    except Exception:
        pass
    try:
        root.quit()
        root.destroy()
    except Exception:
        pass
    os._exit(0)

# --- Auto Exit Feature ---
def check_auto_exit_time():
    global auto_exit_time
    if not hasattr(root, '_exit_time_thread_started') or not root._exit_time_thread_started:
        root._exit_time_thread_started = True
    else:
        return
    while True:
        try:
            widget = auto_exit_time
            if widget and widget.winfo_exists():
                val = widget.get().strip()
                now = datetime.datetime.now().strftime("%H:%M:%S")
                if now == val and val:
                    exit_app()
        except Exception:
            pass
        time.sleep(1)

threading.Thread(target=check_auto_exit_time, daemon=True).start()

# --- UI Redraw & Theme Toggle ---
def get_current_ui_state(num_rows_current):
    dam_vals = [entry.get() if entry.winfo_exists() else "" for entry in values[:num_rows_current]]
    sched_times_vals = [entry.get() if entry.winfo_exists() else "" for entry in scheduled_times[:num_rows_current]]
    exit_time_val = auto_exit_time.get() if auto_exit_time and auto_exit_time.winfo_exists() else ""
    notif_texts_vals = [lbl.cget("text") if lbl.winfo_exists() else "" for lbl in notification_labels[:num_rows_current]]
    notif_colors_vals = [lbl.cget("fg") if lbl.winfo_exists() else "" for lbl in notification_labels[:num_rows_current]]
    return dam_vals, sched_times_vals, exit_time_val, notif_texts_vals, notif_colors_vals

def redraw_ui_with_state(dam_values_state, sched_times_state, exit_time_state, notif_texts_state, notif_colors_state, num_rows, auto_exit_value, current_shift_var_value, current_flood_mode_var_value):
    shift_var.set(current_shift_var_value)
    flood_mode_var.set(current_flood_mode_var_value)

    for widget in root.winfo_children():
        widget.destroy()

    main_app(dam_values_load=dam_values_state, sched_times_load=sched_times_state,
             exit_time_val=exit_time_state, notif_texts_load=notif_texts_state,
             notif_colors_load=notif_colors_state, num_rows=num_rows, auto_exit_value=auto_exit_value)

def toggle_theme(num_rows_current, auto_exit_value_current):
    global current_theme
    was_running = running_event.is_set()
    if was_running:
        running_event.clear()
        time.sleep(0.5)
    
    dam_values_state, sched_times_state, exit_time_state, notif_texts_state, notif_colors_state = get_current_ui_state(num_rows_current)
    current_shift_var_value = shift_var.get()
    current_flood_mode_var_value = flood_mode_var.get()

    if current_theme == "light":
        apply_theme(NIGHT_THEME)
        current_theme = "night"
    else:
        apply_theme(LIGHT_THEME)
        current_theme = "light"

    redraw_ui_with_state(dam_values_state, sched_times_state, exit_time_state, notif_texts_state, notif_colors_state,
                         num_rows_current, auto_exit_value_current, current_shift_var_value, current_flood_mode_var_value)
    
    if was_running:
        root.after(100, lambda: running_event.set())
        root.after(200, lambda: set_status("running"))

# --- generate_shift_times ---
def generate_shift_times(shift_key, minute_range=(9, 12), second_range=(1, 51)):
    shift = shift_definitions[shift_key]
    start_hour = shift["start_hour"]
    end_hour = shift["end_hour"]
    hours = []
    for h in range(start_hour, end_hour + 1):
        hours.append(h % 24)
    times = []
    for h in hours:
        minute = random.randint(*minute_range)
        second = random.randint(*second_range)
        times.append(f"{h:02d}:{minute:02d}:{second:02d}")
    return times

default_auto_exit_times = {
    "SECOND SHIFT (8:00 AM - 4:00 PM)": "16:18:00",
    "STRAIGHT SHIFT (5:00 PM - 7:00 AM)": "07:18:00",
    "THIRD SHIFT (4:00 PM - 11:00 PM)": "23:18:00",
    "Select Shift": ""
}

def on_shift_change(event=None):
    selected_shift_label = shift_var.get()
    num_rows_current = len(values) if values else 0
    was_running = running_event.is_set()
    if was_running:
        running_event.clear()
        time.sleep(0.5)

    dam_values_state, _, exit_time_state, notif_texts_state, notif_colors_state = get_current_ui_state(num_rows_current)
    current_flood_mode_var_value = flood_mode_var.get()

    new_sched_times = []
    new_num_rows = 15
    auto_exit_for_shift = default_auto_exit_times.get(selected_shift_label, "")

    if selected_shift_label in shift_definitions:
        new_num_rows = shift_definitions[selected_shift_label]["rows"]
        new_sched_times = generate_shift_times(selected_shift_label)
    else:
        new_num_rows = 0
        new_sched_times = []

    adjusted_dam_values_state = dam_values_state[:new_num_rows] + [""] * max(0, new_num_rows - len(dam_values_state))
    adjusted_notif_texts_state = notif_texts_state[:new_num_rows] + ["Idle"] * max(0, new_num_rows - len(notif_texts_state))
    adjusted_notif_colors_state = notif_colors_state[:new_num_rows] + [FG_COLOR] * max(0, new_num_rows - len(notif_colors_state))

    redraw_ui_with_state(
        adjusted_dam_values_state,
        new_sched_times,
        auto_exit_for_shift,
        adjusted_notif_texts_state,
        adjusted_notif_colors_state,
        num_rows=new_num_rows,
        auto_exit_value=auto_exit_for_shift,
        current_shift_var_value=selected_shift_label,
        current_flood_mode_var_value=current_flood_mode_var_value
    )
    
    if was_running:
        root.after(100, lambda: running_event.set())
        root.after(200, lambda: set_status("running"))

def update_playwright_browsers_thread():
    def _run_update():
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"])
        messagebox.showinfo("Update", "Chromium browser updated.")
    threading.Thread(target=_run_update, daemon=True).start()

def main_app(dam_values_load=None, sched_times_load=None, exit_time_val=None, notif_texts_load=None, notif_colors_load=None, num_rows=15, auto_exit_value="07:18:00"):
    global start_button, stop_button, status_label, auto_exit_time, values, scheduled_times, toggle_btn, notification_labels

    root.title("😴😴Sweet Dreams😴😴")
    root.configure(background=BG_COLOR)

    header_frame = tk.Frame(root, bg=SECONDARY_COLOR)
    header_frame.pack(fill="x", padx=10, pady=10)

    header_label = tk.Label(header_frame, text="😴😴Sweet Dreams😴😴",
                             font=("Arial", 18, "bold"), bg=SECONDARY_COLOR, fg=FG_COLOR)
    header_label.pack(pady=15, side="left")

    create_animated_reminder_label(header_frame)

    shift_dropdown_label = tk.Label(header_frame, text="Shift:", font=("Arial", 11, "bold"), bg=SECONDARY_COLOR, fg=FG_COLOR)
    shift_dropdown_label.pack(side="left", padx=(10, 2))
    shift_dropdown = ttk.Combobox(header_frame, textvariable=shift_var, values=shift_options, state="readonly", width=32)
    shift_dropdown.pack(side="left", padx=(2, 10))
    shift_dropdown.bind("<<ComboboxSelected>>", on_shift_change)

    flood_dropdown_label = tk.Label(header_frame, text="Mode:", font=("Arial", 11, "bold"), bg=SECONDARY_COLOR, fg=FG_COLOR)
    flood_dropdown_label.pack(side="left", padx=(10, 2))
    flood_dropdown = ttk.Combobox(header_frame, textvariable=flood_mode_var, values=flood_options, state="readonly", width=12)
    flood_dropdown.pack(side="left", padx=(2, 10))
    flood_dropdown.bind("<<ComboboxSelected>>", on_flood_mode_change)


    theme_btn = tk.Button(header_frame, text="🌙 Dark Mode" if current_theme=="light" else "☀️ Light Mode", font=("Arial", 10, "bold"), bg=ACCENT_COLOR, fg="white",
        command=lambda: toggle_theme(num_rows, auto_exit_value))
    theme_btn.pack(side="right", padx=10)

    settings_frame = tk.Frame(root, bg=BG_COLOR)
    settings_frame.pack(fill="x", padx=10, pady=5)

    headless_chk = tk.Checkbutton(settings_frame, text="Headless Mode (no browser UI)", variable=headless_mode,
                                  bg=BG_COLOR, fg=FG_COLOR, selectcolor=SECONDARY_COLOR, font=("Arial", 9))
    headless_chk.pack(side="left", padx=5)

    main_frame = tk.Frame(root, bg=BG_COLOR)
    main_frame.pack(fill="both", expand=True, padx=10, pady=5)

    headers = ["Values", "Scheduled Time", "Status", "Controls"]
    HEADER_COLOR = FG_COLOR
    for col, header in enumerate(headers):
        label = tk.Label(main_frame, text=header, font=("Arial", 10, "bold",),
                         bg=BG_COLOR, fg=HEADER_COLOR)
        label.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")

    for i in range(num_rows):
        main_frame.grid_rowconfigure(i+1, weight=1)
    for i in range(4):
        main_frame.grid_columnconfigure(i, weight=1)

    entry_font = ("Arial", 10)
    button_font = ("Arial", 9, "bold")

    scheduled_times_list = []
    if sched_times_load and len(sched_times_load) == num_rows:
        scheduled_times_list = sched_times_load
    elif shift_var.get() in shift_definitions:
        scheduled_times_list = generate_shift_times(shift_var.get())
    else:
        scheduled_times_list = [f"{h:02d}:{random.randint(7,12):02d}:{random.randint(1,51):02d}" for h in range(num_rows)]

    values.clear()
    scheduled_times.clear()
    notification_labels.clear()

    for row in range(num_rows):
        val = dam_values_load[row] if dam_values_load and row < len(dam_values_load) and dam_values_load[row] else f"{random.uniform(1.40, 1.90):.2f}"
        time_val = scheduled_times_list[row] if scheduled_times_list and row < len(scheduled_times_list) else "00:00:00"

        val_entry = ttk.Entry(main_frame, font=entry_font, justify='center', style="TEntry")
        val_entry.grid(row=row+1, column=0, padx=5, pady=3, sticky="nsew")
        val_entry.insert(0, val)
        val_entry.bind("<KeyRelease>", lambda e, ent=val_entry, idx=row: on_value_change(e, ent, idx))
        values.append(val_entry)

        time_entry = ttk.Entry(main_frame, font=entry_font, justify='center', style="TEntry")
        time_entry.grid(row=row+1, column=1, padx=5, pady=3, sticky="nsew")
        time_entry.insert(0, time_val)
        time_entry.bind("<KeyRelease>", lambda e, ent=time_entry, idx=row: on_time_change(e, ent, idx))
        time_entry.bind("<FocusOut>", lambda e, ent=time_entry: on_time_focus_out(e, ent))
        scheduled_times.append(time_entry)

        if row == 0:
            start_button = tk.Button(main_frame, text="▶ START ", font=button_font,
                                     bg=SUCCESS_COLOR, fg="white", bd=0, padx=10, pady=2,
                                     command=lambda: start_automation_thread(num_rows))
            start_button.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            start_button.bind("<Enter>", lambda e, btn=start_button, col=SUCCESS_COLOR: on_enter(e, btn, col))
            start_button.bind("<Leave>", lambda e, btn=start_button, col=SUCCESS_COLOR: on_leave(e, btn, col))

            def update_start_btn_state(*args):
                if shift_var.get() == "Select Shift" or flood_mode_var.get() == "Select Mode":
                    start_button.config(state="disabled")
                else:
                    start_button.config(state="normal")
            shift_var.trace_add("write", update_start_btn_state)
            flood_mode_var.trace_add("write", update_start_btn_state)
            update_start_btn_state()

        elif row == 1:
            stop_button = tk.Button(main_frame, text="⏹ STOP ", font=button_font,
                                     bg=DANGER_COLOR, fg="white", bd=0, padx=10, pady=2,
                                     command=stop_automation)
            stop_button.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            stop_button.bind("<Enter>", lambda e, btn=stop_button, col=DANGER_COLOR: on_enter(e, btn, col))
            stop_button.bind("<Leave>", lambda e, btn=stop_button, col=DANGER_COLOR: on_leave(e, btn, col))

        elif row == 2:
            gen_btn = tk.Button(main_frame, text="🔄 GENERATE VALUES", font=button_font,
                                     bg=ACCENT_COLOR, fg="white", bd=0, padx=10, pady=2,
                                     command=lambda: regenerate_values(num_rows))
            gen_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            gen_btn.bind("<Enter>", lambda e, btn=gen_btn, col=ACCENT_COLOR: on_enter(e, btn, col))
            gen_btn.bind("<Leave>", lambda e, btn=gen_btn, col=ACCENT_COLOR: on_leave(e, btn, col))

        elif row == 3:
            adj_btn = tk.Button(main_frame, text="📉 ADJUST VALUES", font=button_font,
                                     bg=WARNING_COLOR, fg="white", bd=0, padx=10, pady=2,
                                     command=lambda: adjust_values(num_rows))
            adj_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            adj_btn.bind("<Enter>", lambda e, btn=adj_btn, col=WARNING_COLOR: on_enter(e, btn, col))
            adj_btn.bind("<Leave>", lambda e, btn=adj_btn, col=WARNING_COLOR: on_leave(e, btn, col))

        elif row == 4:
            toggle_btn = tk.Button(main_frame, text="🔴 ALARM OFF", font=button_font,
                                     bg=DANGER_COLOR, fg="white", bd=0, padx=10, pady=2,
                                     command=toggle_alarm)
            toggle_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            toggle_btn.bind("<Enter>", lambda e, btn=toggle_btn, col=DANGER_COLOR: on_enter(e, btn, col))
            toggle_btn.bind("<Leave>", lambda e, btn=toggle_btn, col=DANGER_COLOR: on_leave(e, btn, col))
            update_alarm_button_ui()

        elif row == 5:
            alarm_btn = tk.Button(main_frame, text="🔔 TEST ALARM", font=button_font,
                                     bg="#8B0000", fg="white", bd=0, padx=10, pady=2,
                                     command=play_alarm_sound_async_safe)
            alarm_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            alarm_btn.bind("<Enter>", lambda e, btn=alarm_btn, col="#8B0000": on_enter(e, btn, col))
            alarm_btn.bind("<Leave>", lambda e, btn=alarm_btn, col="#8B0000": on_leave(e, btn, col))

        elif row == 6:
            tray_btn = tk.Button(main_frame, text="⬇ HIDE TO TRAY", font=button_font,
                                     bg="purple", fg="white", bd=0, padx=10, pady=2,
                                     command=hide_to_tray)
            tray_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            tray_btn.bind("<Enter>", lambda e, btn=tray_btn, col="purple": on_enter(e, btn, col))
            tray_btn.bind("<Leave>", lambda e, btn=tray_btn, col="purple": on_leave(e, btn, col))

        elif row == 7:
            exit_frame = tk.Frame(main_frame, bg=BG_COLOR)
            exit_frame.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")

            ttk.Label(exit_frame, text="⏰ Auto Exit:", font=("Arial", 9),
                              background=BG_COLOR, foreground=FG_COLOR).pack(side="left", padx=(0, 5))

            auto_exit_time = ttk.Entry(exit_frame, width=11, justify='center',
                                           font=entry_font, style="TEntry")
            auto_exit_time.pack(side="left", fill="y")
            initial_exit_time = exit_time_val if exit_time_val else auto_exit_value
            if len(initial_exit_time.split(':')) == 2:
                initial_exit_time += ":00"
            auto_exit_time.insert(0, initial_exit_time)
            
        elif row == 8:
            def reset_failures():
                global consecutive_failures
                with consecutive_failures_lock:
                    consecutive_failures = 0
                messagebox.showinfo("Reset", "Counter reset.")
            
            reset_btn = tk.Button(main_frame, text="🔄 Reset Failures", font=button_font,
                                     bg="#FF6B6B", fg="white", bd=0, padx=10, pady=2,
                                     command=reset_failures)
            reset_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            reset_btn.bind("<Enter>", lambda e, btn=reset_btn, col="#FF6B6B": on_enter(e, btn, col))
            reset_btn.bind("<Leave>", lambda e, btn=reset_btn, col="#FF6B6B": on_leave(e, btn, col))

        elif row == 9:
            # Placeholder for alignment
            pass

        elif row == 10:
            update_pw_btn = tk.Button(main_frame, text="⚙️ Update ", font=button_font,
                                     bg="#6A5ACD", fg="white", bd=0, padx=10, pady=2,
                                     command=update_playwright_browsers_thread)
            update_pw_btn.grid(row=row+1, column=3, padx=5, pady=3, sticky="nsew")
            update_pw_btn.bind("<Enter>", lambda e, btn=update_pw_btn, col="#6A5ACD": on_enter(e, btn, col))
            update_pw_btn.bind("<Leave>", lambda e, btn=update_pw_btn, col="#6A5ACD": on_leave(e, btn, col))

        notif_text = notif_texts_load[row] if notif_texts_load and row < len(notif_texts_load) and notif_texts_load[row] else "Idle"
        notif_color = notif_colors_load[row] if notif_colors_load and row < len(notif_colors_load) and notif_colors_load[row] else FG_COLOR
        notif_label = tk.Label(main_frame, text=notif_text, font=entry_font, fg=notif_color, bg=BG_COLOR)
        notif_label.grid(row=row+1, column=2, padx=5, pady=3, sticky="nsew")
        notification_labels.append(notif_label)

    status_frame = tk.Frame(root, bg=BG_COLOR)
    status_frame.pack(side="bottom", fill="x", pady=(0, 15))

    status_label = tk.Label(status_frame, text="◼ STOPPED",
                             font=("Arial", 12, "bold"), bg=BG_COLOR, fg=DANGER_COLOR, )
    status_label.pack(pady=5)
    set_status("stopped")
    
    # NEW: Failure indicator
    def update_failures_display():
        with consecutive_failures_lock:
            current_failures = consecutive_failures
        
        if current_failures == 0:
            failures_text = "✓ No failures"
            failures_color = SUCCESS_COLOR
        elif current_failures == 1:
            failures_text = f"⚠ {current_failures} failure"
            failures_color = WARNING_COLOR
        else:
            failures_text = f"🚨 {current_failures} consecutive failures"
            failures_color = DANGER_COLOR
        
        try:
            if hasattr(root, 'failures_label') and root.failures_label.winfo_exists():
                root.failures_label.config(text=failures_text, fg=failures_color)
        except:
            pass
        root.after(2000, update_failures_display)
    
    failures_label = tk.Label(status_frame, text="✓ No failures", 
                             font=("Arial", 9), bg=BG_COLOR, fg=SUCCESS_COLOR)
    failures_label.pack(pady=(0, 5))
    root.failures_label = failures_label
    update_failures_display()

    notification_label = tk.Label(root, text="No new notifications.", fg="black", bg=BG_COLOR, font=("Arial", 10))
    notification_label.pack(side="bottom")
    root.notification_label = notification_label

    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TEntry", fieldbackground=ENTRY_BG, foreground=FG_COLOR,
                    font=("Arial", 10), borderwidth=1, relief="flat")
    style.map("TEntry", fieldbackground=[("active", ENTRY_BG)],
              foreground=[("active", FG_COLOR)])
    style.configure("Warning.TEntry", fieldbackground="#fff3cd", foreground="#856404",
                    font=("Arial", 10), borderwidth=1, relief="flat")

def password_prompt():
    global pw_window

    pw_window = tk.Toplevel(root)
    pw_window.title("🦉🦉🦉")
    pw_window.geometry("350x200")
    pw_window.resizable(False, False)
    pw_window.configure(bg=BG_COLOR)
    pw_window.grab_set()

    content_frame = tk.Frame(pw_window, bg=SECONDARY_COLOR)
    content_frame.place(relx=0.5, rely=0.5, anchor="center", width=300, height=150)

    ttk.Label(content_frame, text="🦉🦉🦉", font=("Arial", 11, "bold"),
              background=SECONDARY_COLOR, foreground=FG_COLOR).pack(pady=(20, 5))

    pw_entry = ttk.Entry(content_frame, show='*', style="TEntry", font=("Arial", 11))
    pw_entry.pack(pady=5, ipady=2)
    pw_entry.focus_set()

    def verify():
        if pw_entry.get() == "32133qwe":
            pw_window.destroy()
            main_app(num_rows=15)
            root.after(100, lambda: shift_var.set("STRAIGHT SHIFT (5:00 PM - 7:00 AM)"))
            root.after(200, on_shift_change)
        else:
            messagebox.showerror("Error", "Incorrect Password!")
            pw_entry.delete(0, tk.END)

    submit_btn = tk.Button(content_frame, text="✅", font=("Arial", 10, "bold"),
                            bg=ACCENT_COLOR, fg="white", bd=0, command=verify)
    submit_btn.pack(pady=(10, 15), ipadx=20, ipady=3)
    submit_btn.bind("<Enter>", lambda e, btn=submit_btn, col=ACCENT_COLOR: on_enter(e, btn, col))
    submit_btn.bind("<Leave>", lambda e, btn=submit_btn, col=ACCENT_COLOR: on_leave(e, btn, col))

    pw_window.update_idletasks()
    width_pw = pw_window.winfo_width()
    height_pw = pw_window.winfo_height()
    x_pw = (pw_window.winfo_screenwidth() // 2) - (width_pw // 2)
    y_pw = (pw_window.winfo_screenheight() // 2) - (height_pw // 2)
    pw_window.geometry(f"{width_pw}x{height_pw}+{x_pw}+{y_pw}")

width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

root.protocol("WM_DELETE_WINDOW", exit_app)

password_prompt()
root.mainloop()