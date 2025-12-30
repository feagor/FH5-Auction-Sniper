import os
import sys
import time
import logging
import threading
from logging.handlers import RotatingFileHandler
import configparser

import cv2
import numpy as np
import pandas as pd
from openpyxl import load_workbook

import pygetwindow as gw
import pyautogui as pyau
import pydirectinput as pydi
import get_monitors as gm
import colorama

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

# Read configuration from settings.ini
config = configparser.ConfigParser()
config.read(os.path.join(CURRENT_DIR, 'settings.ini'))

def get_config_value(section, key, default=None, value_type=str):
    """
    Safely read value from config. Returns default on any error.
    value_type can be bool, int, float or callable for conversion.
    """
    try:
        value = config.get(section, key)
        if value_type == bool:
            return value.lower() in ['true', '1', 'yes']
        return value_type(value)
    except Exception:
        return default

# Get config data with defaults
LOCAL = get_config_value('GENERAL', 'LOCAL', 'ENG')
EXCEL_FILENAME = get_config_value('GENERAL', 'EXCEL_FILENAME', 'FH5_all_cars_info_v3.xlsx')
EXCEL_SHEET_NAME = get_config_value('GENERAL', 'EXCEL_SHEET_NAME', 'all_cars_info')
LOCAL_MAKE_COL = get_config_value('GENERAL', 'LOCAL_MAKE_COL', 'MAKE LOC (ENG)')
DEBUG_MODE = get_config_value('GENERAL', 'DEBUG_MODE', False, bool)
GAME_TITLE = get_config_value('GENERAL', 'GAME_TITLE', 'Forza Horizon 5')
# Set internal pause in pyautogui / pydirectinput
pyau.PAUSE = 0
pydi.PAUSE = 0
# Ensure colorama initialized for terminal coloring
colorama.init(wrap=True)
# Constants (colorama codes kept for terminal coloring)
RED_CODE = '\033[1;31;40m'
GREEN_CODE = '\033[1;32;40m'
YELLOW_CODE = '\033[1;33;40m'
BLUE_CODE = '\033[1;34;40m'
CYAN_CODE = '\033[1;36;40m'
COLOR_END_CODE = '\033[0m'

try:
    import tkinter as tk
except Exception:
    tk = None

FIRST_RUN = True
MISSED_MATCH_TIMES = 1
PAUSE_EVENT = threading.Event()
STOP_EVENT = threading.Event()
# Paths
EXCEL_PATH = os.path.join(CURRENT_DIR, EXCEL_FILENAME)
# Templates paths
IMAGE_PATH_SA = os.path.join(CURRENT_DIR, 'images', LOCAL, 'SA.png')
IMAGE_PATH_CF = os.path.join(CURRENT_DIR, 'images', LOCAL, 'CF.png')
IMAGE_PATH_AT = os.path.join(CURRENT_DIR, 'images', LOCAL, 'AT.png')
IMAGE_PATH_BF = os.path.join(CURRENT_DIR, 'images', LOCAL, 'BF.png')
IMAGE_PATH_PB = os.path.join(CURRENT_DIR, 'images', LOCAL, 'PB.png')
IMAGE_PATH_BS = os.path.join(CURRENT_DIR, 'images', LOCAL, 'BS.png')
IMAGE_PATH_NB = os.path.join(CURRENT_DIR, 'images', LOCAL, 'NB.png')
IMAGE_PATH_VS = os.path.join(CURRENT_DIR, 'images', LOCAL, 'VS.png')
IMAGE_PATH_AO = os.path.join(CURRENT_DIR, 'images', LOCAL, 'AO.png')
IMAGE_PATH_HMG = os.path.join(CURRENT_DIR, 'images', LOCAL, 'HMG.png')
IMAGE_PATH_HMBS = os.path.join(CURRENT_DIR, 'images', LOCAL, 'HMBS.png')
IMAGE_PATH_HMMF = os.path.join(CURRENT_DIR, 'images', LOCAL, 'HMMF.png')

# --- logging setup ---------------------------------------------------------
class ColorFormatter(logging.Formatter):
    """Inject ANSI colors for console handler only."""

    def format(self, record):
        message = super().format(record)
        color = getattr(record, 'color', None)
        if color:
            return f"{color}{message}{COLOR_END_CODE}"
        return message


def setup_logging(debug_mode: bool):
    """
    Configure logging: file handler (DEBUG+) and console handler (INFO or DEBUG).
    Returns the configured logger.
    """
    log_path = os.path.join(CURRENT_DIR, 'fh5_sniper.log')
    logger = logging.getLogger('fh5_sniper')
    logger.setLevel(logging.DEBUG)  # capture all levels, handlers decide output

    # Remove existing handlers to avoid duplicate logs on re-import
    if logger.hasHandlers():
        logger.handlers.clear()

    # Rotating file handler: keep DEBUG-level history
    fh = RotatingFileHandler(log_path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh_formatter = logging.Formatter('%(asctime)s %(levelname)s %(name)s: %(message)s')
    fh.setFormatter(fh_formatter)
    logger.addHandler(fh)

    # Console handler: INFO by default, DEBUG if requested
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG if debug_mode else logging.INFO)
    ch_formatter = ColorFormatter('%(asctime)s %(levelname)s: %(message)s', datefmt='%H:%M:%S')
    ch.setFormatter(ch_formatter)
    logger.addHandler(ch)

    return logger

logger = setup_logging(DEBUG_MODE)

def log_and_print(level: str, message: str, color: str = None):
    """
    Helper: log to logger while keeping optional coloring for console output only.
    level: 'debug', 'info', 'warning', 'error'
    color: one of color code constants or None
    """
    log_fn = {
        'debug': logger.debug,
        'info': logger.info,
        'warning': logger.warning,
        'error': logger.error,
    }.get(level, logger.info)

    extra = {'color': color} if color else None
    if extra:
        log_fn(message, extra=extra)
    else:
        log_fn(message)


def wait_if_paused(poll_interval: float = 0.1):
    """Block the automation loop while the pause overlay button is active."""
    while PAUSE_EVENT.is_set() and not STOP_EVENT.is_set():
        time.sleep(poll_interval)


def start_overlay(window_bounds=None):
    """Render a minimal always-on-top overlay with pause/stop controls."""
    if tk is None:
        logger.warning('tkinter is not available. Overlay controls disabled.')
        return

    try:
        root = tk.Tk()
    except Exception:
        logger.exception('Failed to initialize overlay window')
        return

    root.title('FH5 Sniper Control')
    root.configure(bg='#111111')
    root.attributes('-topmost', True)
    root.attributes('-alpha', 0.9)
    root.overrideredirect(True)

    status_var = tk.StringVar(value='Running')

    def place_overlay(bounds):
        if not bounds:
            root.geometry('+80+80')
            return
        left, top, width, height = bounds
        margin = 16
        root.update_idletasks()
        ow = root.winfo_width() or 180
        oh = root.winfo_height() or 120
        x = int(left + width - ow - margin)
        y = int(top + height - oh - margin)
        min_x = int(left + margin)
        min_y = int(top + margin)
        root.geometry(f'+{max(x, min_x)}+{max(y, min_y)}')

    place_overlay(window_bounds)

    def toggle_pause():
        if PAUSE_EVENT.is_set():
            PAUSE_EVENT.clear()
            status_var.set('Running')
            pause_btn.configure(text='Pause')
            log_and_print('info', 'Automation resumed from overlay', GREEN_CODE)
        else:
            PAUSE_EVENT.set()
            status_var.set('Paused')
            pause_btn.configure(text='Resume')
            log_and_print('warning', 'Automation paused from overlay', YELLOW_CODE)

    def request_stop():
        if not STOP_EVENT.is_set():
            STOP_EVENT.set()
            status_var.set('Stopping...')
            log_and_print('warning', 'Stop requested from overlay', RED_CODE)
        try:
            root.destroy()
        except tk.TclError:
            pass

    def start_move(event):
        root._drag_start_x = event.x
        root._drag_start_y = event.y

    def do_move(event):
        x = root.winfo_pointerx() - getattr(root, '_drag_start_x', 0)
        y = root.winfo_pointery() - getattr(root, '_drag_start_y', 0)
        root.geometry(f'+{x}+{y}')

    frame = tk.Frame(root, bg='#111111', padx=12, pady=10)
    frame.pack()

    title_lbl = tk.Label(frame, text='Sniper Overlay', fg='white', bg='#111111', font=('Segoe UI', 10, 'bold'))
    title_lbl.pack(anchor='w')
    status_lbl = tk.Label(frame, textvariable=status_var, fg='#66ff99', bg='#111111', font=('Consolas', 10))
    status_lbl.pack(anchor='w', pady=(2, 8))

    btn_style = {
        'bg': '#1e1e1e',
        'fg': 'white',
        'activebackground': '#4b4b4b',
        'activeforeground': 'white',
        'bd': 0,
        'font': ('Segoe UI', 10, 'bold'),
        'width': 12,
        'pady': 4,
    }

    pause_btn = tk.Button(frame, text='Pause', command=toggle_pause, **btn_style)
    pause_btn.pack(pady=(0, 6))
    stop_btn = tk.Button(frame, text='Stop', command=request_stop, bg='#8b0000', activebackground='#a40000', fg='white', width=12, pady=4, bd=0, font=('Segoe UI', 10, 'bold'))
    stop_btn.pack()

    for widget in (frame, title_lbl, status_lbl):
        widget.bind('<Button-1>', start_move)
        widget.bind('<B1-Motion>', do_move)

    def monitor_stop_flag():
        if STOP_EVENT.is_set():
            try:
                root.destroy()
            except tk.TclError:
                pass
        else:
            root.after(200, monitor_stop_flag)

    root.bind('<Escape>', lambda _event: request_stop())
    root.after(200, monitor_stop_flag)
    root.mainloop()


def debug_screenshot(prefix_name, screenshot_cv):
    if DEBUG_MODE:
        debug_dir = os.path.join(CURRENT_DIR, 'debug', 'screen')
        os.makedirs(debug_dir, exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        ms = int((time.time() % 1) * 1000)
        out_name = f"region_{prefix_name}_{ts}_{ms:03d}.png"
        out_path = os.path.join(debug_dir, out_name)
        # save BGR image
        cv2.imwrite(out_path, screenshot_cv)
    else: pass

    
def get_template_match(image_path, region=None, width_ratio=1, height_ratio=1):
    """
    Take a screenshot of region, read template and run cv2.matchTemplate.
    Returns result matrix.
    """
    screenshot = pyau.screenshot(region=region)
    screenshot_cv = np.array(screenshot)
    screenshot_cv = cv2.cvtColor(screenshot_cv, cv2.COLOR_RGB2BGR)
    template = cv2.imread(image_path, cv2.IMREAD_COLOR)
    screenshot_cv = cv2.resize(screenshot_cv, (int(screenshot_cv.shape[1]/width_ratio), int(screenshot_cv.shape[0]/height_ratio)))
    result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
    base_name = image_path.rsplit('\\', 1)[-1].rsplit('.', 1)[0]
    debug_screenshot(base_name, screenshot_cv)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    log_and_print('debug', f"Max match for {base_name}: {max_val*100:.1f} at {max_loc}")
    return result


def get_best_match_img_array(
    images_path,
    region=None,
    width_ratio=1,
    height_ratio=1,
    threshold=0.8
):
    """
    Find best match among one or multiple template images inside region.
    Returns location (x,y) or (location, index) when multiple images provided.
    """
    images_path = [images_path] if isinstance(images_path, str) else images_path
    return_index = True if len(images_path) > 1 else False
    best_prob,best_index,i = 0,0,0
    best_loc = ()
    for each_image_path in images_path:
        result = get_template_match(each_image_path, region=region, width_ratio=width_ratio, height_ratio=height_ratio)
        loc = np.where(result >= threshold)
        for pt in zip(*loc[::-1]):
            if best_prob < result[pt[1], pt[0]]:
                best_prob = result[pt[1], pt[0]]
                best_loc = (int(pt[0]), int(pt[1]))
                best_index = i
        i += 1
    if best_loc: 
        filename = images_path[0].rsplit('\\', 1)[-1]
        log_and_print('debug', f"Best match for { filename } at location { best_loc } with {best_prob*100:.1f}% probability.")
        if return_index:
            return best_loc, best_index
        return best_loc
    return None


def press_image(image_path, search_region, width_ratio, height_ratio, threshold):
    best_loc = get_best_match_img_array(image_path, search_region, width_ratio, height_ratio, threshold)
    left, top, width, height = search_region
    if best_loc:
        pydi.press('enter')
        return True
    return False


def multi_press(button, times: int, interval: float = 0.1) -> int:
    """Press `button` `times` times with `interval` seconds between presses."""
    if times <= 0:
        return 0
    successful = 0
    for _ in range(int(times)):
        result = pydi.press(button)
        if result:
            successful += 1
        if interval > 0:
            time.sleep(interval)
    return successful


def multi_press_cond(button1, button2, times: int, interval: float = 0.1):
    if times > 0:
        multi_press(button1, times, interval)
    else:
        multi_press(button2, abs(times), interval)


def hold_key(button, secs=5):
    pyau.keyDown(button)
    time.sleep(secs)
    pyau.keyUp(button)


def move_mouse(x, y):
    pyau.moveTo(x, y, duration=0.01)


def click_left():
    pydi.mouseDown()
    time.sleep(0.05)
    pydi.mouseUp()


def multi_click_left(n):
    for _ in range(n):
        click_left()
        time.sleep(0.01)


def reset_car_make():
    active_game_window(GAME_TITLE)
    pydi.press('enter')
    time.sleep(0.5)
    hold_key('w', 4.5)
    hold_key('a', 2)
    time.sleep(0.5)
    pydi.press('enter')
    time.sleep(0.5)
    log_and_print('info', 'Car make reset to ANY', GREEN_CODE)


def set_auc_search_cond(
    Old_Make_Pos,
    Old_Model_Pos,
    New_Make_Pos,
    New_Model_Pos
):
    global FIRST_RUN
    if FIRST_RUN:
        reset_car_make()
        FIRST_RUN = False
    Make_X_Delta, Make_Y_Delta = np.array(Old_Make_Pos) - np.array(New_Make_Pos)

    # set make
    if Make_X_Delta != 0 or Make_Y_Delta != 0:
        pydi.press('enter')
        time.sleep(0.5)    
        #select vertical
        multi_press_cond('w', 's', Make_Y_Delta)
        time.sleep(0.5)
        #select horizontal
        multi_press_cond('a', 'd', Make_X_Delta)        
        time.sleep(1)
        pydi.press('enter')
        time.sleep(0.5)
        
    # GOTO model and set it
    pydi.press('s')    
    time.sleep(1.5)
    if Make_X_Delta == 0 and Make_Y_Delta == 0:  # same make
        model_move_delta = New_Model_Pos - Old_Model_Pos
    else:
        model_move_delta = New_Model_Pos
    multi_press_cond('d', 'a', model_move_delta, 0.3)
    multi_press('s', 5, 0.3)


def active_game_window(title=GAME_TITLE):
    try:
        windows = gw.getWindowsWithTitle(title)
        if not windows:
            return None
        game_window = windows[0]
        try:
            game_window.activate()
        except Exception:
            try:
                game_window.minimize()
                game_window.restore()
            except Exception:
                logger.exception("Failed to activate/restore window")
        return game_window
    except Exception:
        logger.exception("Error getting game window")
        return None


def measure_game_window():
    """Measure the game window size and try to resize it to a fixed resolution for matching."""
    try:
        game_window = active_game_window()
        if game_window:
            try:
                game_window.resizeTo(1616, 939)
            except Exception:
                logger.debug("Could not resize game window, continuing with current size")
            left, top, width, height = game_window.left, game_window.top, game_window.width, game_window.height
            return left, top, width, height
        else:
            log_and_print('error', "Game window not found. Check the title.", RED_CODE)
    except Exception:
        logger.exception("An error occurred while measuring the game window")
    return None, None, None, None


def write_excel(data, output_path, sheet_name):
    data.to_excel(output_path, index=False, sheet_name=sheet_name)
    workbook = load_workbook(output_path)
    sheet = workbook.active
    sheet.auto_filter.ref = sheet.dimensions
    for col in sheet.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        sheet.column_dimensions[col_letter].width = max_length + 2
    workbook.save(output_path)


def exit_script():
    log_and_print('error', 'Script exits in 2 seconds!', RED_CODE)
    time.sleep(2)
    log_and_print('error', 'Script stops!', RED_CODE)
    STOP_EVENT.set()
    sys.exit(0)


def convert_seconds(seconds):
    minutes = int(seconds // 60)
    remaining_seconds = int(seconds % 60)
    return minutes, remaining_seconds


def something_wrong():
    global MISSED_MATCH_TIMES
    log_and_print('warning', f'Fail to match anything. {MISSED_MATCH_TIMES}-th try to press ESC to see whether it works!', RED_CODE)
    active_game_window()
    pydi.press('esc')
    time.sleep(2)
    if MISSED_MATCH_TIMES >= 10:
        log_and_print('error', 'Fail to detect anything, try to restart the script or game!', RED_CODE)
        exit_script()
    MISSED_MATCH_TIMES += 1


def main():
    log_and_print('info', 'Welcome to the Forza 5 CAR BUYOUT Sniper', YELLOW_CODE)
    log_and_print('info', 'Running pre-check: monitor and game resolution', BLUE_CODE)

    monitors = gm.get_monitors()
    logger.debug("Monitors data: %s", monitors)

    forzaWinInfo = gw.getWindowsWithTitle(GAME_TITLE)
    if not forzaWinInfo:
        log_and_print('error', 'Game window not found.', RED_CODE)
        exit_script()
    else:
        w = forzaWinInfo[0]
        left, top, width, height = w.left, w.top, w.width, w.height
        log_and_print('info', f"Game window {GAME_TITLE} found at ({left},{top}) with resolution {width}x{height}", CYAN_CODE)
        cx, cy = left + width / 2, top + height / 2
        mon = gm.find_monitor_for_point(cx, cy, monitors)
        if mon:
            log_and_print('info', f"Center of the game window is at ({cx:.0f}, {cy:.0f}) on monitor: {mon.get('name')} {mon.get('resolution')}", CYAN_CODE)
        else:
            log_and_print('error', f"Center of the game window is at ({cx:.0f}, {cy:.0f}) not found on any monitor", RED_CODE)
            exit_script()

    left, top, width, height = measure_game_window()
    if None in (left, top, width, height):
        exit_script()
    log_and_print('info', f"Resize {GAME_TITLE} resolution to {width}x{height} pixels!", CYAN_CODE)

    game_bounds = (left, top, width, height)
    if tk:
        threading.Thread(target=start_overlay, args=(game_bounds,), name='OverlayUI', daemon=True).start()
    else:
        logger.debug('Overlay disabled: tkinter module not available')

    # screenshot params and regions
    threshold = 0.8
    width_ratio, height_ratio = 1, 1
    REGION_HOME_TABS = (520 + left, 164 + top, 570, 40)
    REGION_AUCTION_MAIN = (230 + left, 590 + top, 910, 310)
    REGION_AUCTION_CAR_DESCR = (790 + left, 190 + top, 810, 90)
    REGION_AUCTION_ACTION_MENU = (525 + left, 330 + top, 530, 190)
    REGION_AUCTION_RESULT = (60 + left, 150 + top, 180, 40)

    log_and_print('info', 'The script will start in 5 seconds', YELLOW_CODE)
    time.sleep(5)
    log_and_print('info', 'Script started', YELLOW_CODE)

    car_needs_swap_fl = True
    failed_snipe = False
    New_Make_Loc, New_Model_Loc = (0, 0), 0
    start_time, all_snipe_index = time.time(), []

    while not STOP_EVENT.is_set():
        wait_if_paused()
        end_time = time.time()
        if end_time - start_time > 1800:
            car_needs_swap_fl = True
            failed_snipe = True

        time.sleep(0.35)
        wait_if_paused()
        if STOP_EVENT.is_set():
            break
        is_search_auc_pressed = press_image(IMAGE_PATH_SA, REGION_AUCTION_MAIN, width_ratio, height_ratio, threshold)
        time.sleep(0.5)
        wait_if_paused()
        if STOP_EVENT.is_set():
            break
        if not is_search_auc_pressed:
            Home_Page_found = get_best_match_img_array([IMAGE_PATH_HMG, IMAGE_PATH_HMBS, IMAGE_PATH_HMMF], REGION_HOME_TABS, width_ratio, height_ratio, threshold)
            if Home_Page_found:
                hold_key('a', 5)
                pydi.press('w')
                pydi.press('enter')
                time.sleep(1)
            else:
                something_wrong()
            continue

        if car_needs_swap_fl:
            log_and_print('info', 'Car need to be changed', GREEN_CODE)
            is_confirm_button_found = get_best_match_img_array(IMAGE_PATH_CF, REGION_AUCTION_MAIN, width_ratio, height_ratio, threshold)
            if is_confirm_button_found:
                if failed_snipe and not FIRST_RUN:
                    end_time = time.time()
                    minutes, remaining_seconds = convert_seconds(end_time - start_time)
                    log_and_print('info', f'[{minutes}:{remaining_seconds}] TIME OUT, Switching to Next Auction Sniper!', YELLOW_CODE)
                failed_snipe = False
                start_time = time.time()
                car_needs_swap_fl = False

                # read file and filter non-zero cars
                df = pd.read_excel(EXCEL_PATH, EXCEL_SHEET_NAME)
                if len(df[df['BUYOUT NUM'] > 0]) == 0:
                    log_and_print('info', 'Finish Sniping!', GREEN_CODE)
                    STOP_EVENT.set()
                    break
                # ignore car model location =-1
                all_snipe_index = df[(df['BUYOUT NUM'] > 0) & (df['MODEL LOC']!=-1)].index.tolist() if all_snipe_index == [] else all_snipe_index
                index = all_snipe_index.pop()
                Old_Make_Loc, Old_Model_Loc = New_Make_Loc, New_Model_Loc

                row = df.iloc[index]
                Make_Name = row.iloc[0]
                Make_Loc = row[LOCAL_MAKE_COL]
                Model_FName = row['CAR MODEL(Full Name)']
                Model_Loc = row['MODEL LOC']
                New_Make_Loc, New_Model_Loc = eval(Make_Loc), Model_Loc
                # reset cursor
                active_game_window()
                move_mouse(left + 10, top + 40)
                multi_click_left(3)
                hold_key('w', 1.5)
                log_and_print('info', f'Setting search to: {Make_Name}, {Model_FName}', GREEN_CODE)
                set_auc_search_cond(Old_Make_Loc, Old_Model_Loc, New_Make_Loc, New_Model_Loc)
                log_and_print('info', f'Start sniping {Model_FName}', GREEN_CODE)
                if STOP_EVENT.is_set():
                    break
            else:
                something_wrong()
                continue

        is_confirm_button_pressed = press_image(IMAGE_PATH_CF, REGION_AUCTION_MAIN, width_ratio, height_ratio, threshold)
        time.sleep(1)
        wait_if_paused()
        if STOP_EVENT.is_set():
            break
        is_auc_res_found = get_best_match_img_array(IMAGE_PATH_NB, REGION_AUCTION_RESULT, width_ratio, height_ratio, threshold)
        if is_auc_res_found:
            logger.debug('Auction results found')
            is_car_found = get_best_match_img_array(IMAGE_PATH_AT, REGION_AUCTION_CAR_DESCR, width_ratio, height_ratio, threshold)
            if is_car_found:
                log_and_print('debug', 'Car found in stock')
                stop = False
                found_PB = found_VS = found_AO = None
                while not stop:
                    if STOP_EVENT.is_set():
                        stop = True
                        break
                    wait_if_paused()
                    time.sleep(0.1)
                    pydi.press('y')
                    found_PB = get_best_match_img_array(IMAGE_PATH_PB, REGION_AUCTION_ACTION_MENU, width_ratio, height_ratio, threshold)
                    found_VS = get_best_match_img_array(IMAGE_PATH_VS, REGION_AUCTION_ACTION_MENU, width_ratio, height_ratio, threshold)
                    found_AO = get_best_match_img_array(IMAGE_PATH_AO, REGION_AUCTION_ACTION_MENU, width_ratio, height_ratio, threshold)
                    if found_PB or found_VS or found_AO:
                        stop = True
                    time.sleep(0.3)

                if found_PB:
                    pydi.press('s')
                    pydi.press('enter')
                    time.sleep(2)
                    pydi.press('enter')
                    time.sleep(5)
                    stop = False

                    while not stop:
                        if STOP_EVENT.is_set():
                            stop = True
                            break
                        wait_if_paused()
                        found_buyoutfail = get_best_match_img_array(IMAGE_PATH_BF, REGION_AUCTION_ACTION_MENU, width_ratio, height_ratio, threshold)
                        found_buyoutsuccess = get_best_match_img_array(IMAGE_PATH_BS, REGION_AUCTION_ACTION_MENU, width_ratio, height_ratio, threshold)
                        if found_buyoutfail:
                            end_time = time.time()
                            minutes, remaining_seconds = convert_seconds(end_time - start_time)
                            log_and_print('info', f'[{minutes}:{remaining_seconds}] BUYOUT Failed!', RED_CODE)
                            pydi.press('enter')
                            pydi.press('esc')
                            stop = True
                        if found_buyoutsuccess:
                            end_time = time.time()
                            minutes, remaining_seconds = convert_seconds(end_time - start_time)
                            log_and_print('info', f'[{minutes}:{remaining_seconds}] BUYOUT Success!', GREEN_CODE)
                            df.loc[index, 'BUYOUT NUM'] = df['BUYOUT NUM'][index] - 1
                            write_excel(df, EXCEL_PATH, EXCEL_SHEET_NAME)
                            if df.loc[index, 'BUYOUT NUM'] == 0:
                                car_needs_swap_fl = True
                                Old_Make_Loc, Old_Model_Loc = New_Make_Loc, New_Model_Loc
                            pydi.press('enter')
                            pydi.press('esc')
                            stop = True
                        time.sleep(3)
                else:
                    end_time = time.time()
                    minutes, remaining_seconds = convert_seconds(end_time - start_time)
                    log_and_print('info', f'[{minutes}:{remaining_seconds}] BUYOUT Missed!', YELLOW_CODE)
                    pydi.press('esc')
                    time.sleep(0.1)
                    if STOP_EVENT.is_set():
                        break
            elif is_car_found is None and is_auc_res_found and is_confirm_button_pressed:
                log_and_print('debug', 'Car not found in stock')
                MISSED_MATCH_TIMES = 1
                pydi.press('esc')
                time.sleep(0.5)
                continue
        else:
            log_and_print('debug', 'Auction results not found :(')
            Home_Page_found = get_best_match_img_array([IMAGE_PATH_HMG, IMAGE_PATH_HMBS, IMAGE_PATH_HMMF], REGION_HOME_TABS, width_ratio, height_ratio, threshold)
            something_wrong()
            continue

    STOP_EVENT.set()
    log_and_print('info', 'Automation stopped.', YELLOW_CODE)


if __name__ == "__main__":
    main()
