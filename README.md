# Forza-Horizon-5-Auction-Buyout-Sniper

Forza Horizon 5 auction sniping automation that combines OpenCV template matching, Excel-driven targeting, and PyDirectInput controls to rotate through any number of cars. It automatically focuses the FH5 window, resizes it to the supported capture size, and keeps a Tk overlay on top so you can pause/resume or stop the run without touching the console.

> ⚠️ This automation explicitly violates the [Forza Code of Conduct](https://support.forzamotorsport.net/hc/en-us/articles/360035563914-Forza-Code-of-Conduct#:~:text=Forza%20Code%20of%20Conduct%201%20The%20Driver%E2%80%99s%20Code,Suspensions%2FBanning%20...%204%20Appeals%20...%205%20Reporting%20). Extended use will almost certainly lead to an account ban. You are solely responsible for any consequences—run it entirely at your own risk.

## Highlights

- **Excel-controlled queue** – `FH5_all_cars_info_v4.xlsx` stores make/model coordinates, model picker offsets, and personal buyout quotas. The loop only touches cars with `BUYOUT NUM > 0` and auto-rotates when a timer or quota runs out.
- **Smart search conditioning** – `set_auc_search_cond` resets the auction filters, navigates straight to the stored make/model coordinates, snaps the buyout slider to the closest supported ladder value, and keeps the overlay timer in sync.
- **Tk overlay with pause/resume** – the overlay shows current target, remaining rotation time, remaining buyouts, and purchased count. Pause freezes the countdown, resume re-focuses FH5 automatically so global hotkeys don’t leak into the overlay.
- **Robust window management** – `pre_check` validates monitor DPI via `get_monitors.py`, focuses the FH5 window, resizes it to 1616×939, and optionally takes debug screenshots for every template match.
- **Multi-locale templates** – drop pixel-perfect screenshots under `images/ENG` or `images/RUS`, then point `LOCAL` and `LOCAL_MAKE_COL` at the right workbook columns.

## Performance Preview (2MIN Demo)

In this demo, we let the script snipe these four cars `AUDI RS`, `AUDI R1`, `MEGANE R26 R`, `MINI COUNTRYMAN`.

![preview](archive/demo.gif)


## Result Preview
![Result](archive/script_result.PNG)
![ingame result](archive/game_success.png)

## Capability Overview

- **Multi-car rotation** – timer-based rotation (`SNIPE_MIN_LIMIT`) plus per-car quotas (`BUYOUT NUM`) prevent stalling on a single target.
- **Automated buyout workflow** – template-based navigation handles Search Auctions, Confirm/Finding, Auction actions (`PB.png`, `VS.png`, `AO.png`), and buyout result banners with auto-dismiss + Excel persistence.
- **Pausing without drift** – the overlay freezes the countdown when paused and resumes from the same second, keeping the console loop and UI perfectly aligned.
- **Debug-friendly** – enable `DEBUG_MODE=true` to capture every template region in `debug/screen/` and write a detailed rotating log via `fh5_sniper.log`.
- **Locale flexibility** – each locale has its own `(x,y)` make coordinates and model offsets stored in the workbook, so code changes aren’t required when you re-scan UI grids.

## Data Requirements

- Keep [FH5_all_cars_info_v4.xlsx](https://github.com/feagor/FH5-Auction-Sniper/blob/main/FH5_all_cars_info_v4.xlsx) up to date. Wrong `MODEL LOC` or stale menu coordinates will snipe unintended cars.
- Only rows with `BUYOUT NUM > 0` and `MODEL LOC != -1` are considered. Each car tracks both remaining buyouts and how many you’ve purchased during the current session.
- Locale-specific columns:
    - `LOCAL_MAKE_COL` (e.g., `MAKE LOC (ENG)` or `MAKE LOC (RUS)`) must contain `(x,y)` tuples for the Make grid.
    - `MODEL LOC` is the horizontal offset within the Model row for that make.

## Current settings snapshot

The default `settings.ini` bundled with this repo is tuned to the Russian locale profile:

| Key | Value | Purpose |
| --- | --- | --- |
| `LOCAL` | `RUS` | Selects the localized screenshot set under `images/RUS`. Switch to `ENG` if you want to use the English assets. |
| `LOCAL_MAKE_COL` | `MAKE LOC (RUS)` | Column in the workbook that stores make/model labels for the selected locale. Update when you change `LOCAL`. |
| `EXCEL_FILENAME` | `FH5_all_cars_info_v4.xlsx` | Car catalog that also stores your `BUY NUM` values. |
| `EXCEL_SHEET_NAME` | `all_cars_info` | Worksheet to read from within the Excel file. |
| `DEBUG_MODE` | `false` | Set to `true` if you need verbose logging while troubleshooting runs. |
| `GAME_TITLE` | `Forza Horizon 5` | Window caption used to make sure the game is focused before automation starts. |

Adjust `settings.ini` if your environment differs (e.g., switch locale, rename the workbook, or run the game through a different launcher title).

- ## System Requirements

- Windows 10/11 host with FH5 configured. The automation has been tested on FHD (1920×1080), 2K, and 4K panels thanks to the enforced resize to 1616×939 during `pre_check`.
- Python < 3.13 with the packages in `requirements.txt` **plus** `pywin32` and `wmi` (needed by `get_monitors.py`).
   
## Running the Sniper

### Python workflow

```bash
git clone https://github.com/feagor/FH5-Auction-Sniper.git
cd FH5-Auction-Sniper
pip install -r requirements.txt
python main.py
```

### Packaged release

- Download the latest ZIP from the [releases](https://github.com/feagor/FH5-Auction-Sniper/releases) page or build it locally (next section).
- Unpack next to `settings.ini`, `FH5_all_cars_info_v4.xlsx`, and the `images/` folder. The executable locates everything relative to its own directory (`CURRENT_DIR`).
- Start FH5, navigate to Auction Search, ensure the focus stays on the game window, then launch `FH5Sniper.exe`.

## Configuration Checklist

1. Update `settings.ini`:
   - `LOCAL` + `LOCAL_MAKE_COL` must match the template pack you intend to use.
   - `MAX_BUYOUT_PRICE` snaps to the nearest supported ladder value (see the list in `set_auc_search_cond`).
   - Timing knobs: `WAIT_RESULT_TIME`, `INPUT_DELAY_SCALE`, `SNIPE_MIN_LIMIT` (minutes per rotation) / `SNIPE_SEC_LIMIT` (auto-derived seconds).
   - `DEBUG_MODE=true` to enable region captures and verbose logs.
2. Edit the Excel workbook:
   - Set `BUYOUT NUM` to the number of copies you still need. The script decrements and persists the value using `update_buyout`.
   - Ensure `MODEL LOC` matches the horizontal index inside the grid row.
3. Prepare locale assets:
   - Keep template names identical (`SA.png`, `CF.png`, etc.).
   - Place new assets in `images/<LOCAL>/`.

## Overlay & Controls

- **Pause/Resume** – click the overlay button or use the pause hotkey (if configured); the countdown freezes instantly, and resuming re-focuses FH5 so global ESC presses don’t hit the overlay.
- **Stop** – exits gracefully by setting `STOP_EVENT`, closing the overlay, and stopping the automation loop.
- **Status fields** – car name, time left in the current rotation, remaining buyouts for that car, and purchased count for the whole session update from any thread via `overlay_controller.update_status`.

## Operating Tips

1. Keep FH5 focused on the Search Auction screen. `active_game_window()` will re-focus periodically, but switching away mid-loop may cause template misses.
2. Use the Tk overlay to pause instead of Alt+Tabbing; the timer stays accurate and resume logic will click back into the game automatically.
3. When templates stop matching, enable `DEBUG_MODE` to capture fresh screenshots, then update the Excel coordinates or the image packs.

![Auction House](archive/auction_house.png)
