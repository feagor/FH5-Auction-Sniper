# Forza-Horizon-5-Auction-Buyout-Sniper

This is the first script using image matching (e.g., OpenCV) to create a much faster and more stable macro for sniping a variety of desired cars in the auction house. Rather than sniping single specific cars, this script aims at fully collection for this game.

Note: This script DOESN'T gaurantee 100% to snipe the auctions. Due to network and other potential issues, you may run it for nothing or get quite a few cars within a long time.

This script is against [Forza Code of Conduct](https://support.forzamotorsport.net/hc/en-us/articles/360035563914-Forza-Code-of-Conduct#:~:text=Forza%20Code%20of%20Conduct%201%20The%20Driver%E2%80%99s%20Code,Suspensions%2FBanning%20...%204%20Appeals%20...%205%20Reporting%20), Use it as YOUR OWN RISK!

## Performance Preview (2MIN Demo)

In this demo, we let the script snipe these four cars `AUDI RS`, `AUDI R1`, `MEGANE R26 R`, `MINI COUNTRYMAN`.

![preview](archive/demo.gif)


## Result Preview
![Result](archive/script_result.PNG)
![ingame result](archive/game_success.png)

## Features

|Name         |Added version           |Breif introduction            |
| ------------- |:-------------:|:-------------:|
| ✅ Fast sniping                             |  v1.0          | Fast speed buyout |
| ✅ Enable single or multi auction snipers   |  v2.0          | Support one or many different car snipers      |
| ✅ Smart auto switch cars                   |  v3.0          | If one auction takes more than 30mins, switch to another car  |
| ✅ Easy set-up                              |  v4.0          | Only needs to set how many cars you want to buy |
| ✅ Memory efficient with 40MB(->80MB)       |  v1.1(->v4.0)  | Less memory costs      |
| ✅ Include all car info                     |  v4.0          | Include short_name, seasons, DLC, Autoshow,etc    |
| ✅ Game pre-check                           |  v4.0          | Game and windows resolution pre-check |
| ✅ Few auction house setting                |  v5.0          | Only needs to set Price|
| ✅ Added multilocales support                |  v6.0         | Added rus locale, support more screen resolution( tested on 2k,4k)

## Limits:
1. [FH5_all_cars_info_v4.xlsx](https://github.com/feagor/FH5-Auction-Sniper/blob/main/FH5_all_cars_info_v4.xlsx) must be up to date. Otherwise, it may buy different cars. (PS: Update at 01.01.2026)

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

## Pre-Requirements
1. System Requirements:

    This script only tests well on windows 10 with 1920*1080 (100% scale).

    ![system requirement](archive/system_setting.png)

2. Game setting: 
    
    I am using [Hyper-V](https://github.com/jamesstringerparsec/Easy-GPU-PV), a GPU Paravirtualization on Windows like virtual box on MacOS. Therefore, the HDR setting shows wired here. But it doesn't matter.

    ![video setting](archive/video_setting.png)

    To save energy and gpu cost, strongly suggest to set "VERY LOW" in grahic setting.

    ![Graphic setting](archive/graphics_setting.png)

3. The shipped locale is `RUS`, so the script expects the screenshots from `images/RUS`. If you prefer another language, change `LOCAL` in `settings.ini`, supply a matching screenshot set (keep the file names identical), and point `LOCAL_MAKE_COL` to the proper workbook columns.

4. Update the workbook defined by `EXCEL_FILENAME` (currently `FH5_all_cars_info_v4.xlsx`).

    For introduction of `CAR MAKE LOCATION` and `CAR MODEL LOCATION`, please see previous tags.
    
    Now, only need to set `BUY NUM` in the file. Super simple and easy!!!
   
## How to run it
1. Run with Python
    
    Python version must below 3.13
```
git clone https://github.com/feagor/FH5-Auction-Sniper.git
cd FH5-Auction-Sniper
pip install -r requirements.txt
python main.py
```

2. Use Compiled Zip 

    Steps: 
    1. Download zip file on [release page](https://github.com/YiwenLu-yiwen/Forza-Horizon-5-Auction-Buyout-Sniper/releases).
    2. Modify the images folder.(No need if you are satisfied pre-requirements)
    3. Modify the `FH5_all_cars_info_v4.csv`.
    4. Run the exe.

## Start and Enjoy
1. Make sure you have checked all above info.

2. Modify the `FH5_all_cars_info_v4.xlsx` file (or whichever workbook you set in `settings.ini`) for your own needs.

3. Optional: Set auction price and other toggles directly in `settings.ini`. 

4. Stay with this screen (Search auctions must be active), then run the script or exe.

![Auction House](archive/auction_house.png)
