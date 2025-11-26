# AutoProcess – Gwyddion Python Plugin

Fast batch toolbox for SPM data processing inside Gwyddion (Python 2.7 / pygwy).

### Screenshot
![AutoProcess GUI](screenshot.png)

### Features
- Fixed or full color range, invert mapping, zero-to-min (reversible)
- Apply any Gwyddion gradient/palette
- Batch rename selected channels
- Crop by coordinates or live rectangle selection (in-place or new channel, keep lateral offsets)
- Load & replay processing macro from log file (exact reproduction of recorded tools)
- Batch save per original file as .gwy (preserves logs & color metadata)
- Save all selected channels merged into one perfect .gwy
- Live file/channel browser with select-all, index-based selection, and file closing

### How to replay a macro 
1. Perform desired processing steps on one SPM file  
2. Export the processing log (`File → Save Log…` → save as plain text)  
3. In AutoProcess → "Data Process" tab → click **Load Log File** and choose the txt file  
4. Check the channels you want to process in the right panel  
5. Click **Replay Selected Channels** → all recorded operations are executed instantly

### Installation (Gwyddion + Python 2.7 + pygwy)

Enabling Python scripting (pygwy) is not trivial due to version dependencies.  
Recommended working setup (still valid in 2025 for maximum compatibility):

1. Download **32-bit Gwyddion**   
   → http://gwyddion.net/download.php

2. Install **Python 2.7.13** (32-bit)  

3. Install the three required packages (gwy, gtk, pygtk) – easiest way: use the pre-built pygwy console bundle from Gwyddion website or follow the official guide:  
   → http://gwyddion.net/documentation/user-guide-en/python-scripting.html

4. Place the plugin file (`autoprocess.py`) into the pygwy folder:  
   - Windows: `C:\Users\<you>\.gwyddion\pygwy\`  
   - Linux/macOS: `~/.gwyddion/pygwy/`

5. Restart Gwyddion → **AutoProcess** appears under the menu **/AutoProcess**

For full Python API reference and examples:  
http://gwyddion.net/documentation/user-guide-en/python-scripting.html

