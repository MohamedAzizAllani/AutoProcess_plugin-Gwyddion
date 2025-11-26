# AutoProcess – Gwyddion Python Plugin

Batch processing toolbox for SPM/AFM data.

 ### Screenshot
<img src="https://github.com/user-attachments/assets/baf0cec1-ff62-4096-82e6-f244d9f6dad2" width="400"/>

### Features
- Fixed or full color range, invert mapping, zero-to-min 
- Apply any Gwyddion gradient/palette
- Batch rename selected channels
- Crop by coordinates or live rectangle selection 
- Load & replay processing macro from log file
- Batch save per original file as .gwy 
- Save all selected channels merged into one  .gwy
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
   → [https://gwyddion.net/documentation/user-guide-en/pygwy.html](https://gwyddion.net/documentation/user-guide-en/pygwy.html)

4. Place the plugin file (`autoprocess.py`) into the pygwy folder:  
   - Windows: `C:\Users\<you>\.gwyddion\pygwy\`  
   - Linux/macOS: `~/.gwyddion/pygwy/`

5. Restart Gwyddion → **AutoProcess** appears under the menu **/AutoProcess**

For full Python API reference and examples:  
http://gwyddion.net/documentation/user-guide-en/python-scripting.html

