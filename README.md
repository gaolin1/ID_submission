### **Requirements**
- Python 3 (tested on versions 3.9 and 3.10 on mac/windows)
- Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (command: `pip install pandas`)
2. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (command: `pip install msoffcrypto-tool`)
3. [streamlit](https://streamlit.io) (command: `pip install streamlit`)
4. [altair](https://altair-viz.github.io) (command: `pip install altair`)
---

to run the script, either 
1. enter `python3 asp_visual.py` in terminal (mac) or cmd (windows)
2. use right click on the asp_visual.py file and launch with **python launcher** 
---

### This script has two function options
> <img width="493" alt="image" src="https://user-images.githubusercontent.com/28236780/152463994-bacdf740-2d2c-4055-b615-ef46ac540e41.png">
* Process Epic exported file
  * To process Epic exported report, enter Y.
  * Then select either the exported report are DDD or DOT.
    * *sample file ddd_feb_2.xlsx* 
  * Depending on the exported report, select either to process location, department or both location and department reports (1 or 2 or 3).
  * Enter path to exported file, then enter password.
  * Finally, enter file path for the processed file or file name directly to export on the same folder (with extension (xlsx)): 
    * example: 'output file name' *(to export on the same folder of the script)*
    * example: /Users/'User Name'/Downloads/'output file name' *(to export to an exact folder)*
---

* Launch dashboard visualization app
  * To launch visualization page, enter Y.
  * <img width="754" alt="image" src="https://user-images.githubusercontent.com/28236780/152464782-17f28c6e-2a95-4f47-bceb-d97f3dc72532.png">
  * Click **Browse file** to open file explorer and choose the output processed file.
  * Then select either DDD or DOT.
  * Select location, department or both for later summarization table and graphs.
  * <img width="760" alt="image" src="https://user-images.githubusercontent.com/28236780/152464961-1a7e9115-d332-424f-b3e0-611b4d5aa5f6.png">
  * Select grouper of interested.
  * Then select level of data, list of departments will automatically filter based on locations selected.
  * <img width="733" alt="image" src="https://user-images.githubusercontent.com/28236780/152465697-d0c34f2a-29ae-403e-968c-34bccc0e65f8.png">
  * Location or department data summary tables will display.
  * <img width="869" alt="image" src="https://user-images.githubusercontent.com/28236780/153302590-da92432d-0df8-4a26-b060-7e7dff74e0ad.png">
  * Location or department line graphs will also display.
  * **NOTE** yellow dashed line indicates sum of all selected data.

to stop the program, use "control + c" in terminal to stop the program
