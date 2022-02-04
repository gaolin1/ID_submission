### **Requirements**
- Python 3 (tested on versions 3.9 and 3.10 on mac/windows)
- Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (command: `pip install pandas`)
2. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (command: `pip install msoffcrypto-tool`)
3. [streamlit](https://streamlit.io) (command: `pip install streamlit`)
4. [altair](https://altair-viz.github.io) (command: `pip install altair`)


to run the script, either 
1. enter `python3 asp_visual.py` in terminal (mac) or cmd (windows)
2. use right click on the asp_visual.py file and launch with **python launcher** 


### This script has two function options
* Process Epic exported file
  * <img width="493" alt="image" src="https://user-images.githubusercontent.com/28236780/152463994-bacdf740-2d2c-4055-b615-ef46ac540e41.png">
  * To process Epic exported report, enter Y
  * Then select either the exported report are DDD or DOT
  * Depending on the exported report, select either to process location, department or both location and department reports (1 or 2 or 3)
  * Enter path to exported file, then enter password
  * Finally, enter path to processed output file with extension (xlsx) (e.g. to export within the same directory, simply enter the file name (sample.xlsx))
* Launch dashboard visualization app
  * To launch visualization page, enter Y
  * <img width="754" alt="image" src="https://user-images.githubusercontent.com/28236780/152464782-17f28c6e-2a95-4f47-bceb-d97f3dc72532.png">
  * 
