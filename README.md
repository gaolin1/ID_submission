### **Requirements**
- Python 3 (tested on versions 3.9 and 3.10 on mac/windows)
- Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (install command: `pip install pandas`)
2. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (install command: `pip install msoffcrypto-tool`)
3. [streamlit](https://streamlit.io) (install command: `pip install streamlit`)
4. [altair](https://altair-viz.github.io) (install command: `pip install altair`)
5. [tkinter](https://docs.python.org/3/library/tkinter.html) (install command: `pip install tk`)
---

to run the script, either 
1. enter `python3 asp_visual.py` in terminal (mac) or cmd (windows)
2. use right click on the asp_visual.py file and launch with **python launcher** 
---

### This script has two main options
> <img width="654" alt="image" src="https://user-images.githubusercontent.com/28236780/155802819-bcc90e75-fcdf-4cf6-9020-5be44d7b33e1.png">
* Process Epic exported file
  * To process Epic exported report, enter Y.
  * Then select either the exported report are DDD or DOT by entering either 1 or 2.
  * Depending on the exported report, select either to process location, department or both location and department reports (1 or 2 or 3).
  * select the exported excel file from Epic in the dialog window
    * <img width="456" alt="image" src="https://user-images.githubusercontent.com/28236780/155803043-923c6e4c-d892-4431-b089-933dd3b62a93.png">
    * *sample file: HHS_RX_RW_DDDsummary_20220111_1729.xlsx* 
  * enter the password
  * Finally, enter the file name to save under the same folder: 
    * <img width="459" alt="image" src="https://user-images.githubusercontent.com/28236780/155802354-3a5b29c7-b9d6-4ff7-9d8b-08afef6fc25a.png">
    * to save under different folder, select the down arrow to expand the file dialog    
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
  * <img width="869" alt="image" src="https://user-images.githubusercontent.com/28236780/154320518-a58f7290-2ff8-4bc3-a6e1-23151465824e.png">
  * Location or department line graphs will also display.
  * **NOTE** yellow dashed line indicates weighted sum of all selected data.

to stop the program, use "control + c" in terminal to stop the program
