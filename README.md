### **Setup (ONLY REQUIRED FOR FIRST TIME/SIGNIFICANT UPDATE)**
1. Install python 3. (tested on versions 3.9, 3.10 and 3.11 on mac/windows)
2. To download the latest version of this script, select the release on the right sidebar and download the latest release.
3. Extract the folder to a convenient location.
4. ***Recommended***: to install all required packages, witin the extracted folder, right click on `setup.py` and select "launch with python launcher".
---
- Alternatively, required packages can also be installed inidivudally:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (install command: `pip install pandas`)
2. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (install command: `pip install msoffcrypto-tool`)
3. [streamlit](https://streamlit.io) (install command: `pip install streamlit`)
4. [plotly](https://plotly.com/python/getting-started/) (install command: `pip install plotly`)
5. [tkinter](https://docs.python.org/3/library/tkinter.html) (install command: `pip install tk`)
6. [PySimpleGUI](https://www.pysimplegui.org/en/latest/) (install command: `pip install pysimplegui`)
---

- to run the script:
1. Right click on the `asp_analytics.gui.py` file and launch with **python launcher**  <br> OR enter `python3 asp_analytics_gui.py` in terminal (mac) or cmd (windows)
---

### This script has two main options
> <img width="433" alt="image" src="https://user-images.githubusercontent.com/28236780/232595134-992c707a-de36-4c60-a974-fdbe8a8a453c.png">
* Process Epic exported file
  * First select either the exported report are `DDD` or `DOT`.
  * Then click on "Process Epic Export File"
> <img width="578" alt="image" src="https://user-images.githubusercontent.com/28236780/232595517-9ad416b8-1f2f-426c-bd3e-882506ea21ca.png">
  * To select the Epic output file using `Browse`, enter the file password
    * then click `Import and Process`
    * once finished, click `Save As` to enter the desired output name and then click `Export`

---

* Launch dashboard visualization app
> <img width="433" alt="image" src="https://user-images.githubusercontent.com/28236780/232595134-992c707a-de36-4c60-a974-fdbe8a8a453c.png">
  * To launch visualization page, select `Launch Analytics Dashboard`.
  * <img width="325" alt="image" src="https://user-images.githubusercontent.com/28236780/230796824-67de34cf-e593-4fbc-afa9-2975fc5894c9.png">
  * Click **Browse file** to open file explorer and choose the output processed file.
  * Then select either `DDD` or `DOT`.
  * Select `location`, `department` or `both` for later summarization table and graphs.
  * <img width="1205" alt="image" src="https://user-images.githubusercontent.com/28236780/230796728-eef5d9e8-43b6-4fbf-8807-5381bc25aaa8.png">
  * (Optional: Check "Combine Data For All Imported Hospitals/Departments" For HHS-wide data (medications can be selected at the next step).)
  * <img width="760" alt="image" src="https://user-images.githubusercontent.com/28236780/152464961-1a7e9115-d332-424f-b3e0-611b4d5aa5f6.png">
  * Select `grouper` of interested.
  * Then select level of data, list of departments will automatically filter based on locations selected.
  * <img width="733" alt="image" src="https://user-images.githubusercontent.com/28236780/152465697-d0c34f2a-29ae-403e-968c-34bccc0e65f8.png">
  * Location or department data summary tables will display.
  * <img width="1220" alt="image" src="https://user-images.githubusercontent.com/28236780/230796896-31ae3e09-06ff-4303-b807-7e5bd0ec62c1.png">
  * Location or department line graphs will also display.
  * to stop the web program, use "control + c" in terminal to stop the program
