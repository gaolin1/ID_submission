### **Requirements**
- Recommended: to install all required packages, run `pip3 install -r requirements.txt`
- Python 3 (tested on versions 3.9, 3.10 and 3.11 on mac/windows)
- Required packages:
1. [pandas](https://pandas.pydata.org/docs/getting_started/install.html) (install command: `pip install pandas`)
2. [msoffcypto](https://github.com/nolze/msoffcrypto-tool) (install command: `pip install msoffcrypto-tool`)
3. [streamlit](https://streamlit.io) (install command: `pip install streamlit`)
4. [altair](https://altair-viz.github.io) (install command: `pip install altair`)
5. [tkinter](https://docs.python.org/3/library/tkinter.html) (install command: `pip install tk`)
6. [PySimpleGUI](https://www.pysimplegui.org/en/latest/) (install command: `pip install pysimplegui`)
---

to run the script, either 
1. select release on the right sidebar and download the latest release
2. use right click on the asp_visual.py file and launch with **python launcher** OR enter `python3 asp_visual.py` in terminal (mac) or cmd (windows)
---

### This script has two main options
> <img width="547" alt="image" src="https://user-images.githubusercontent.com/28236780/189789926-a728ce53-704a-4bde-8f3d-d4398407fcbc.png">
* Process Epic exported file
  * First select either the exported report are `DDD` or `DOT`.
  * Depending on the exported report, select either to process `location`, `department` or `both` reports.
> <img width="734" alt="image" src="https://user-images.githubusercontent.com/28236780/189790240-5fd82c26-8ccc-422d-875b-43fb6252e3ce.png">
  * for `location` or `department`:
    * select the Epic output file using `Browse`, enter the file password
    * then click `Import and Process`
    * once finished, click `Save As` to enter the desired output name and then click `Export`
> <img width="734" alt="image" src="https://user-images.githubusercontent.com/28236780/189790626-b4aea7ee-2bd7-47dd-aaf6-a403d9e4a1e8.png">
  * for `both`:
    * follow the prompts to select both location and department files and corresponding passwords
    * then click `Import and Process`
    * once finished, click `Save As` to enter the desired output name and then click `Export`

---

* Launch dashboard visualization app
> <img width="547" alt="image" src="https://user-images.githubusercontent.com/28236780/189789926-a728ce53-704a-4bde-8f3d-d4398407fcbc.png">
  * To launch visualization page, select `Launch Analytics Dashboard`.
  * <img width="754" alt="image" src="https://user-images.githubusercontent.com/28236780/152464782-17f28c6e-2a95-4f47-bceb-d97f3dc72532.png">
  * Click **Browse file** to open file explorer and choose the output processed file.
  * Then select either `DDD` or `DOT`.
  * Select `location`, `department` or `both` for later summarization table and graphs.
  * <img width="760" alt="image" src="https://user-images.githubusercontent.com/28236780/152464961-1a7e9115-d332-424f-b3e0-611b4d5aa5f6.png">
  * Select `grouper` of interested.
  * Then select level of data, list of departments will automatically filter based on locations selected.
  * <img width="733" alt="image" src="https://user-images.githubusercontent.com/28236780/152465697-d0c34f2a-29ae-403e-968c-34bccc0e65f8.png">
  * Location or department data summary tables will display.
  * <img width="869" alt="image" src="https://user-images.githubusercontent.com/28236780/154320518-a58f7290-2ff8-4bc3-a6e1-23151465824e.png">
  * Location or department line graphs will also display.
  * **NOTE** yellow dashed line indicates weighted sum of all selected data.
  * to stop the web program, use "control + c" in terminal to stop the program
