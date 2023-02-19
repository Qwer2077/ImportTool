import time

import keyboard
# import time
# import pynput
# import pyautogui
import pydirectinput
import pandas as pd
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import json
import pathlib
import os

root = tk.Tk()
root.title("Import Glass Data From Excel")
root.geometry("500x250+700+350")
file_name = None
label1 = None
label2 = None
is_success = False
df = None

pathlib.Path("./etc").mkdir(parents=True, exist_ok=True)


# Read json file
def startup_check(file_path):
    if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
        # checks if file exists
        # print("File exists and is readable")
        pass
    else:
        # print("Either file is missing or is not readable, creating file...")
        with open(file_path, 'w+') as fp:
            config = {
                "defaultPath": "/"
            }

            json.dump(config, fp, sort_keys=True, indent=4)


def select_file():
    file_types_tuple = (
        ("Spreadsheets", ("*.xls", "*.xlsx", "*.csv")),
        ("all files", "*.*")
    )

    global file_name

    file_name = tk.filedialog.askopenfilename(title="HI",
                                              initialdir=json_file["defaultPath"],
                                              filetypes=file_types_tuple)

    check_file()


def check_file():
    global label1, label2, df, is_success

    if label1 is not None:
        label1.destroy()
    if label2 is not None:
        label2.destroy()

    if file_name[-4:] == ".csv":
        df = pd.read_csv(file_name)
    elif file_name[-4:] == ".xls" or file_name[-5:] == ".xlsx":
        df = pd.read_excel(file_name)
    else:
        is_success = False
        return

    try:
        # start_time = time.time()
        quantity_regex = "(?i)Quantity|QTY|Q"
        height_regex = "(?i)Height|H"
        width_regex = "(?i)Width|Base|W"
        mark_regex = "(?i)Mark|Note1|Unit|M"

        quantity_string = str(df.filter(regex=quantity_regex).columns[0])
        height_string = str(df.filter(regex=height_regex).columns[0])
        width_string = str(df.filter(regex=width_regex).columns[0])
        mark_string = str(df.filter(regex=mark_regex).columns[0])

        df = df[[quantity_string, height_string, width_string, mark_string]]
        # print(time.time() - start_time)
    except IndexError as e:
        tk.messagebox.showerror(f"Error: ", f"Check columns name")
        is_success = False
        return
    except Exception as e:
        tk.messagebox.showerror(f"Error: ", f"{e}")
        is_success = False
        return

    label1 = tk.Label(root, text=file_name)
    label2 = tk.Label(root, text="Read Success")
    label1.pack()
    label2.pack()

    is_success = True


def import_glass():
    pydirectinput.PAUSE = False
    global label1, label2, df, is_success, file_name

    if not is_success:
        return

    is_first = True
    is_success = False
    df = df.fillna('')

    df_size = len(tuple(df.itertuples()))

    for i, row in enumerate(df.itertuples()):
        if keyboard.is_pressed("shift"):
            label2.destroy()
            label2 = tk.Label(root, text="stopped")
            label2.pack()

            return

        height_string = str(row[2])
        width_string = str(row[3])
        quantity_string = str(row[1])
        mark_string = str(row[4])


        if is_first:
            is_first = False
            keyboard.write(height_string)
            keyboard.press("tab")
            keyboard.write(width_string)
            pydirectinput.press("tab")
            pydirectinput.press("tab")
            # keyboard.press("tab")
            # keyboard.press("tab")
            keyboard.write(quantity_string)
            pydirectinput.press("tab")
            keyboard.write(mark_string)
            keyboard.send("ctrl+g")

            # pydirectinput.typewrite(height_string)
            # pydirectinput.press("tab")
            # pydirectinput.typewrite(width_string)
            # pydirectinput.press("tab")
            # pydirectinput.press("tab")
            # pydirectinput.typewrite(quantity_string)
            # pydirectinput.press("tab")
            # # pydirectinput.typewrite(mark_string)
            # keyboard.write(mark_string)

            # pydirectinput.keyDown('ctrl')
            # pydirectinput.press("g")
            # pydirectinput.keyUp("ctrl")
            continue

        keyboard.press("tab")
        keyboard.write(height_string)
        keyboard.press("tab")
        keyboard.write(width_string)
        keyboard.press("tab")
        keyboard.write(quantity_string)
        keyboard.press("tab")
        keyboard.write(mark_string)

        # pydirectinput.press("tab")
        # pydirectinput.typewrite(height_string)
        # pydirectinput.press("tab")
        # pydirectinput.typewrite(width_string)
        # pydirectinput.press("tab")
        # pydirectinput.typewrite(quantity_string)
        # pydirectinput.press("tab")
        # # pydirectinput.typewrite(mark_string)
        # keyboard.write(mark_string)

        # if last row, stop ctrl + g
        if i != df_size - 1:
            keyboard.send("ctrl+g")

            # pydirectinput.keyDown('ctrl')
            # pydirectinput.press("g")
            # pydirectinput.keyUp("ctrl")

    label2.destroy()
    label2 = tk.Label(root, text="Finished Importing")
    label2.pack()
    df = None
    file_name = None


def select_default():
    directory = tk.filedialog.askdirectory()
    json_file["defaultPath"] = directory

    with open("etc/config.json", "w") as f:
        json.dump(json_file, f, indent=4)


startup_check("etc/config.json")
json_file = json.load(open('etc/config.json'))

button2 = tk.Button(root, text="Select Default Directory", command=select_default).pack()
button1 = tk.Button(root, text="Select File", command=select_file).pack()

keyboard.on_release_key("esc", lambda _: import_glass())

root.mainloop()
