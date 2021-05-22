import pyautogui
import pydirectinput
import tkinter as tk
import tkinter.messagebox
import time
import pandas as pd
import keyboard

root = tk.Tk()


def write_the_words():
    pydirectinput.PAUSE = False
    isFirst = True
    quantity_regex = "Quantity|QTY"
    height_regex = "Height|HEIGHT"
    width_regex = "Width|Base|WIDTH"
    mark_regex = "Mark|Note1|UNIT"

    quantity_string = str(df.filter(regex=quantity_regex).columns[0])
    height_string = str(df.filter(regex=height_regex).columns[0])
    width_string = str(df.filter(regex=width_regex).columns[0])
    mark_string = str(df.filter(regex=mark_regex).columns[0])

    for i in range(total_rows):

        quantity = str(df[quantity_string].iloc[i])
        height = str(df[height_string].iloc[i])
        width = str(df[width_string].iloc[i])
        mark = str(df[mark_string].iloc[i])

        if isFirst:
            isFirst = False
            pydirectinput.typewrite(height)
            pydirectinput.press("tab")
            pydirectinput.typewrite(width)
            pydirectinput.press("tab")
            pydirectinput.press("tab")
            pydirectinput.typewrite(quantity)
            pydirectinput.press("tab")
            pyautogui.typewrite(mark)

            pydirectinput.keyDown('ctrl')
            pydirectinput.press("g")
            pydirectinput.keyUp("ctrl")

            continue

        pydirectinput.press("tab")
        pydirectinput.typewrite(height)
        pydirectinput.press("tab")
        pydirectinput.typewrite(width)
        pydirectinput.press("tab")
        pydirectinput.typewrite(quantity)
        pydirectinput.press("tab")
        pyautogui.typewrite(mark)

        if i != total_rows - 1:
            pydirectinput.keyDown('ctrl')
            pydirectinput.press("g")
            pydirectinput.keyUp("ctrl")


try:
    print("Start reading")

    df = pd.read_excel("test.xlsx")

    print("Finished reading")

    keyboard.wait("esc")

    total_rows = len(df.index)
    print(df.columns)

    write_the_words()

except Exception as e:
    tk.messagebox.showerror(f"Error: {e} not found", f"{e}")

tk.mainloop()