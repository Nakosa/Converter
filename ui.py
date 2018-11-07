"""GUI for converter script."""
import tkinter as tk
from tkinter import filedialog

from convert import convert_gett, convert_uber

def __main():
    try:
        import os
        os.system('cls')
    except:
        pass
    root = tk.Tk()
    root.withdraw()
    while True:
        print('select a file (xlsx for gett & csv for uber):')
        file_path = filedialog.askopenfilename()
        if file_path.endswith('.xlsx'):
            convert_gett(file_path)
            print(' gett converter complited')
        elif file_path.endswith('.csv'):
            convert_uber(file_path)
            print(' uber converter complited')
        else:
            print(' bad format of file')
        print('press Enter for load new file')
        input()
        print()
    input()

if __name__ == "__main__":
    __main()
