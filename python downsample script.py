# -*- coding: utf-8 -*-
"""
Created on Fri May  3 12:47:01 2024

@author: 100575352
"""
import pandas as pd
import numpy as np
from tkinter import filedialog, Tk, simpledialog
import matplotlib.pyplot as plt

def select_excel_file():
    """ Open a dialog to select an Excel file """
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return file_path

def get_user_input(prompt, type_=None, min_=None, max_=None, range_=None):
    """ Get user input with type checking and range validation """
    root = Tk()
    root.withdraw()
    input_str = simpledialog.askstring("Input", prompt, parent=root)
    try:
        val = type_(input_str)
        if ((min_ is not None and val < min_) or
            (max_ is not None and val > max_) or
            (range_ is not None and val not in range_)):
            raise ValueError("Input out of valid range.")
    except TypeError:
        print("Input type must be", type_.__name__)
    except ValueError as e:
        print(e)
    else:
        return val

def downsample_data(df, original_freq, target_freq):
    """ Downsample the data by averaging over chunks """
    factor = int(np.ceil(original_freq / target_freq))  # Adjust factor for frequencies less than 1 Hz
    return df.groupby(np.arange(len(df)) // factor).mean()

def save_dataframe(df):
    """ Save DataFrame to an Excel file chosen by the user """
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    df.to_excel(file_path, index=False)
    root.destroy()

def plot_data(original_df, downsampled_df):
    """ Plot original and downsampled data for each column """
    for column in original_df.columns:
        plt.figure()
        plt.plot(original_df.index, original_df[column], label='Original Data')
        plt.plot(downsampled_df.index * (len(original_df) / len(downsampled_df)), downsampled_df[column], label='Downsampled Data', linestyle='--')
        plt.title(column)
        plt.xlabel('Sample Index')
        plt.ylabel('Value')
        plt.legend()
        plt.show()

def main():
    # File selection and user input
    file_path = select_excel_file()
    starting_row = get_user_input("Enter the starting row (1-indexed):", int, 1)
    original_freq = get_user_input("Enter the original sampling frequency (Hz):", float, 0.001)
    target_freq = get_user_input("Enter the desired output frequency (Hz):", float, 0.001)

    # Load the data
    df = pd.read_excel(file_path, skiprows=starting_row - 1)
    print("Data loaded successfully. Column names are:", df.columns.tolist())

    # Downsample the data
    downsampled_df = downsample_data(df, original_freq, target_freq)
    print("Data downsampled successfully.")

    # Plot the data
    plot_data(df, downsampled_df)

    # Save the downsampled data
    save_dataframe(downsampled_df)

if __name__ == "__main__":
    main()
