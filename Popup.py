import tkinter as tk
from tkinter import filedialog
import sys
import os
class Popup:
    def __init__(self):
        self.snapshot_local_axes = 0
        
    def tranformation_popup(self,nozzle_number):
        # GUI window for user input
        self.root = tk.Tk()
        self.root.withdraw()
        input_window=tk.Toplevel(self.root)

        input_window.title(f"Transformation for Nozzle at point: {nozzle_number} ")
        self.local_axes = tk.IntVar(input_window)

        # Add input fields
        self.transform_input = self.create_input_field(input_window, f"Coordinate Transformation for Nozzle at point: {nozzle_number} (XYZ)")
        self.check_box = self.create_check_button(input_window, "Use Local Axes")

        # Add submit button
        submit_button = tk.Button(input_window, text="Submit", command=lambda: self.transform_submit(input_window))
        submit_button.pack(padx=10, pady=10)

        input_window.protocol("WM_DELETE_WINDOW", lambda: self.transform_submit(input_window))

        input_window.mainloop()

   
    def create_input_field(self, window, label_text):
        label = tk.Label(window, text=label_text)
        label.pack(padx=10, pady=5)
        entry = tk.Entry(window)
        entry.pack(padx=10, pady=5)
        return entry
    
    def create_check_button(self,window,label_text):
        axes_checkbox = tk.Checkbutton(window,text = label_text,variable= self.local_axes)
        axes_checkbox.pack(padx=10,pady=5)
        return axes_checkbox    

    def transform_submit(self,window):
        user_transform = self.transform_input.get()
        #print(f"Transform String : {user_transform}")
        self.transform_string = user_transform
        self.snapshot_local_axes=self.local_axes.get()
        window.destroy()
        self.root.quit()
        self.root.destroy()

    def get_transformstring(self):
        return self.transform_string
    def get_local_axes(self):
        return self.snapshot_local_axes

    def destroy(self):
        self.root.quit()
        self.root.destroy()