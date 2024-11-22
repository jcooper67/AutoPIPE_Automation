import tkinter as tk
from tkinter import filedialog
import sys
import os
class Menu:
    def __init__(self,controller):
        self.controller = controller
        self.axes = 0     

    def run(self):
        # GUI window for user input
        self.root = tk.Tk()
        self.root.withdraw()
        try:
            wd = sys._MEIPASS 
        except AttributeError:
            wd = os.getcwd()
        file_path = os.path.join(wd,'files','logo.png')
        icon = tk.PhotoImage(file=file_path)
        self.root.iconphoto(True,icon)

        # Ask for the file path
        file_path = filedialog.askopenfilename(
            title="Select a file", 
            filetypes=[("Excel Files", "*.xlsx"), ("Excel Files", "*.xls"), ("All Files", "*.*")]
        )

        output_path = filedialog.askopenfilename(
            title = "Select The Template Output File",
            filetypes=[("Word Files","*.docx")]
        )


        if file_path:
            input_window = tk.Toplevel(self.root)
            input_window.title("AutoPIPE Automation")

            # Add input fields
            self.add_input_fields(input_window)

            # Add submit button
            submit_button = tk.Button(input_window, text="Submit", command=lambda: self.on_submit(input_window, file_path,output_path))
            submit_button.pack(padx=10, pady=10)

            input_window.mainloop()

    def add_input_fields(self, window):
        self.nozzle_input = self.create_input_field(window, "Nozzle Node Numbers (Comma or Space Separated):")
        self.support_input = self.create_input_field(window, "Support Node Numbers (Comma or Space Separated):")

        
        # self.transform_input = self.create_input_field(window, "Coordinate Transformation (XYZ)")
        # self.axes_input = self.create_check_button(window,"Use Local Axes")

    def create_input_field(self, window, label_text):
        label = tk.Label(window, text=label_text)
        label.pack(padx=10, pady=5)
        entry = tk.Entry(window)
        entry.pack(padx=10, pady=5)
        return entry
    


    def on_submit(self, window, file_path,output_path):
        nozzle_nodes = self.nozzle_input.get().split()
        support_nodes = self.support_input.get().split()

        

        self.controller.handle_user_input(file_path, nozzle_nodes, support_nodes,output_path)
        
        window.destroy()
        self.root.quit()
        self.root.destroy()

    def transform_submit(self,window):
        user_transform = self.transform_input.get()
        self.transform_string = user_transform
        window.destroy()
        self.root.quit()
        self.root.destroy()

    def get_transformstring(self):
        return self.transform_string

    def destroy(self):
        self.root.quit()
        self.root.destroy()