import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import zipfile
import os

class SpectraApp:
    def __init__(self, root):
        self.root = root
        self.root.title("dx extractor vs 1.8")  # Set the title of the window
        self.root.geometry("400x200")  # Set the size of the window
        self.center_window()  # Center the window on the screen
        self.create_widgets()

    def center_window(self):
        window_width = self.root.winfo_reqwidth()
        window_height = self.root.winfo_reqheight()
        position_right = int(self.root.winfo_screenwidth() / 2 - window_width / 2)
        position_down = int(self.root.winfo_screenheight() / 2 - window_height / 2)
        self.root.geometry("+{}+{}".format(position_right, position_down))
    def create_widgets(self):
        title_label = tk.Label(self.root, text="Extract Spectra", font=("Arial", 20))
        title_label.pack()

        input_label = tk.Label(self.root, text="File for processing path")
        input_label.pack()
        self.input_button = tk.Button(self.root, text="Select File(s)", command=self.select_input_files)
        self.input_button.pack()

        output_label = tk.Label(self.root, text="Export Excel file path")
        output_label.pack()
        self.output_button = tk.Button(self.root, text="Select Directory", command=self.select_output_directory)
        self.output_button.pack()

        self.combine_var = tk.IntVar()
        combine_check = tk.Checkbutton(self.root, text="To a single file", variable=self.combine_var)
        combine_check.pack()

        process_button = tk.Button(self.root, text="Process", command=self.process_files)
        process_button.pack()

    def select_input_files(self):
        self.input_file_paths = filedialog.askopenfilenames(filetypes=[('Zip and DX files', '*.zip *.dx')])

    def select_output_directory(self):
        self.output_directory = filedialog.askdirectory()

    def process_files(self):
        if not hasattr(self, 'input_file_paths') or not hasattr(self, 'output_directory'):
            print("Please select an input file and an output directory.")
            return

        # Initialize an empty DataFrame
        df = pd.DataFrame()

        for input_file_path in self.input_file_paths:
            file_name, file_extension = os.path.splitext(input_file_path)
            if file_extension == '.zip':
                # Extract the .dx file from the zipped input file
                with zipfile.ZipFile(input_file_path, 'r') as zip_ref:
                    for file_name in zip_ref.namelist():
                        if file_name.endswith('.dx'):
                            dx_file_name = file_name
                            zip_ref.extract(dx_file_name)
                            break  # Stop searching after finding the first .dx file
            elif file_extension == '.dx':
                dx_file_name = input_file_path

            # Read the data from the .dx file
            with open(dx_file_name, 'r') as dx_file:
                dx_content = dx_file.read()

            # Split the content into blocks based on '##TITLE'
            blocks = dx_content.split('##TITLE')

            # Process each block
            for block in blocks[1:]:  # Skip the first block as it does not contain spectra data
                lines = block.split('\n')
                # Extract sample name from the TITLE line
                sample_name = lines[0].split('=')[1].strip()

                # Find the start of the spectra data
                spectra_start_index = next((i for i, line in enumerate(lines) if line.strip().startswith('##XYDATA')), None)
                if spectra_start_index is None:
                    print(f"Skipping block due to missing spectra data: {lines[0]}")
                    continue

                # Extract spectra data from subsequent lines
                spectra_data = {}
                for line in lines[spectra_start_index+1:]:
                    line = line.strip()  # Remove leading and trailing whitespaces
                    values = line.split(' ')
                    try:
                        wavelength = float(values[0])
                        for val in values[1:]:
                            try:
                                spectra_data[f'{wavelength:.1f}'] = float(val)
                                wavelength += 5
                            except ValueError:
                                continue
                    except ValueError:
                        print(f"Skipping line due to ValueError: {line}")
                        continue

                # Create a dictionary with sample name and spectra data
                data_dict = {'sample_name': sample_name, **spectra_data, 'file_name': os.path.basename(dx_file_name)}
                data_df = pd.DataFrame([data_dict])            
                # Concatenate the DataFrame to the existing DataFrame
                df = pd.concat([df, data_df], ignore_index=True)

            if self.combine_var.get() == 0:
                output_file_path = os.path.join(self.output_directory, f'{os.path.basename(dx_file_name)}_output.csv')
                df = df.loc[:, (df != 0).any(axis=0)]  # Remove the columns that only contain zeros
                df.to_csv(output_file_path, index=False)
                df = pd.DataFrame()  # Reset the DataFrame for the next file

        if self.combine_var.get() == 1:
            output_file_path = os.path.join(self.output_directory, 'combined_output.csv')
            df = df.loc[:, (df != 0).any(axis=0)]  # Remove the columns that only contain zeros
            df.to_csv(output_file_path, index=False)

        print(f"Processing completed. Output saved to:\n{output_file_path}")

root = tk.Tk()
app = SpectraApp(root)
root.mainloop()
