import csv
import pandas as pd
from tkinter import Tk, filedialog, Label, Button
from io import StringIO

# Function to open file dialog and return selected file path
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        label_file.config(text="Selected File: " + file_path)
    return file_path

# Function to convert Excel to VCF
def convert_to_vcf():
    file_path = label_file.cget("text").replace("Selected File: ", "")
    if file_path:
        # Convert Excel to CSV in memory
        read_file = pd.read_excel(file_path)
        csv_data = read_file.to_csv(index=None, header=True,encoding='utf-8-sig')

        # Convert CSV data to VCF
        csv_buffer = StringIO(csv_data)
        reader = csv.reader(csv_buffer)
        
        vcf_path = filedialog.asksaveasfilename(defaultextension=".vcf", filetypes=[("VCF files", "*.vcf")])
        if vcf_path:
            with open(vcf_path, 'w',encoding='utf-8-sig') as vcf:
                i = 0
                for row in reader:
                    vcf.write('BEGIN:VCARD' + "\n")
                    vcf.write('VERSION:3.0' + "\n")
                    vcf.write('NAME:' + row[0] + "\n")
                    vcf.write('FN:' + row[0] + "\n")
                    vcf.write('TEL;CELL:' + row[1] + "\n")
                    vcf.write('END:VCARD' + "\n")
                    vcf.write("\n")
                    i += 1  # Count

            label_status.config(text="VCF file saved at: " + vcf_path)
            

# Create Tkinter window
root = Tk()
root.title("Excel to VCF Converter")

# Set window size
window_width = 400
window_height = 200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width/2) - (window_width/2)
y_coordinate = (screen_height/2) - (window_height/2)
root.geometry("%dx%d+%d+%d" % (window_width, window_height, x_coordinate, y_coordinate))

# Label to display selected file
label_file = Label(root, text="No file selected.")
label_file.pack()

# Button to choose file
btn_choose_file = Button(root, text="Choose File", command=browse_file)
btn_choose_file.pack()

# Button to convert
btn_convert = Button(root, text="Convert", command=convert_to_vcf)
btn_convert.pack()

# Label to display conversion status
label_status = Label(root, text="")
label_status.pack()

# Start the Tkinter event loop
root.mainloop()
