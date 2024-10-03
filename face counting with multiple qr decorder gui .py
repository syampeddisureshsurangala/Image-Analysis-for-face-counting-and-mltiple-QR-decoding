import cv2
import pyzbar.pyzbar as pyzbar
import pandas as pd
from tkinter import Tk, filedialog, Text, Entry
from datetime import datetime

# Load the pre-trained Haar cascade for face detection
face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')

# Create a GUI to select the image file
root = Tk()
root.title("QR Code and Face Detection")

# Create a Tkinter Text widget to display the output values
text_widget = Text(root)
text_widget.pack()

# Create a Tkinter Entry widget to input the given values
entry_widget = Entry(root)
entry_widget.pack()

# Ask the user to select an image file
image_path = filedialog.askopenfilename(title="Select an image file", filetypes=[("Image files", "*.png *.jpg")])

if not image_path:
    print("No image selected. Exiting...")
    exit()

# Load the input image
img = cv2.imread(image_path)

# Convert the input image to grayscale
gray_image = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

# Perform face detection
faces = face_cascade.detectMultiScale(gray_image, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))

# Calculate the face count
face_count = len(faces)

# Detect and decode the QR codes in the image
codes = pyzbar.decode(img)

# Create an empty list to store the QR code data
data = []

# Loop through the detected codes
for code in codes:
    # Get the data and the type of the code
    code_data = code.data.decode("utf-8")
    code_type = code.type

    # Append the data to the list
    data.append(code_data)

# Create a pandas DataFrame from the list
df = pd.DataFrame(data, columns=["QR Code Data"])

# Add a new row with the face count
df.loc[len(df)] = [face_count]
df.loc[len(df)-1, "QR Code Data"] = f"Face Count: {face_count}"

# Insert the given values into the Entry widget
entry_widget.insert("end", "Given Values")

# Print the DataFrame to the Text widget
text_widget.insert("end", df.to_string(index=False))

# Get the current time
timestr = datetime.now().strftime("%Y%m%d_%H%M%S")

# Save the DataFrame to an Excel file
df.to_excel("data_"+timestr+".xlsx", index=False)

# Run the Tkinter event loop
root.mainloop()