import tkinter as tk
from PIL import ImageTk, Image, ImageWin
from pdf2image import convert_from_path
import win32printing
import win32ui

# === Print Function ===
def print_preview_image(pil_image):
    printer_name = win32printing.GetDefaultPrinter()
    hprinter = win32ui.CreateDC()
    hprinter.CreatePrinterDC(printer_name)

    # Start the print job
    hprinter.StartDoc("Reçu")
    hprinter.StartPage()

    # Get printable area (HORZRES, VERTRES)
    printable_area = hprinter.GetDeviceCaps(8), hprinter.GetDeviceCaps(10)

    # Resize image to fit page
    img = pil_image.copy()
    img = img.resize(printable_area, Image.LANCZOS)

    # Send image to printer
    dib = ImageWin.Dib(img)
    dib.draw(hprinter.GetHandleOutput(), (0, 0, printable_area[0], printable_area[1]))

    hprinter.EndPage()
    hprinter.EndDoc()
    hprinter.DeleteDC()

# === GUI & Preview ===

# Convert PDF to image
pages = convert_from_path("recu_06_2025.pdf", dpi=150)
image = pages[0]

# Create Tkinter window
root = tk.Tk()
root.title("Receipt Preview")

# Convert to Tkinter-compatible image
tk_image = ImageTk.PhotoImage(image)

# Display the image
label = tk.Label(root, image=tk_image)
label.pack(pady=10)

# Add Print button
print_button = tk.Button(root, text="Imprimer", command=lambda: print_preview_image(image))
print_button.pack(pady=10)

# Run the app
root.mainloop()