import tkinter as tk
from tkinter import font
from tkinter.font import Font

from data.excel import generate_excel


def download_excel():
    try:
        generate_excel()
        show_status("Excel file saved successfully", "green")
    except Exception as e:
        show_status(f"Error: {str(e)}", "red")


def show_status(text, color):
    statusLabel.config(text=text, fg=color)
    statusLabel.update()
    window.after(2000, hide_status)


def hide_status():
    statusLabel.config(text="")
    statusLabel.update()


window = tk.Tk()
window.title("SnapshotTrello")
window.geometry("500x300")

window.configure(bg="#ffb6c1")
boldFont = Font(weight="bold")

welcomeLabel = tk.Label(window, text="~ Generate the Excel file ~", fg="#b6fff4", bg="#ffb6c1", font=boldFont)
welcomeLabel.pack(pady=10)

buttonFont = font.Font(size=12, weight="bold")
downloadButton = tk.Button(window, text="Download Excel", command=download_excel, bg="#b6fff4", fg="#ffb6c1", height=2,
                           width=15, font=buttonFont)
downloadButton.pack(pady=10)

statusLabel = tk.Label(window, text="", fg="green", bg="#ffb6c1", font=boldFont)
statusLabel.pack(pady=10)

window.mainloop()
