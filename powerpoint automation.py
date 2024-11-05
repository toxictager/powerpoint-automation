import tkinter as tk
from tkinter import messagebox

# Create the main window
root = tk.Tk()
root.title("PowerPoint Generator")
root.geometry("400x300")

# Welcome message
welcome_label = tk.Label(root, text="What would you like for your presentation?", font=("Arial", 14))
welcome_label.pack(pady=10)

# Subject input
subject_label = tk.Label(root, text="Subject:")
subject_label.pack()
subject_entry = tk.Entry(root, width=30)
subject_entry.pack(pady=5)

# Number of slides input
slides_label = tk.Label(root, text="Number of slides:")
slides_label.pack()
slides_entry = tk.Entry(root, width=10)
slides_entry.pack(pady=5)

# Number of pictures per slide input
pictures_label = tk.Label(root, text="Pictures per slide:")
pictures_label.pack()
pictures_entry = tk.Entry(root, width=10)
pictures_entry.pack(pady=5)


# Function to capture input and verify entries
def start_creation():
    subject = subject_entry.get()
    slides = slides_entry.get()
    pictures = pictures_entry.get()

    # Basic validation
    if not subject or not slides.isdigit() or not pictures.isdigit():
        messagebox.showerror("Input Error", "Please fill in all fields with valid data.")
        return

    # Convert slides and pictures to integers for further processing
    slides = int(slides)
    pictures = int(pictures)

    messagebox.showinfo("Success",
                        f"Creating a presentation on '{subject}' with {slides} slides and {pictures} pictures per slide.")
    # Here we'll call the API and create the PowerPoint file in the next steps


# Button to start the process
generate_button = tk.Button(root, text="Generate PowerPoint", command=start_creation)
generate_button.pack(pady=20)

root.mainloop()
