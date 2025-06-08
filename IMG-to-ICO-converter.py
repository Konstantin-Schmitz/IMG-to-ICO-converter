import sys
import os
import ctypes
import tkinter as tk
import pyperclip
import subprocess
from math import sqrt, ceil
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from colorsys import hls_to_rgb

# Make the application DPI aware
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Sharpness, because of DPI
except Exception:
    pass


selected_images_paths = []


#       |=====================|
#       |      FUNCTIONS      |
#       |=====================|


# Function to open the default file explorer and highlight specific files (Windows only)
def open_file_locations(successful_conversion_images_paths):
    unique_folders = set()
    for file_path in successful_conversion_images_paths:
        unique_folders.add(os.path.abspath(os.path.dirname(file_path)))

    for folder_path in unique_folders:
        subprocess.run(["explorer", folder_path])


# Function to show success message with an Open Folder option
def show_conversion_completed_message(successful_conversion_images_paths):
    global selected_images_paths
    files_count = len(selected_images_paths)
    success_count = len(successful_conversion_images_paths)

    # Build a list of filenames with success/failure status
    checkmark = "✔"
    cross = "✖"
    file_status_list = [
        f"{checkmark if img in successful_conversion_images_paths else cross} {os.path.splitext(os.path.basename(img))[0]}.ico"
        for img in selected_images_paths
    ]

    # Show only the first 10 files, then add "..."
    file_status_text = "\n".join(file_status_list[:10])
    if len(file_status_list) > 10:
        file_status_text += "\n..."
    response = messagebox.askyesno(
        "Completed",
        f"Conversion completed\n{success_count}/{files_count} images converted successfully.\n\n{file_status_text}\n\n"
        "Do you want to open the file locations?",
    )

    # Open file locations if the user confirms
    if response:
        open_file_locations(successful_conversion_images_paths)


# Function to convert the selected images to icon files
def convert_to_icon():
    print(f"\033[38;5;81mFUNCTION -> convert_to_icon()\033[0m")

    global selected_images_paths
    if not selected_images_paths:
        messagebox.showerror("Error", "Please select one or more files.")
        return

    successful_conversion_images_paths = []
    icon_sizes = [
        (256, 256),
        (128, 128),
        (96, 96),
        (64, 64),
        (48, 48),
        (32, 32),
        (24, 24),
        (16, 16),
    ]

    display_status(
        "info",
        f"Converting Images... (0/{len(selected_images_paths)})",
        3000,
    )

    for index, image_path in enumerate(selected_images_paths):
        folder = os.path.dirname(image_path)
        base_name = os.path.splitext(os.path.basename(image_path))[0]
        icon_file_name = f"{base_name}.ico"
        output_icon_path = os.path.join(folder, icon_file_name).replace(
            "\\", "/"
        )  # Compatibility fix

        try:
            with Image.open(image_path) as img:
                if img.mode not in ["RGBA", "RGB"]:
                    print(f"image converted tot RGBA")
                    img = img.convert("RGBA")

                width, height = img.size

                if width != height:
                    max_size = max(width, height)
                    square_img = Image.new("RGBA", (max_size, max_size), (0, 0, 0, 0))
                    square_img.paste(
                        img, ((max_size - width) // 2, (max_size - height) // 2)
                    )
                else:
                    square_img = img

                square_img.save(output_icon_path, format="ICO", sizes=icon_sizes)

            successful_conversion_images_paths.append(image_path)

        except Exception as e:
            print(f"Error saving ICO ({image_path}): {e}")
            display_status("error", f"error:\n{str(e)}", 3000)

            error_message = f"An error occurred during conversion:\n{str(e)}\n\n"
            error_question = "Would you like to copy the error message?"
            response = messagebox.askyesno(
                "Error", error_message + error_question, icon=messagebox.ERROR
            )

            if response:
                pyperclip.copy(error_message)
                display_status("info", "Error message copied to clipboard.", 3000)

            continue

        display_status(
            "info",
            f"Converting Images... ({index + 1}/{len(selected_images_paths)})",
            3000,
        )

    show_conversion_completed_message(successful_conversion_images_paths)


# Store the reference to the scheduled "after" callback
status_after_id = None


# Function to display a status
def display_status(type, status_message, duration=None):
    global status_after_id

    formatted_status_message = status_message.replace("\n", " ")

    if status_after_id:  # Cancel the previous countdown if there was one
        root.after_cancel(status_after_id)

    if type == "error":
        label_status_message.config(text=formatted_status_message, fg="red")
    else:
        label_status_message.config(text=formatted_status_message, fg="black")

    if duration:
        status_after_id = root.after(
            duration, lambda: label_status_message.config(text="")
        )
    label_status_message.update_idletasks()


# Function to update the filepath display text
def update_filepath_display():
    global selected_images_paths

    text_filepath_display.configure(state="normal")

    # Clear the display if no image paths are provided
    if not selected_images_paths:
        text_filepath_display.delete(1.0, "end")
        text_filepath_display.insert("end", "No Images Selected...")
    else:
        # If there are image paths, process them
        short_image_paths = []
        for image_path in selected_images_paths:
            normalized_path = os.path.normpath(image_path)
            path_parts = normalized_path.split(os.sep)

            # Only shorten the path if there are more than 4 parts
            if len(path_parts) > 4:
                short_path = os.sep.join(
                    [
                        path_parts[0],
                        path_parts[1],
                        "...",
                        path_parts[-2],
                        path_parts[-1],
                    ]
                )
                short_image_paths.append(short_path)
            else:
                short_image_paths.append(image_path)

        # Update the display with all the processed paths
        text = "\n".join(short_image_paths)
        text_filepath_display.delete(1.0, "end")
        text_filepath_display.insert("end", text)

    text_filepath_display.configure(state="disabled")

    # Adjust the height of the filepath display
    text_lines = min(len(selected_images_paths), text_filepath_display_max_lines)
    text_filepath_display.config(height=text_lines)


def update_preview_image():
    global selected_images_paths

    # Create empty canvas for grid
    grid_width, grid_height = 1980, 1080

    if not selected_images_paths:
        preview_image = default_image
    else:
        preview_image = Image.new(
            "RGBA", (grid_width, grid_height), (255, 255, 255, 255)
        )
        # rastering
        n = len(selected_images_paths)
        columns = ceil(sqrt(n))
        rows = ceil(n / columns)
        cell_width = grid_width // columns
        cell_height = grid_height // rows
        padding = 5  # px

        # Resize and paste each image into the grid
        for i, image_path in enumerate(selected_images_paths):
            img = Image.open(image_path).convert("RGBA")

            # Resize image while maintaining aspect ratio
            img_width = (grid_width // columns) - (2 * padding)
            img_height = (grid_height // rows) - (2 * padding)
            img.thumbnail((img_width, img_height))  # Resize to cell
            # img.thumbnail((img_width, img_height), Image.NEAREST)  # Resize to cell

            # Calculate the padding needed to center the image
            img_width, img_height = img.size
            padding_x = (cell_width - img_width) // 2
            padding_y = (cell_height - img_height) // 2

            # Create cell image and paste it on the grid image
            x = (i % columns) * cell_width + padding_x
            y = (i // columns) * cell_height + padding_y
            preview_image.paste(img, (x, y))

    # Load the preview image for preview
    resize_image(img=preview_image)

    # Bind resizing event for the preview label to trigger dynamic resizing (only once the image is selected)
    label_preview.bind("<Configure>", lambda event: resize_image(event, preview_image))


def resize_image(event=None, img=None):
    if img is None:
        return

    label_width = label_preview.winfo_width()
    label_height = label_preview.winfo_height()
    img_resized = img.copy()
    img_resized.thumbnail((label_width, label_height))
    img_preview = ImageTk.PhotoImage(img_resized)
    label_preview.config(image=img_preview)
    label_preview.image = img_preview  # Keep a reference to avoid garbage collection


def add_images():
    global selected_images_paths

    # Open the file dialog to select one or more images
    add_image_paths = filedialog.askopenfilenames(
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")]
    )

    # Check if the user canceled the file dialog
    if not add_image_paths:
        return

    selected_images_paths.extend(add_image_paths)
    update_filepath_display()
    update_preview_image()
    display_status(
        "info",
        f"{len(add_image_paths)} image{'s' if len(add_image_paths) != 1 else ''} added",
        3000,
    )
    label_selected_images.config(text=f"{len(selected_images_paths)} Images Selected")


def remove_all_images():
    global selected_images_paths

    selected_images_paths = []
    update_filepath_display()
    update_preview_image()
    display_status("info", "All Images Removed", 3000)
    label_selected_images.config(text=f"{len(selected_images_paths)} Images Selected")


def quit_application():
    global selected_images_paths

    # If there are any files selected
    if selected_images_paths:
        response = messagebox.askyesno(
            "Quit?",
            f"Are you sure you want to quit? You currently have {len(selected_images_paths)} images selected",
        )

        # quit the application if the user confirms
        if response:
            root.quit()
    else:
        root.quit()


#       |================|
#       |      MAIN      |
#       |================|


# Determine the correct path to the image, accounting for PyInstaller
if hasattr(sys, "_MEIPASS"):
    image_path = os.path.join(
        sys._MEIPASS, "Source Files", "default-placeholder - 1920x1080.png"
    )
else:
    image_path = os.path.join("Source Files", "default-placeholder - 1920x1080.png")


# Set up the Tkinter window
root = tk.Tk()
root.title("IMG to ICO Converter")

# Set up the layout
window_width = 600
window_height = 600
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x}+{y}")
root.minsize(window_width, window_height)

# ROOT: Frame: Main Frame
frame_main = tk.Frame(root, bg="gray")
frame_main.grid(row=0, column=0, sticky="nsew")

# MAIN FRAME: Frame: Content Frame
frame_content = tk.Frame(frame_main, bg="gray")
frame_content.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# MAIN FRAME: Frame: Status Bar Frame
frame_status_bar = tk.Frame(frame_main, bg="lightgray", height=30)
frame_status_bar.grid(row=1, column=0, sticky="ew")


# CONTENT FRAME: Label: Title
label_title = tk.Label(
    frame_content,
    text="Select one or more image files to convert into .ico format",
    font=("TkDefaultFont", 12),
)
label_title.grid(row=0, column=0, pady=10, sticky="n")

# CONTENT FRAME: Frame: Selection Buttons
frame_selection_buttons = tk.Frame(frame_content, bg="gray")
frame_selection_buttons.grid(row=1, column=0, pady=5, sticky="n")

# CONTENT FRAME: SELECTION BUTTONS FRAME: Button: Add Images
button_add = tk.Button(frame_selection_buttons, text="Add Images", command=add_images)
button_add.grid(row=0, column=0, padx=5, sticky="n")

# CONTENT FRAME: SELECTION BUTTONS FRAME: Button: Remove All Images
button_remove = tk.Button(
    frame_selection_buttons, text="Remove All Images", command=remove_all_images
)
button_remove.grid(row=0, column=1, padx=5, sticky="n")


# CONTENT FRAME: Text: Display of Filepaths of Selected Files
text_filepath_display = tk.Text(frame_content, height=1, wrap="none")
text_filepath_display.insert("1.0", "No Images Selected...")
text_filepath_display.config(state="disabled")
text_filepath_display.grid(row=2, column=0, pady=5, sticky="nsew")
text_filepath_display_max_lines = 5

# CONTENT FRAME: Label: Image Preview
label_preview = tk.Label(frame_content)
label_preview.grid(row=3, column=0, pady=5, sticky="nsew")
label_preview_height = 200
label_preview.grid_configure(ipady=label_preview_height)

# Load the default image
default_image = Image.open(image_path)
label_preview.update_idletasks()
default_image.thumbnail((label_preview.winfo_width(), label_preview_height))
default_image_tk = ImageTk.PhotoImage(default_image)
label_preview.config(image=default_image_tk)
label_preview.image = default_image_tk

# CONTENT FRAME: Frame: Convert / Quit Buttons
frame_convert_quit_buttons = tk.Frame(frame_content, bg="gray")
frame_convert_quit_buttons.grid(row=4, column=0, pady=5, sticky="n")

# CONTENT FRAME: CONVERT QUIT BUTTONS FRAME: Button: Convert to ICO
button_convert = tk.Button(
    frame_convert_quit_buttons,
    text="Convert to ICO",
    command=convert_to_icon,
)
button_convert.grid(row=0, column=0, padx=5, sticky="s")

# CONTENT FRAME: CONVERT QUIT BUTTONS FRAME: Button: Quit Application
button_quit = tk.Button(
    frame_convert_quit_buttons,
    text="Quit",
    command=quit_application,
    fg="#9C0006",
    bg="#FFC7CE",
)
button_quit.grid(row=0, column=1, padx=5, sticky="s")


# STATUS BAR FRAME: Label: Version Label
label_version = tk.Label(
    frame_status_bar,
    text="Iconizer v1.0",
    anchor="w",
    bg="lightgray",
    fg="black",
    padx=5,
)
label_version.grid(row=0, column=0, pady=0, sticky="w")

# STATUS BAR FRAME: Label: Selected Images
label_selected_images = tk.Label(
    frame_status_bar,
    text="0 Images Selected",
    anchor="w",
    bg="lightgray",
    fg="black",
    padx=5,
)
label_selected_images.grid(row=0, column=1, pady=0, sticky="w")

# STATUS BAR FRAME: Label: Status Message
label_status_message = tk.Label(
    frame_status_bar,
    anchor="w",
    bg="lightgray",
    fg="black",
    padx=5,
)
label_status_message.grid(row=0, column=2, pady=0, sticky="ew")

# Add separators between all children of the status bar frame
children = frame_status_bar.winfo_children()
for i in range(len(children)):
    # Move the current child to column 2 * i + 1
    child = children[i]
    child.grid_forget()
    child.grid(row=0, column=2 * i + 1, pady=5, sticky="s")

    if i < len(children) - 1:
        # Add the separator between children
        separator = ttk.Separator(frame_status_bar, orient="vertical")
        separator.grid(row=0, column=2 * i + 2, pady=5, sticky="ns")

# Configure Main Frame
root.grid_rowconfigure(0, weight=1)  #                Frame:  Main
root.grid_columnconfigure(0, weight=1)  #             Frame:  Main

# Configure grid inside the main frame
frame_main.grid_columnconfigure(0, weight=1)
frame_main.grid_rowconfigure(0, weight=1)  #          Frame:  Content
frame_main.grid_rowconfigure(1, weight=0)  #          Frame:  Statusbar

# Configure grid inside the content frame
frame_content.grid_columnconfigure(0, weight=1)
frame_content.grid_rowconfigure(0, weight=0)  #       Label:  Title
frame_content.grid_rowconfigure(1, weight=0)  #       Frame:  Selection Buttons
frame_content.grid_rowconfigure(2, weight=0)  #       Text:   Filepath
frame_content.grid_rowconfigure(3, weight=1)  #       Label:  Image preview
frame_content.grid_rowconfigure(4, weight=0)  #       Frame:  Convert / Quit Buttons

# Configure grid inside the status bar frame
frame_status_bar.grid_columnconfigure(0, weight=0)  # Label:  Version Number
frame_status_bar.grid_columnconfigure(1, weight=0)  # Label:  Selected Images
frame_status_bar.grid_columnconfigure(2, weight=0)  # Label:  Status Message


def generate_color_palette(base_hue, hue_range, colors_amount, iteration):
    brightness = max(0.3, 1.0 - (iteration * 0.15))  # Reduce brightness per level
    saturation = 0.8
    step = hue_range / (colors_amount + 1)  # Divide into (x+1) parts
    color_palette = []

    print(f"Base Hue: {base_hue}")
    print(f"Hue Range: {hue_range}")
    print(f"Generated Color Palette (Iteration {iteration}):")

    for i in range(colors_amount):
        hue = base_hue + (i + 1) * step  # Skip the first slot (reserved for parent)
        hue %= 1.0  # Ensure valid hue range

        # Print each color's calculated range and the resulting hue
        color_range_start = base_hue + (i * step)
        color_range_end = base_hue + ((i + 1) * step)
        color_range_start %= 1.0
        color_range_end %= 1.0

        print(
            f"Color {i + 1}: Range [{color_range_start:.3f}, {color_range_end:.3f}] -> Hue: {hue:.3f}"
        )

        r, g, b = hls_to_rgb(hue, brightness, saturation)
        color_palette.append(f"#{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}")

    return color_palette


def colorize_widgets(widget, base_hue=0.0, hue_range=1.0, iteration=0):
    children = widget.winfo_children()
    if not children:
        return  # Stop if there are no children

    # Parent color (first in range)
    r, g, b = hls_to_rgb(base_hue, max(0.3, 1.0 - (iteration * 0.15)), 0.8)
    widget.config(bg=f"#{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}")

    # Generate child colors
    color_palette = generate_color_palette(
        base_hue, hue_range, len(children), iteration
    )
    print(f"color_palette: {color_palette}")

    # Apply colors and recurse
    for i, child in enumerate(children):
        if isinstance(child, (tk.Frame, tk.Label, tk.Button, tk.Entry, tk.Text)):
            child.config(bg=color_palette[i])  # Use generated color
        else:
            print(f"instance ignored {i}")

        new_base_hue = base_hue + ((i + 1) * (hue_range / (len(children) + 1)))
        new_hue_range = hue_range / (len(children) + 1)  # Subdivide range
        colorize_widgets(child, new_base_hue, new_hue_range, iteration + 1)


debug_mode = False
if debug_mode:
    root.config(bg="#888888")  # Neutral gray for root
    colorize_widgets(root)


# Start the Tkinter main loop
root.mainloop()
