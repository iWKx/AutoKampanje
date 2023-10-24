# Import necessary libraries
import tkinter as tk
import pygetwindow as gw
import json
from PIL import Image, ImageTk

# File paths
logo_path = "./data/pics/logo.ico"
con_logo_path = "./data/pics/pwr.png"
json_path = "./data/config.json"

# Window dimensions
window_x = 480
window_y = 620

#functions
def get_active_window_titles():
    try:
        active_windows = gw.getAllTitles()
        active_windows = [title for title in active_windows if title]
        return active_windows or ["No active windows found."]
    except Exception as e:
        return [f"An error occurred: {str(e)}"]

def auto_select_shoppa_window():
    active_windows = get_active_window_titles()
    shoppa_windows = [title for title in active_windows if "shoppa" in title.lower()]

    if shoppa_windows:
        window_list.selection_clear(0, tk.END)
        index = active_windows.index(shoppa_windows[0])
        window_list.selection_set(index)
        on_select_button_click()
    else:
        target_program_title_entry.delete(0, tk.END)
        target_program_title_entry.insert(0, "Start Shoppa først")
        target_program_title_entry.config(fg="#fc032d", highlightbackground="#fc032d")

def update_window_list():
    active_windows = get_active_window_titles()
    shoppa_windows = [title for title in active_windows if "shoppa" in title.lower()]

    window_list.delete(0, tk.END)
    for title in shoppa_windows:
        window_list.insert(tk.END, title)

def refresh_window_list():
    window_list.delete(0, tk.END)
    titles = get_active_window_titles()
    for title in titles:
        window_list.insert(tk.END, title)

def on_select_button_click():
    selected_item = window_list.get(window_list.curselection())
    target_program_title_entry.delete(0, tk.END)
    target_program_title_entry.insert(0, selected_item)

    if "shoppa" in selected_item.lower():
        target_program_title_entry.config(fg="#03fc94", highlightbackground="#03fc94")
    else:
        target_program_title_entry.config(fg="#fc032d", highlightbackground="#fc032d")

def load_config():
    try:
        with open(json_path, 'r') as config_file:
            return json.load(config_file)
    except FileNotFoundError:
        return {}

def save_config():
    existing_config = load_config()
    existing_config["target_program_title"] = target_program_title_entry.get()
    existing_config["power_butikk"] = power_butikk_entry.get()
    existing_config["delay_start"] = float(delay_start_entry.get())
    existing_config["pause_autogui"] = float(delay_autogui_entry.get())

    with open(json_path, 'r') as config_file:
        data = json.load(config_file)

    data.update(existing_config)
    with open(json_path, 'w') as config_file:
        json.dump(data, config_file, indent=2)
    con.destroy()


def on_enter(event):
    widget = event.widget
    if widget['state'] == tk.NORMAL:
        widget.configure(background="white")

def on_leave(event):
    widget = event.widget
    widget.configure(background="#f15c25")

def on_enter2(event):
    widget = event.widget
    if widget['state'] == tk.NORMAL:
        widget.configure(fg="#f15c25")

def on_leave2(event):
    widget = event.widget
    widget.configure(fg="white")

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")

# Create the config tkinter window
con = tk.Tk()
con.title("Setup")
con.configure(bg="#2b2e33")
con.iconbitmap(logo_path)
con.attributes("-alpha", 0.98)
window_width = window_x
window_height = window_y
center_window(con, window_width, window_height)
con.resizable(False, True)

existing_config = load_config()
get_active_window_titles()

# Create the logo image and label
logo_image = Image.open(con_logo_path)
logo_image = logo_image.resize((350, 100))
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(con, image=logo_photo, bg="#2b2e33")
logo_label.image = logo_photo
logo_label.pack(pady=(10, 0))


# Create widgets and buttons
intro_felt = tk.Label(con, text="Config", fg="white", bg="#2b2e33", font=("Roboto", 10, "bold italic"))
intro_felt.pack(pady=(0, 10))
intro_felt.bind("<Enter>", on_enter2)
intro_felt.bind("<Leave>", on_leave2)


window_list = tk.Listbox(con, selectmode=tk.SINGLE, height=2)
window_list.configure(bg="#2b2e33", fg="white", highlightcolor="#f15c25", highlightbackground="#f7b154",
                     selectbackground="#f15c25", selectforeground="black", font=("Verdana", 10))
window_list.pack(padx=55, pady=(10, 10),ipady=5, fill=tk.BOTH, expand=False)
window_list.configure(highlightthickness=1, borderwidth=1, relief=tk.RIDGE)

button_frame = tk.Frame(con, bg="#2b2e33")
button_frame.pack()

copy_button = tk.Button(button_frame, text="Select", command=on_select_button_click)
copy_button.configure(cursor="hand2", fg="white", bg="#f15c25", font=("Verdana", 10), width=10, relief=tk.RAISED, borderwidth=0)
copy_button.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

refresh_button = tk.Button(button_frame, text="Refresh", command=refresh_window_list)
refresh_button.configure(cursor="hand2", fg="white", bg="grey", font=("Verdana", 10), width=10, relief=tk.RAISED, borderwidth=0)
refresh_button.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)

refresh_window_list()

tk.Label(con, text="Target Program Title:", fg="white", bg="#2b2e33", font=("Verdana", 10)).pack(pady=(10, 1))
target_program_title_entry = tk.Entry(con)
target_program_title_entry.config(fg="#03fc98", font=("Verdana", 9), width=45, justify="center", bg="#202226")
target_program_title_entry.pack(pady=(5, 15),ipady=10)
target_program_title_entry.configure(highlightthickness=1, borderwidth=1, relief=tk.RIDGE, highlightcolor="#f15c25")

tk.Label(con, text="Power Butikk:", fg="white", bg="#2b2e33", font=("Verdana", 10)).pack(pady=(5, 0))
power_butikk_entry = tk.Entry(con)
power_butikk_entry.insert(0, existing_config.get("power_butikk", ""))
power_butikk_entry.config(justify="center", fg="#f15c25", bg="#202226", font=("Verdana", 11), width=22)
power_butikk_entry.pack(pady=(5,15),ipady=7)
power_butikk_entry.configure(highlightthickness=1, borderwidth=1, relief=tk.RIDGE, highlightcolor="#f15c25", highlightbackground="#18191c")

delay_frame = tk.Frame(con, bg="#2b2e33")
delay_frame.pack()

tk.Label(delay_frame, text="Delay Start:", fg="white", bg="#2b2e33", font=("Verdana", 10)).grid(row=0, column=0, padx=5, pady=(5, 0), sticky="w")
delay_start_entry = tk.Entry(delay_frame)
delay_start_entry.insert(0, existing_config.get("delay_start", ""))
delay_start_entry.config(justify="center", fg="#f15c25", bg="#202226", font=("Verdana", 11), width=5)
delay_start_entry.grid(row=1, column=0, padx=5, pady=(5, 15),ipady=3)
delay_start_entry.configure(highlightthickness=1, borderwidth=1, relief=tk.RIDGE, highlightcolor="#f15c25", highlightbackground="#18191c")

tk.Label(delay_frame, text="Delay:", fg="white", bg="#2b2e33", font=("Verdana", 10)).grid(row=0, column=1, padx=5, pady=(5, 0), sticky="w")
delay_autogui_entry = tk.Entry(delay_frame)
delay_autogui_entry.insert(0, existing_config.get("pause_autogui", ""))
delay_autogui_entry.config(justify="center", fg="#f15c25", bg="#202226", font=("Verdana", 11), width=5, highlightcolor="#f15c25")
delay_autogui_entry.grid(row=1, column=1, padx=5, pady=(5, 15),ipady=3)
delay_autogui_entry.configure(highlightthickness=1, borderwidth=1, relief=tk.RIDGE, highlightcolor="#f15c25", highlightbackground="#18191c")

submit_button = tk.Button(con, text="Save Config", bg="#f15c25", fg="black", font="Verdana", cursor="hand2", command=save_config, relief=tk.FLAT, borderwidth=2, width=20, height=2)
submit_button.pack(pady=15)
submit_button.bind("<Enter>", on_enter)
submit_button.bind("<Leave>", on_leave)

signature_label = tk.Label(con, text="Made with 🖤 by Wael for Power")
signature_label.configure(cursor="heart", bg="#2b2e33", fg="white", font=("Verdana", 10))
signature_label.pack(side="bottom", pady=(5, 5))
signature_label.bind("<Enter>", on_enter2)
signature_label.bind("<Leave>", on_leave2)

auto_select_shoppa_window()

con.mainloop()