import customtkinter as ctk
import tkinter as tk
import math
import random
try:
    from ctypes import windll, byref, sizeof, c_int
except:
    pass

THEME_TOP = "#33475d" # R G B
PANEL_TOP = 0x005d4733  # B G R

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Custom Weighted Spinner")
        self.geometry("620x450")
        
        # Initial Data
        self.data = [{"name": "Pizza", "percent": 50}, {"name": "Tacos", "percent": 50}]
        
        # Layout
        self.input_panel = InputPanel(self, self.update_wheel)
        self.input_panel.pack(side="left", fill="y", padx=10, pady=10)
        
        self.wheel_spinner = WheelSpinner(self, data=self.data)
        self.wheel_spinner.pack(side="right", expand=True, fill="both", padx=10, pady=10)

    def update_wheel(self, new_data):
        self.wheel_spinner.update_data(new_data)

    # Aesthetics 
    def ChangeTitleBar(self) -> None:
        '''Sync title bar colour (windows only)'''
        try:
            HWND = windll.user32.GetParent(self.winfo_id())
            DWMWA_ATTRIBUTE = 35
            COLOR = PANEL_TOP
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_ATTRIBUTE, byref(c_int(COLOR)), sizeof(c_int))
        except:
            pass

class WheelSpinner(ctk.CTkFrame):
    def __init__(self, parent, data, **kwargs):
        super().__init__(parent, **kwargs)
        self.data = data
        self.current_angle = 0
        self.is_spinning = False
        
        # UI Elements
        self.canvas = tk.Canvas(self, width=300, height=300, 
                                bg="#2b2b2b", highlightthickness=0)
        self.canvas.pack(pady=10, padx=10)
        
        self.spin_button = ctk.CTkButton(self, text="SPIN!", height=35, 
                                         font=("Arial", 14, "bold"), 
                                         command=self.start_spin)
        self.spin_button.pack(pady=5)

        # NEW: The Result Label
        self.result_label = ctk.CTkLabel(self, text="", 
                                         font=("Arial", 18, "bold"),
                                         text_color="#FFCC00") # Gold color for the winner
        self.result_label.pack(pady=10)
        
        self.setup_ghost_button()
        self.draw_wheel()

    def update_data(self, new_data):
        self.data = new_data
        self.draw_wheel(self.current_angle)

    def draw_wheel(self, offset=0):
        self.canvas.delete("all")
        colors = ["#FF5733", "#33FF57", "#3357FF", "#F333FF", "#FF33A1", "#F3FF33"]
        start_angle = offset
        for i, item in enumerate(self.data):
            extent = (item['percent'] / 100) * 360
            color = colors[i % len(colors)]
            self.canvas.create_arc(5, 5, 295, 295, start=start_angle, 
                                   extent=extent, fill=color, outline="white")
            
            mid_angle = math.radians(start_angle + extent/2)
            x = 150 + 80 * math.cos(mid_angle)
            y = 150 - 80 * math.sin(mid_angle)
            
            self.canvas.create_text(x, y, text=item['name'], fill="white", 
                                    font=("Arial", 10, "bold"), angle=start_angle + extent/2)
            start_angle += extent
        
        self.canvas.create_polygon(285, 150, 300, 140, 300, 160, fill="#FFCC00")

    def start_spin(self):
        if not self.is_spinning and self.data:
            self.result_label.configure(text="Spinning...")
            self.is_spinning = True

            if self.rigged_index is not None:
                # Calculate angles based on the current list
                current_sum = 0
                for i in range(self.rigged_index):
                    current_sum += (self.data[i]['percent'] / 100) * 360
                
                extent = (self.data[self.rigged_index]['percent'] / 100) * 360
                # Target the exact middle of that slice
                target_pos = 360 - (current_sum + extent / 2)
                
                # Make it spin at least 5 times for effect
                total_needed = (360 * 5) + (target_pos - (self.current_angle % 360))
                self.animate_rigged(total_needed)
            else:
                self.animate(random.uniform(15, 25))

    def animate_rigged(self, remaining):
        if remaining > 0.5:
            # This creates a "smooth" deceleration feel
            step = max(remaining * 0.04, 0.2) 
            self.current_angle += step
            self.draw_wheel(self.current_angle)
            self.after(10, lambda: self.animate_rigged(remaining - step))
        else:
            self.is_spinning = False
            self.get_winner()
            self.rigged_index = None # Reset after one rigged use

    def animate(self, velocity):
        if velocity > 0.1:
            self.current_angle += velocity
            self.draw_wheel(self.current_angle)
            self.after(10, lambda: self.animate(velocity * 0.97))
        else:
            self.is_spinning = False
            self.get_winner()

    def setup_ghost_button(self):
        # A tiny button in the bottom right, invisible to others
        self.ghost_btn = ctk.CTkButton(
            self, text="", width=10, height=10, 
            fg_color="transparent", hover_color="#2b2b2b", # Match your bg
            command=self.open_backstage_menu
        )
        self.ghost_btn.place(relx=0.98, rely=0.98, anchor="center")

    def open_backstage_menu(self):
        # Create a small popup window
        top = ctk.CTkToplevel(self)
        top.title("Backstage")
        top.geometry("200x150")
        top.attributes("-topmost", True) # Keep it on top

        ctk.CTkLabel(top, text="Set Forced Winner:").pack(pady=10)
        
        # Dropdown or Entry for the index
        options = [f"{i}: {item['name']}" for i, item in enumerate(self.data)]
        selector = ctk.CTkOptionMenu(top, values=options, 
                                     command=lambda val: self.set_rig(val))
        selector.pack(pady=5)
        
        ctk.CTkButton(top, text="Clear Rig", command=lambda: self.set_rig(None, top)).pack(pady=10)

    def set_rig(self, value, window=None):
        if value is None:
            self.rigged_index = None
        else:
            self.rigged_index = int(value.split(":")[0])
        print(f"Rigged to index: {self.rigged_index}")

    def get_winner(self):
        normalized_angle = (360 - (self.current_angle % 360)) % 360
        current_sum = 0
        winner_name = "Unknown"
        
        for item in self.data:
            extent = (item['percent'] / 100) * 360
            if current_sum <= normalized_angle < current_sum + extent:
                winner_name = item['name']
                break
            current_sum += extent
            
        # Update the UI Label instead of just printing
        self.result_label.configure(text=f"Winner: {winner_name}!")
        
class InputPanel(ctk.CTkFrame):
    def __init__(self, parent, sync_callback):
        super().__init__(parent, width=220) # Narrower sidebar
        self.sync_callback = sync_callback
        
        ctk.CTkLabel(self, text="Wheel Items", font=("Arial", 16, "bold")).pack(pady=5)
        
        # Shorter scrollable frame
        self.entries_frame = ctk.CTkScrollableFrame(self, width=200, height=200)
        self.entries_frame.pack(pady=5, padx=5, fill="both", expand=True)
        
        self.rows = []
        self.add_btn = ctk.CTkButton(self, text="+ Add", height=28, command=self.add_row)
        self.add_btn.pack(pady=2)

        self.status_label = ctk.CTkLabel(self, text="Total: 0%", font=("Arial", 11))
        self.status_label.pack()

        self.update_btn = ctk.CTkButton(self, text="Update", height=32, command=self.sync)
        self.update_btn.pack(pady=10)
        
        for _ in range(2): self.add_row()

    def add_row(self):
        # Container for the single row
        row_frame = ctk.CTkFrame(self.entries_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=2)

        name_entry = ctk.CTkEntry(row_frame, placeholder_text="Name", width=110)
        name_entry.pack(side="left", padx=2)

        perc_entry = ctk.CTkEntry(row_frame, placeholder_text="%", width=50)
        perc_entry.pack(side="left", padx=2)

        # The Delete Button
        remove_btn = ctk.CTkButton(row_frame, text="✕", width=30, fg_color="#a13232", 
                                   hover_color="#7d2626", 
                                   command=lambda r=row_frame: self.remove_row(r))
        remove_btn.pack(side="left", padx=2)

        # Store references so we can extract data later
        self.rows.append({
            "frame": row_frame,
            "name": name_entry,
            "percent": perc_entry
        })

    def remove_row(self, frame_to_remove):
        # 1. Remove from the internal list
        self.rows = [row for row in self.rows if row["frame"] != frame_to_remove]
        # 2. Destroy the physical widgets
        frame_to_remove.destroy()
        # 3. Recalculate status
        self.update_status()

    def update_status(self):
        total = 0
        for row in self.rows:
            try:
                total += float(row["percent"].get())
            except: continue
        
        self.status_label.configure(text=f"Total: {total}%", 
                                    text_color="white" if total == 100 else "#e67e22")

    def sync(self):
        new_data = []
        total_p = 0
        
        for row in self.rows:
            name = row["name"].get()
            try:
                p = float(row["percent"].get())
                if name:
                    new_data.append({"name": name, "percent": p})
                    total_p += p
            except ValueError:
                continue
        
        if not new_data: return

        # Auto-normalization if they don't hit 100
        if total_p != 100 and total_p > 0:
            for item in new_data:
                item['percent'] = (item['percent'] / total_p) * 100
            
        self.update_status()
        self.sync_callback(new_data)

if __name__ == "__main__":
    app = App()
    app.mainloop()