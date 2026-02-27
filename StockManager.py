#region imports
import customtkinter as ctk
import yfinance as yf
import threading
try:
    from ctypes import windll, byref, sizeof, c_int
except:
    pass

# COLOURS
THEME_BG = "#19233c"        # Main Background
THEME_DARK = "#121a2d"      # Darker Background (Summary)
BTN_BLUE = "#1f538d"        # Standard Button
BTN_RED = "#C53434"         # Reset/Delete Button
BTN_HOVER = "#14375e"
#endregion

class App(ctk.CTk):
    def __init__(self):
        # setup
        super().__init__(fg_color = THEME_BG)
        self.title("Stock Manager")
        self.geometry("600x500")
        self.minsize(500,400)
        self.ChangeTitleBar()

        # widgets
        self.CreateFrames()
    
    def CreateFrames(self):
        '''Adds frame widgets onto window'''
        self.control_frame = ControlFrame(self, self.AddRowCallback, self.UpdateCallback, self.ResetCallback)
        self.control_frame.place(relx = 0, rely = 0, relwidth = 1.0, relheight = 0.15)

        self.main_frame = MainFrame(self)
        self.main_frame.place(relx = 0, rely = 0.15, relwidth = 1.0, relheight = 0.7)

        self.summary_frame = SummaryFrame(self)
        self.summary_frame.place(relx = 0, rely = 0.85, relwidth = 1.0, relheight = 0.15)

    # callback triggers for control frame buttons
    def AddRowCallback(self) -> None:
        '''Triggered when new row required'''
        self.main_frame.AddRow()

    def ResetCallback(self) -> None:
        '''Triggered to clear all rows'''
        for row_dict in self.main_frame.rows_data:
            for widget in row_dict.values():
                widget.destroy()
        
        self.main_frame.rows_data.clear()

        # Only one layout recalculation needed
        self.main_frame.RefreshGrid() 
    
    def UpdateCallback(self) -> None:
        '''Triggered when tickers need updating'''
        tickers = [row["ticker"].get().strip().upper() for row in self.main_frame.rows_data if row["ticker"].get()]
        if not tickers: 
            return
        
        threading.Thread(target = self.FetchPrices, args = (tickers,), daemon = True).start()

    def FetchPrices(self, tickers: list) -> list:
        '''Background task to fetch data and update UI'''
        try:
            # retrieves most recent data
            data = yf.download(tickers, period = "2d", interval = "1d", progress = False, prepost = True)
            previous_prices = data['Close'].iloc[0]
            current_prices = data['Close'].iloc[-1]
            self.after(0, lambda: self.ApplyPricesToUI(previous_prices, current_prices))
        except Exception as e:
            print(f"Error fetching data: {e}")
    
    def ApplyPricesToUI(self, prev_prices: list, current_prices: list) -> None:
        '''Pushes values onto associated widgets'''
        total_value = 0.0
        total_change = 0.0

        for row in self.main_frame.rows_data:
            ticker = row["ticker"].get().strip().upper()

            if ticker in current_prices:
                # gathers and displays appropriate data
                price = current_prices[ticker]
                prev_close = prev_prices[ticker]
                quantity = float(row["amount"].get() or 0)

                change = ((price - prev_close) / prev_close) * 100
                quantity_change = (price - prev_close) * quantity
                colour = "#0F9D58" if quantity_change >= 0 else "#DB4437"
                prefix = "+" if change >= 0 else ""

                # update row info
                row_total = price * quantity
                row["price"].configure(text = f"${price:,.2f}")
                row["total"].configure(text = f"${row_total:,.2f}")
                row["change"].configure(text = f"{prefix}${quantity_change:+.2f} ({change:+.2f}%)", text_color = colour)
                total_value += row_total
                total_change += (price - prev_close) * quantity
            else:
                row["total"].configure(text = "NULL")

            self.summary_frame.UpdateSummary(total_value, total_change)

    def ChangeTitleBar(self):
        '''Sync title bar colour (windows only)'''
        try:
            HWND = windll.user32.GetParent(self.winfo_id())
            DWMWA_ATTRIBUTE = 35
            COLOR = 0x003c2319
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_ATTRIBUTE, byref(c_int(COLOR)), sizeof(c_int))
        except:
            pass

class SummaryFrame(ctk.CTkFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color = THEME_DARK, **kwargs)

        self.total_label = ctk.CTkLabel(
            self,
            text = "Total: $0.00",
            font = ("Helvetica", 20, "bold"),
            text_color = "white"
        )
        self.total_label.pack(expand = True, pady = (5, 0))

        self.change_label = ctk.CTkLabel(
            self,
            text = "$Day change: $0.00 (0.00%)",
            font = ("Helvetica", 12, "bold"),
            text_color = "gray"
        )
        self.change_label.pack(expand = True, pady = (0, 10)) 
    
    def UpdateSummary(self, total_amount: float, change: float):
        '''Method to modify total and conigure profit/loss'''
        self.total_label.configure(text = f"Total: ${total_amount:,.2f}")

        # change label config
        colour = "#0F9D58" if change >= 0 else "#DB4437"
        prefix = "+" if change >= 0 else ""
        percent = (change / (total_amount - change) * 100) if total_amount != change else 0
        self.change_label.configure(
            text = f"Day Change: {prefix}${change:,.2f} ({prefix}{percent:.2f}%)", 
            text_color = colour
        )

class MainFrame(ctk.CTkScrollableFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color = "transparent", corner_radius = 0, **kwargs)

        self.grid_columnconfigure((1, 3), weight = 1) #ticker and quantity have expansive nature
        self.grid_columnconfigure((0, 2, 4, 5, 6), weight = 0)
        self.rows_data = []
    
    def AddRow(self) -> None:
        '''Allows the addition of a new row'''
        vcmd = (self.register(self.ValidateNumber), '%P') 
        index = len(self.rows_data)
        row_dictionary = {}

        row_dictionary["num"] = ctk.CTkLabel(self, text = f"{index + 1}", width = 30)
        row_dictionary["num"].grid(row = index, column = 0, padx = 5, pady = 5)

        row_dictionary["ticker"] = ctk.CTkEntry(self, placeholder_text = "Ticker", width = 120)
        row_dictionary["ticker"].grid(row=index, column = 1, padx = 5, pady = 5, sticky = "ew")

        row_dictionary["price"] = ctk.CTkLabel(self, text = "$0.00", width = 80)
        row_dictionary["price"].grid(row = index, column = 2, padx = 5, pady = 5)

        row_dictionary["amount"] = ctk.CTkEntry(
            self, 
            placeholder_text = "0.0", 
            width = 80,
            validate = "key",
            validatecommand = vcmd
        )
        row_dictionary["amount"].grid(row = index, column = 3, padx = 5, pady = 5, sticky = "ew")

        row_dictionary["total"] = ctk.CTkLabel(self, text = "$0.00", width = 80)
        row_dictionary["total"].grid(row = index, column = 4, padx = 5, pady = 5)

        row_dictionary["change"] = ctk.CTkLabel(self, text = "0.00%", width = 70, text_color = "gray")
        row_dictionary["change"].grid(row = index, column = 5, padx = 5, pady = 5)

        row_dictionary["delete"] = ctk.CTkButton(
            self, 
            text = "-", 
            width = 30, 
            fg_color = BTN_BLUE,
            command = lambda: self.RemoveRow(row_dictionary)
        )
        row_dictionary["delete"].grid(row = index, column = 6, padx = 5, pady = 5)

        # Store references so we can access their data later
        self.rows_data.append(row_dictionary)
    
    def RemoveRow(self, row_dict: dict) -> None:
        '''Allows the removal of an existing row'''
        for widget in row_dict.values():
            widget.destroy()
        self.rows_data.remove(row_dict)
        self.RefreshGrid()

    def RefreshGrid(self) -> None:
        '''Re-indexes and moves data after deletion'''
        for index, row in enumerate(self.rows_data): 
            row["num"].grid(row = index, column = 0)
            row["ticker"].grid(row = index, column = 1)
            row["price"].grid(row = index, column = 2)
            row["amount"].grid(row = index, column = 3)
            row["total"].grid(row = index, column = 4)
            row["delete"].grid(row = index, column = 5)
            
            # Update the visual row number
            row["num"].configure(text = str(index + 1))

    def ValidateNumber(self, input: str) -> bool:
        '''Only allows numbers and a single decimal place'''
        if input == "":
            # allows deletion
            return True
        try:
            # verifies if string can become a float
            float(input)
            return True
        except ValueError:
            return False

class ControlFrame(ctk.CTkFrame):
    def __init__(self, parent, add_command: function, update_command: function, reset_command: function, **kwargs):
        super().__init__(parent, fg_color = "transparent", **kwargs)

        # Setup Grid for even spacing
        self.grid_columnconfigure((0, 1, 2, 3), weight = 1)
        self.grid_rowconfigure(0, weight = 1)

        # settings widgets
        self.button_add = ctk.CTkButton(
            self, text = "+ Add Row", fg_color = BTN_BLUE, hover_color = BTN_HOVER, 
            command = add_command  
        )
        self.button_add.grid(row = 0, column = 0, padx = 5)

        # 2. Update Button
        self.button_update = ctk.CTkButton(
            self, text = "Update", fg_color = BTN_BLUE, hover_color = BTN_HOVER, 
            command = update_command
        )
        self.button_update.grid(row = 0, column = 1, padx = 5)

        # 3. NZD Switch
        self.switch_currency = ctk.CTkSwitch(
            self, text = "NZD", progress_color = BTN_BLUE, text_color = "white"
        )
        self.switch_currency.grid(row = 0, column = 2, padx = 5)

        # 4. Reset Button
        self.button_reset = ctk.CTkButton(
            self, text = "Reset", fg_color = BTN_RED, hover_color = "#8a2424", width = 80,
            command = reset_command
        )
        self.button_reset.grid(row = 0, column = 3, padx = 5)

if __name__ == "__main__":
    app = App()
    app.mainloop()