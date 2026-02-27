#region IMPORTS
import customtkinter as ctk
import yfinance as yf
import threading
import json 
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

        self.exchange_rate = 1.0
        self.last_prices = (None, None)

        # widgets
        self.CreateFrames()
        self.LoadData()

        # detection
        self.protocol("WM_DELETE_WINDOW", self.OnClose)
    
    def CreateFrames(self):
        '''Adds frame widgets onto window'''
        self.control_frame = ControlFrame(self, self.AddRowCallback, self.UpdateCallback, self.ResetCallback, self.ToggleCallback)
        self.control_frame.place(relx = 0, rely = 0, relwidth = 1.0, relheight = 0.15)

        self.main_frame = MainFrame(self)
        self.main_frame.place(relx = 0, rely = 0.15, relwidth = 1.0, relheight = 0.7)

        self.summary_frame = SummaryFrame(self)
        self.summary_frame.place(relx = 0, rely = 0.85, relwidth = 1.0, relheight = 0.15)

    # CALLBACK FUNCTIONS
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

    def ToggleCallback(self) -> None:
        '''Called when switch is flipped or for fresh data'''
        if self.last_prices is None: return

        is_nzd = self.control_frame.currency_var.get() == "NZD"
        multiplier = self.exchange_rate if is_nzd else 1.0

        self.ApplyPricesToUI(self.last_prices[0], self.last_prices[1], multiplier)

    def UpdateCallback(self) -> None:
        '''Triggered when tickers need updating'''
        tickers = [row["ticker"].get().strip().upper() for row in self.main_frame.rows_data if row["ticker"].get()]
        if not tickers: 
            return
        
        threading.Thread(target = self.FetchPrices, args = (tickers,), daemon = True).start()

    def FetchPrices(self, tickers: list) -> None:
        '''Background task to fetch data and update UI'''
        try:
            # retrieves most recent data
            data = yf.download(tickers + ["NZD=X"], period = "2d", interval = "1d", progress = False, prepost = True)
            close_data = data['Close'].ffill() 

            # 2. Assign the rows
            self.last_prices = (close_data.iloc[0], close_data.iloc[-1])
            self.exchange_rate = float(self.last_prices[1]["NZD=X"])
            self.after(0, lambda: self.ToggleCallback())
        except Exception as e:
            print(f"Error fetching data: {e}")
    
    def ApplyPricesToUI(self, prev_prices: list, current_prices: list, multiplier: float = 1.0) -> None:
        '''Pushes values onto associated widgets'''
        total_value = 0.0
        total_change = 0.0
        currency_sym = "NZ$" if multiplier != 1.0 else "$"

        for row in self.main_frame.rows_data:
            ticker = row["ticker"].get().strip().upper()

            if ticker in current_prices:
                # gathers and displays appropriate data
                price = current_prices[ticker] * multiplier
                prev_close = prev_prices[ticker] * multiplier
                quantity = float(row["amount"].get() or 0)

                change = ((price - prev_close) / prev_close) * 100
                quantity_change = (price - prev_close) * quantity
                colour = "#0F9D58" if quantity_change >= 0 else "#DB4437"
                prefix = "+" if change >= 0 else "-"

                # update row info
                row_total = price * quantity
                row["price"].configure(text = f"{currency_sym}{price:,.2f}")
                row["total"].configure(text = f"{currency_sym}{row_total:,.2f}")
                row["change"].configure(text = f"{prefix}{currency_sym}{abs(quantity_change):,.2f} ({change:+.2f}%)", text_color = colour)
                total_value += row_total
                total_change += (price - prev_close) * quantity

            self.summary_frame.UpdateSummary(total_value, total_change)

    # DATA FUNCTIONS
    def OnClose(self) -> None:
        '''Executes when application is closed'''
        self.SaveData()
        self.destroy()

    def SaveData(self) -> None:
        '''Extract tickers and quantity'''
        saved_data = []
        for row in self.main_frame.rows_data:
            saved_data.append({
                "ticker": row["ticker"].get(),
                "amount": row["amount"].get() 
            })
        with open("portfolio.json", "w") as file:
            json.dump(saved_data, file, indent = 4)

    def LoadData(self) -> None:
        '''Extracts saved data, loading it'''
        try:
            with open("portfolio.json", "r") as file:
                saved_data = json.load(file)

                # feeds data using existing architecture
                for item in saved_data:
                    self.main_frame.AddRow()

                    new_row = self.main_frame.rows_data[-1]
                    new_row["ticker"].insert(0, item["ticker"])
                    new_row["amount"].insert(0, item["amount"])
            
            self.UpdateCallback()
        except FileNotFoundError:
            pass # no current file save

    def ChangeTitleBar(self) -> None:
        '''Sync title bar colour (windows only)'''
        try:
            HWND = windll.user32.GetParent(self.winfo_id())
            DWMWA_ATTRIBUTE = 35
            COLOR = 0x003c2319
            windll.dwmapi.DwmSetWindowAttribute(HWND, DWMWA_ATTRIBUTE, byref(c_int(COLOR)), sizeof(c_int))
        except:
            pass

#region FRAMES
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

        self.grid_columnconfigure((1, 3, 5), weight = 1) #ticker and quantity have expansive nature
        self.grid_columnconfigure((0, 2, 4, 6), weight = 0)
        self.rows_data = []
    
    def AddRow(self) -> None:
        '''Allows the addition of a new row'''
        vcmd = (self.register(self.ValidateNumber), '%P') 
        index = len(self.rows_data)
        row_dictionary = {}

        row_dictionary["num"] = ctk.CTkLabel(self, text = f"{index + 1}", width = 30)
        row_dictionary["num"].grid(row = index, column = 0, padx = 5, pady = 5)

        row_dictionary["ticker"] = ctk.CTkEntry(self, placeholder_text = "Ticker", width = 70)
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
    def __init__(self, parent, add_command: function, update_command: function, reset_command: function, toggle_command: function, **kwargs):
        super().__init__(parent, fg_color = "transparent", **kwargs)

        # Setup
        self.currency_var = ctk.StringVar(value="USD")
        self.grid_columnconfigure((0, 1, 2, 3), weight = 1)
        self.grid_rowconfigure(0, weight = 1)

        # setting widgets
        self.button_add = ctk.CTkButton(
            self, text = "+ Add Row", fg_color = BTN_BLUE, hover_color = BTN_HOVER, 
            command = add_command  
        )
        self.button_add.grid(row = 0, column = 0, padx = 5)

        self.button_update = ctk.CTkButton(
            self, text = "Update", fg_color = BTN_BLUE, hover_color = BTN_HOVER, 
            command = update_command
        )
        self.button_update.grid(row = 0, column = 1, padx = 5)

        self.switch_currency = ctk.CTkSwitch(
            self, 
            text = "NZD/USD",
            variable = self.currency_var, 
            command = toggle_command,
            progress_color = BTN_BLUE, 
            onvalue = "NZD", offvalue = "USD",
            text_color = "white"
        )
        self.switch_currency.grid(row = 0, column = 2, padx = 5)

        self.button_reset = ctk.CTkButton(
            self, text = "Reset", fg_color = BTN_RED, hover_color = "#8a2424", width = 80,
            command = reset_command
        )
        self.button_reset.grid(row = 0, column = 3, padx = 5)
    
#endregion

if __name__ == "__main__":
    app = App()
    app.mainloop()