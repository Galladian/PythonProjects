#region IMPORTS
import customtkinter as ctk
import yfinance as yf
import threading
import json 
from tksheet import Sheet
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
        self.control_frame = ControlFrame(self, self.AddRowCallback, self.UpdateCallback, self.ResetCallback, self.ToggleCallback, self.SortCallback)
        self.control_frame.place(relx = 0, rely = 0, relwidth = 1.0, relheight = 0.15)

        self.main_frame = MainFrame(self)
        self.main_frame.place(relx = 0, rely = 0.15, relwidth = 1.0, relheight = 0.7)

        self.summary_frame = SummaryFrame(self)
        self.summary_frame.place(relx = 0, rely = 0.85, relwidth = 1.0, relheight = 0.15)

    # CALLBACK FUNCTIONS
    def AddRowCallback(self) -> None:
        '''Triggered when new row required'''
        self.main_frame.AddRow()

    def SortCallback(self, metric: str) -> None:
        '''Called to sort data by specificed metric'''
        if not hasattr(self.main_frame, "raw_data") or not self.main_frame.raw_data:
            print("No data available to sort yet. Please update prices first.")
            return
        
        self.main_frame.SortData(metric)

    def ResetCallback(self) -> None:
        '''Triggered to clear all rows'''
        self.main_frame.sheet.set_sheet_data(data = [])
        self.main_frame.sheet.set_column_widths([80, 80, 80, 100, 130])
        self.main_frame.AddRow()
        self.main_frame.sheet.redraw()

    def ToggleCallback(self) -> None:
        '''Called when switch is flipped or for fresh data'''
        if self.last_prices is None: return

        is_nzd = self.control_frame.currency_var.get() == "NZD"
        multiplier = self.exchange_rate if is_nzd else 1.0

        self.ApplyPricesToUI(self.last_prices[0], self.last_prices[1], multiplier)

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
    
    def UpdateCallback(self) -> None:
        '''Retrieves data from sheet'''
        table_data = self.main_frame.GetTableData()
        tickers = [row[0].strip().upper() for row in table_data if row[0].strip()]
        
        if not tickers: return
        threading.Thread(target=self.FetchPrices, args=(tickers,), daemon=True).start()

    def ApplyPricesToUI(self, prev_prices, current_prices, multiplier=1.0) -> None:
        total_value = 0.0
        total_change = 0.0
        currency_sym = "NZ$" if multiplier != 1.0 else "$"
        
        table_data = self.main_frame.GetTableData()
        new_raw_data = [] # Temporary list to build the new state

        for idx, row in enumerate(table_data):
            ticker = row[0].strip().upper() 
            if ticker in current_prices:
                price = current_prices[ticker] * multiplier
                prev_close = prev_prices[ticker] * multiplier
                
                try:
                    quantity = float(row[2] or 0)
                except ValueError:
                    quantity = 0.0

                change_percent = ((price - prev_close) / prev_close) * 100
                quantity_change = (price - prev_close) * quantity
                row_total = price * quantity
                
                # Store everything as raw numbers AND pre-formatted strings
                new_raw_data.append({
                    'ticker': ticker,
                    'price': price,
                    'quantity': quantity,
                    'total': row_total,
                    'pct': change_percent,
                    'price_str': f"{currency_sym}{price:,.2f}",
                    'total_str': f"{currency_sym}{row_total:,.2f}",
                    'change_str': f"{'+' if quantity_change >= 0 else '-'}{currency_sym}{abs(quantity_change):,.2f} ({change_percent:+.2f}%)"
                })

                total_value += row_total
                total_change += quantity_change

        # Update MainFrame's data and refresh view
        self.main_frame.raw_data = new_raw_data
        self.main_frame.SyncSheetWithRaw()
        
        self.summary_frame.UpdateSummary(total_value, total_change)

    # Data persistence functions
    def OnClose(self) -> None:
        '''Executes when application is closed'''
        self.SaveData()
        self.destroy()

    def SaveData(self) -> None:
        '''Extract tickers and quantity from tksheet'''
        raw_data = self.main_frame.sheet.get_sheet_data()
        
        saved_data = []
        for row in raw_data:
            if row and len(row) > 0 and str(row[0]).strip():
                saved_data.append({
                    "ticker": str(row[0]).strip().upper(),
                    "amount": str(row[2]).strip()
                })
                
        with open("portfolio.json", "w") as file:
            json.dump(saved_data, file, indent = 4)

    def LoadData(self) -> None:
        '''Extracts saved data and populates tksheet'''
        try:
            with open("portfolio.json", "r") as file:
                saved_data = json.load(file)

            new_sheet_data = []
            for item in saved_data:
                # [Ticker, Price (R), Amount, Total (R), Change(R)]
                row = [
                    item["ticker"], 
                    "$0.00",        
                    item["amount"], 
                    "$0.00",        
                    "0.00%"        
                ]
                new_sheet_data.append(row)
            
            # Replace existing sheet data with the new list
            if new_sheet_data:
                self.main_frame.sheet.set_sheet_data(new_sheet_data)
                self.main_frame.sheet.redraw()
                
                self.UpdateCallback()
            else:
                self.ResetCallback()
                
        except (FileNotFoundError, json.JSONDecodeError, TypeError):
            # Called on empty file
            self.main_frame.AddRow()

    # Aesthetics 
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
        prefix_value = "+" if change >= 0 else "-"
        percent = (change / (total_amount - change) * 100) if total_amount != change else 0
        self.change_label.configure(
            text = f"Day Change: {prefix_value}${abs(change):,.2f} ({prefix}{percent:.2f}%)", 
            text_color = colour
        )

class MainFrame(ctk.CTkFrame):
    def __init__(self, parent):
        # setup
        super().__init__(parent, fg_color = THEME_DARK)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.raw_data = []

        # 0:Ticker, 1:Price, 2:Amount, 3:Total, 4:Change
        self.sheet = Sheet(
            self, 
            headers = ["Ticker", "Price", "Amount", "Total", "Daily Change"],
            empty_horizontal = 0, 
            empty_vertical = 0)
        self.sheet.grid(row = 0, column = 0, sticky = "nsew")
        self.ModifyUsage()

    # Functionality
    def AddRow(self) -> None:
        '''Insert a new row with default values'''
        self.sheet.insert_row(["", "$0.00", "", "$0.00", "0.00%"])
        self.sheet.redraw()

    def GetTableData(self) -> None:
        '''Returns all row data as a list of lists'''
        return self.sheet.get_sheet_data()

    def UpdateRow(self, row_idx, values_dict) -> None:
        '''Helper to update specific columns in a row'''
        if "price" in values_dict: self.sheet.set_cell_data(row_idx, 1, values_dict["price"])
        if "total" in values_dict: self.sheet.set_cell_data(row_idx, 3, values_dict["total"])
        if "change" in values_dict: self.sheet.set_cell_data(row_idx,4, values_dict["change"])
    
    def SortData(self, sort_metric: str) -> None:
        '''Sorts the sheet based on the selected metric using raw_data keys'''
        if not self.raw_data: return

        # mapping
        metric_lookup = sort_metric.lower()
        mapping = {
            "stock price": "price",
            "amount": "quantity",
            "total value": "total",
            "daily change": "pct"
        }
        key = mapping.get(metric_lookup)
        
        if key is None:
            # This will now print the lowercase version if it fails
            print(f"Sort Error: '{metric_lookup}' not found in mapping.")
            return

        # 4. Sort (using .get() with a default of 0 is safer against missing keys)
        self.raw_data.sort(key=lambda x: x.get(key, 0), reverse=True)
        
        # 5. Update display
        self.SyncSheetWithRaw()
    
    def SyncSheetWithRaw(self) -> None:
        '''Converts raw_data back into formatted strings for the sheet'''
        formatted_table = []
        for row in self.raw_data:
            formatted_table.append([
                row['ticker'],
                row['price_str'],
                row['quantity'],
                row['total_str'],
                row['change_str']
            ])
        
        self.sheet.set_sheet_data(formatted_table)
        
        # Re-apply colors based on the raw pct
        for idx, row in enumerate(self.raw_data):
            colour = "#0F9D58" if row['pct'] >= 0 else "#DB4437"
            self.sheet.highlight_cells(row=idx, column=4, bg=colour, fg="white")
        self.sheet.set_column_widths([80, 80, 80, 100, 130])
        self.sheet.redraw()

    # Aesthetics
    def ModifyUsage(self) -> None:
        '''Changes how the chart works'''
        self.sheet.readonly_columns(columns = [1, 3, 4])
        self.sheet.enable_bindings((
            "single_select", 
            "row_select", 
            "edit_cell", 
            "column_width_resize",
            "row_deletion", 
            "arrowkeys", 
            "rc_delete_row",
            "copy",
            "cut",
            "paste",
            "undo"
        ))

        # Styling
        self.sheet.set_options(
            table_bg = THEME_DARK,    
            frame_bg = THEME_DARK, 
            header_bg = THEME_DARK,
            index_bg = THEME_DARK,
            index_fg = "#ADD8E6",
            top_left_bg = THEME_DARK,
            empty_horizontal=0, 
            empty_vertical=0
        )
        self.sheet.highlight_columns(
            columns = [1, 3, 4], 
            bg = "#57B9DA", 
            fg = "black"
        )
        self.sheet.highlight_columns(
            columns = [0, 2], 
            bg = "#D3D3D3", 
            fg = "black"
        )
        self.after(5, lambda: self.sheet.set_column_widths([80, 80, 80, 100, 130]))

class ControlFrame(ctk.CTkFrame):
    def __init__(self, parent, add_command: function, update_command: function, reset_command: function, toggle_command: function, sort_command: function, **kwargs):
        super().__init__(parent, fg_color = "transparent", **kwargs)

        # Setup
        self.currency_var = ctk.StringVar(value = "USD")
        self.sort_var = ctk.StringVar(value = "Sort by...")
        self.grid_columnconfigure((0, 1, 2, 3, 4), weight = 1)
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

        self.menu_sort = ctk.CTkOptionMenu(
            self,
            values = ["Total value", "Amount", "Daily change", "Stock price"],
            variable = self.sort_var,
            command = sort_command,
            fg_color = BTN_BLUE,
            button_hover_color = BTN_HOVER
        )
        self.menu_sort.grid(row = 0, column = 2, padx = 5)

        self.switch_currency = ctk.CTkSwitch(
            self, 
            text = "NZD/USD",
            variable = self.currency_var, 
            command = toggle_command,
            progress_color = BTN_BLUE, 
            onvalue = "NZD", offvalue = "USD",
            text_color = "white"
        )
        self.switch_currency.grid(row = 0, column = 3, padx = 5)

        self.button_reset = ctk.CTkButton(
            self, text = "Reset", fg_color = BTN_RED, hover_color = "#8a2424", width = 80,
            command = reset_command
        )
        self.button_reset.grid(row = 0, column = 4, padx = 5)
    
#endregion

if __name__ == "__main__":
    app = App()
    app.mainloop()