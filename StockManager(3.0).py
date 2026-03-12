#region IMPORTS + SETTINGS

# standard imports
import json
import threading
from ctypes import byref, c_int, sizeof, windll

# Third party libraries
import customtkinter as ctk
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import yfinance as yf
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
from tksheet import Sheet

# general theme
THEME_BG = "#19233c"        
THEME_DARK = "#121a2d"      

# buttons
BTN_REG = "#1f538d"        
BTN_RESET = "#C53434"        
BTN_HOVER = "#14375e"
BTN_RESET_HOVER = "#8a2424"

# for graph
ANNOT_BG = "#2B2B2B"
LINE_PLOT = "#90D5FF"
#endregion

class App(ctk.CTk):
    def __init__(self):
        # setup
        super().__init__(fg_color = THEME_BG)
        self.title("Stock Manager")
        self.geometry("525x500")
        self.minsize(525,350)
        self.ChangeTitleBar()

        self.exchange_rate = 1.0
        self.last_prices = (None, None)

        # widgets
        self.CreateFrames()
        self.LoadData()

        # detection
        self.protocol("WM_DELETE_WINDOW", self.OnClose)
    
    def CreateFrames(self) -> None:
        '''Adds frame widgets onto window'''
        self.control_frame = ControlFrame(self, self.AddRowCallback, self.UpdateCallback, self.ResetCallback, self.ToggleCallback, self.SortCallback)
        self.control_frame.place(relx = 0, rely = 0, relwidth = 1.0, relheight = 0.15)

        self.graph_frame = GraphFrame(self)
        self.graph_frame.place(relx = 0.6, rely = 0.15, relwidth = 0.4, relheight = 0.85)

        self.main_frame = MainFrame(self)
        self.main_frame.place(relx = 0, rely = 0.15, relwidth = 0.6, relheight = 0.7)

        self.summary_frame = SummaryFrame(self)
        self.summary_frame.place(relx = 0, rely = 0.85, relwidth = 0.6, relheight = 0.15)

    # Callback functions
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
        self.main_frame.AddRow()
        self.main_frame.DynamicColumnResize(None)
        self.main_frame.sheet.redraw()

    def ToggleCallback(self) -> None:
        '''Called when switch is flipped or for fresh data'''
        if self.last_prices is None: return

        is_nzd = self.control_frame.currency_var.get() == "NZD"
        multiplier = self.exchange_rate if is_nzd else 1.0

        self.ApplyPricesToUI(self.last_prices[0], self.last_prices[1], multiplier)

    def UpdateCallback(self) -> None:
        '''Single entry point to trigger the background update chain'''
        self.control_frame.button_update.configure(state = "disabled", text = "Fetching..")
        table_data = self.main_frame.GetTableData()
        tickers = [row[0].strip().upper() for row in table_data if row[0].strip()]
        
        if not tickers: return

        # Start ONE thread that handles the entire sequence
        threading.Thread(target = self.SequentialUpdateTask, args = (tickers, table_data), daemon = True).start()

    # Update data and apply
    def FetchPrices(self, tickers: list) -> None:
        '''Background task to fetch data and update UI'''
        try:
            # retrieves most recent data
            data = yf.download(tickers + ["NZD=X"], period = "7d", interval = "1d", progress = False, prepost = True)
            close_data = data['Close'].ffill().bfill()

            # 2. Assign the rows
            index = -1
            while close_data.iloc[index][tickers[0]] == close_data.iloc[index - 1][tickers[0]]:
                index -= 1

            self.last_prices = (close_data.iloc[index - 1], close_data.iloc[-1])
            self.exchange_rate = float(self.last_prices[1]["NZD=X"])

            self.after(0, lambda: self.ToggleCallback())
        except Exception as e:
            print(f"Error fetching data: {e}")

    def FetchHistoricalData(self, portfolio_map: dict) -> None:
        '''Fetches 1yr history and calculates performance'''
        try:
            # grab and filter data
            tickers = list(portfolio_map.keys())
            data = yf.download(tickers, period="1y", interval="1wk", progress=False)
            
            if 'Close' in data:
                close_data = data['Close']
            else:
                close_data = data
                
            close_data = close_data.ffill().bfill()
            total_history = None
            
            for ticker, qty in portfolio_map.items():
                if ticker in close_data.columns:
                    # fillna(0) ensures that if a stock didn't exist yet, it just counts as $0 
                    series = close_data[ticker].fillna(0) * qty
                    if total_history is None:
                        total_history = series
                    else:
                        total_history = total_history.add(series, fill_value=0)

            if total_history is not None:
                # updates graph 
                dates = total_history.index
                values = total_history.values
                self.after(0, lambda: self.graph_frame.UpdateChart(dates, values))             
        except Exception as e:
            print(f"Graph Error Logic: {e}") 
        
        # reactivate update button
        self.control_frame.button_update.configure(state = "normal", text = "Update")

    def SequentialUpdateTask(self, tickers: dict, table_data: dict) -> None:
        '''Guarantees that Table finishes before Graph starts to avoid yfinance collisions'''
        self.FetchPrices(tickers) 
        portfolio_map = {row[0].strip().upper(): float(row[2] or 0) for row in table_data if row[0].strip()}
        if portfolio_map:
            self.FetchHistoricalData(portfolio_map)

    def ApplyPricesToUI(self, prev_prices: dict, current_prices: dict, multiplier: float = 1.0) -> None:
        '''Updates the UI safely, skipping corrupted data to prevent NaN/zeroing errors.'''
        total_value = 0.0
        total_change = 0.0
        currency_sym = "NZ$" if multiplier != 1.0 else "$"
        
        table_data = self.main_frame.GetTableData()
        new_raw_data = [] 

        for row in table_data:
            ticker = row[0].strip().upper() 
            
            if ticker in current_prices and ticker in prev_prices:
                # ensure data is present 
                p_curr = current_prices[ticker]
                p_prev = prev_prices[ticker]
                
                if pd.isna(p_curr) or pd.isna(p_prev) or p_prev == 0:
                    continue 

                price = p_curr * multiplier
                prev_close = p_prev * multiplier
                
                try:
                    quantity = float(row[2] or 0)
                except ValueError:
                    quantity = 0.0

                # sorta daa
                change_percent = ((price - prev_close) / prev_close) * 100
                quantity_change = (price - prev_close) * quantity
                row_total = price * quantity
                
                new_raw_data.append({
                    'ticker': ticker,
                    'price': price,
                    'quantity': quantity,
                    'total': row_total,
                    'pct': change_percent,
                    'qty_change': quantity_change,
                    'price_str': f"{currency_sym}{price:,.2f}",
                    'total_str': f"{currency_sym}{row_total:,.2f}",
                    'change_str': f"{'+' if quantity_change >= 0 else '-'}{currency_sym}{abs(quantity_change):,.2f} ({change_percent:+.2f}%)"
                })

                total_value += row_total
                total_change += quantity_change

        self.main_frame.raw_data = new_raw_data
        self.main_frame.SyncSheetWithRaw()
        self.summary_frame.UpdateSummary(total_value, total_change)

    # Data persistence functions
    def OnClose(self) -> None:
        '''Executes when application is closed'''
        self.SaveData()
        for after_id in self.tk.eval('after info').split():
            self.after_cancel(after_id)
        self.quit()
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
        super().__init__(parent, fg_color = THEME_DARK, corner_radius = 0, **kwargs)

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
        self.bind("<Configure>", self.RescaleText)
    
    def RescaleText(self, event):
        '''Calculates and updates font sizes based on frame width'''

        combined_metric = (event.width + (event.height*4)) / 2
        total_font_size = max(14, int(combined_metric / 25)) 
        change_font_size = max(10, int(combined_metric / 50))

        # Apply new sizes
        self.total_label.configure(font=("Helvetica", total_font_size, "bold"))
        self.change_label.configure(font=("Helvetica", change_font_size, "bold"))
    
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
    def __init__(self, parent, **kwargs):
        # setup
        super().__init__(parent, fg_color = THEME_DARK, corner_radius = 0, **kwargs)
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
        self.base_widths = [60, 60, 60, 80, 110]
        self.total_base = sum(self.base_widths)
        self.ModifyUsage()

        self.bind("<Configure>", self.DynamicColumnResize)

    # Functionality
    def AddRow(self) -> None:
        '''Insert a new row with default values'''
        self.sheet.insert_row(["", "$0.00", "", "$0.00", "0.00%"])
        self.sheet.redraw()

    def GetTableData(self) -> None:
        '''Returns all row data as a list of lists'''
        return self.sheet.get_sheet_data()

    def UpdateRow(self, row_idx: int, values_dict: dict) -> None:
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
            "percent change": "pct",
            "quantity change": "qty_change"
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
        
        self.DynamicColumnResize(None)
    
    def DynamicColumnResize(self, event = None) -> None:
        '''Adjusts column widths based on frame width while maintaining ratios'''
        current_width = (event.width if event else self.winfo_width()) - 60
    
        if current_width > 100:
            new_widths = []
            for w in self.base_widths:
                calculated_width = int((w / self.total_base) * current_width)
                new_widths.append(calculated_width)
            
            self.sheet.set_column_widths(new_widths)

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

class ControlFrame(ctk.CTkFrame):
    def __init__(self, parent, add_command: function, update_command: function, reset_command: function, toggle_command: function, sort_command: function, **kwargs):
        super().__init__(parent, fg_color = THEME_BG, corner_radius = 0, **kwargs)

        # Setup
        self.currency_var = ctk.StringVar(value = "USD")
        self.sort_var = ctk.StringVar(value = "Sort by...")

        # setting widgets
        self.button_add = ctk.CTkButton(
            self, 
            text = "+ Add Row", 
            fg_color = BTN_REG, 
            hover_color = BTN_HOVER, 
            command = add_command  
        )
        self.button_add.place(relx = 0.01, rely = 0.2, relwidth = 0.18, relheight = 0.6)

        self.button_update = ctk.CTkButton(
            self, 
            text = "Update", 
            fg_color = BTN_REG, 
            hover_color = BTN_HOVER, 
            command = update_command
        )
        self.button_update.place(relx = 0.21, rely = 0.2, relwidth = 0.18, relheight = 0.6)

        self.menu_sort = ctk.CTkOptionMenu(
            self,
            values = ["Total value", "Amount", "Percent change", "Quantity change", "Stock price"],
            variable = self.sort_var,
            command = sort_command,
            fg_color = BTN_REG,
            button_color = BTN_REG,
            button_hover_color = BTN_HOVER
        )
        self.menu_sort.place(relx = 0.41, rely = 0.2, relwidth = 0.18, relheight = 0.6)

        self.switch_currency = ctk.CTkSwitch(
            self, 
            text = "USD/NZD",
            variable = self.currency_var, 
            command = toggle_command,
            progress_color = BTN_REG, 
            onvalue = "NZD", offvalue = "USD",
            text_color = "white"
        )
        self.switch_currency.place(relx = 0.61, rely = 0.2, relwidth = 0.18, relheight = 0.6)

        self.button_reset = ctk.CTkButton(
            self, text = "Reset", fg_color = BTN_RESET, hover_color = BTN_RESET_HOVER, width = 80,
            command = reset_command
        )
        self.button_reset.place(relx = 0.81, rely = 0.2, relwidth = 0.18, relheight = 0.6)

        # to make the drop down menu more uniform
        self.menu_sort.bind("<Enter>", lambda e: self.menu_sort.configure(fg_color = BTN_HOVER, button_color = BTN_HOVER))
        self.menu_sort.bind("<Leave>", lambda e: self.menu_sort.configure(fg_color = BTN_REG, button_color = BTN_REG))

class GraphFrame(ctk.CTkFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color = THEME_DARK, corner_radius = 0, **kwargs)
        
        self.fig, self.ax = plt.subplots(figsize = (5, 4), dpi = 100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self)
        self.canvas.get_tk_widget().pack(fill = "both", expand = True, padx = 5, pady = 5)

        # Store data references for the hover logic
        self.line_data_x = []
        self.line_data_y = []

        self.SetStyle()
        
        # 1. Create the hover tooltip (initially invisible)
        self.annotation_box = self.ax.annotate(
            "", xy = (0,0), xytext = (10, 10),
            textcoords = "offset points",
            bbox = dict(boxstyle = "round", fc = "white", alpha = 0.8),
            arrowprops = dict(arrowstyle = "->", color = 'white')
        )
        self.annotation_box.set_visible(False)
        
        # 2. Bind the motion event
        self.canvas.mpl_connect("motion_notify_event", self.OnHover)
        self.bind("<Configure>", self.OnResize)

    def UpdateChart(self, dates: list, values: list) -> None:
        '''Clears existing plot and draws new data.'''
        if dates is None or values is None or len(dates) == 0: return

        self.line_data_x = dates
        self.line_data_y = values

        self.ax.clear()
        self.SetStyle()

        # Re-initialize annotation after clearing
        self.annotation_box = self.ax.annotate(
            "", xy = (0,0), xytext = (10, 10),
            textcoords = "offset points",
            bbox = dict(boxstyle = "round", fc = ANNOT_BG, ec = "white"),
            color = "white", fontsize = 8,
            arrowprops = dict(arrowstyle = "->", color = "white")
        )
        self.annotation_box.set_visible(False)
               
        # Personalised style for data
        self.ax.plot(dates, values, color = LINE_PLOT, linewidth = 2, zorder = 2)

        minimum_value = min(values) 
        self.ax.set_ylim(bottom = minimum_value * 0.99) # Add 1% breathing room  
        y_min = self.ax.get_ylim()[0]
        self.ax.fill_between(
            dates, 
            values, 
            y2 = y_min, 
            color = BTN_REG, 
            alpha = 0.1, 
            zorder = 1,
            clip_on = False
        )          
        self.ax.margins(x = 0)
        
        self.fig.tight_layout()
        self.canvas.draw() 

    def OnHover(self, event) -> None:
        '''Calculates nearest point and toggles visibility of the tooltip.'''
        is_visible = self.annotation_box.get_visible()
        
        if event.inaxes == self.ax and event.xdata is not None:
            try:
                # find point
                x_data_nums = mdates.date2num(self.line_data_x)
                index = np.argmin(np.abs(x_data_nums - event.xdata))
                self.annotation_box.xy = (self.line_data_x[index], self.line_data_y[index])
                
                # format
                date_string = self.line_data_x[index].strftime("%b %d, %Y")
                text = f"{date_string}\n${self.line_data_y[index]:,.2f}"
                
                self.annotation_box.set_text(text)
                self.annotation_box.set_visible(True)
                self.canvas.draw_idle()
            except (ValueError, TypeError) as e:
                print(f"Error on hover: {e}")
        else:
            if is_visible:
                self.annotation_box.set_visible(False)
                self.canvas.draw_idle()

    # Aesthetics and design
    def SetStyle(self) -> None:
        '''Sets up the visual theme that doesn't change with data updates.'''
        self.fig.patch.set_facecolor(THEME_DARK)
        self.ax.set_facecolor(THEME_DARK)
        
        # Spine & Tick Configuration
        self.ax.spines["top"].set_visible(False)
        self.ax.spines["right"].set_visible(False)
        for spine in self.ax.spines.values():
            spine.set_color("white")
            
        self.ax.tick_params(colors = "white", labelsize = 8)
        self.ax.xaxis.set_tick_params(rotation = 45)
        
        # Grid Configuration
        self.ax.yaxis.grid(True, linestyle = "--", alpha = 0.3, color = 'gray', zorder = 1)
        self.ax.set_title("Portfolio Performance (1Y)", color = "white", fontsize = 10, pad = 10)
        self.OnResize()

    def OnResize(self, event = None):
        '''Adjusts tick density and font size based on current width.'''
        current_width = (event.width if event else self.winfo_width())
        
        # Find the first threshold that fits
        thresholds = [
            (300, 4, 6, "%b"),
            (500, 7, 7, "%b %d"),
            (float('inf'), 9, 8, "%b %d %Y") # Default/Large
        ]

        for limit, bins, size, fmt in thresholds:
            if current_width < limit:
                nbins, font_size, date_format = bins, size, fmt
                break

        #nApply to axis without clearing the whole plot
        self.ax.xaxis.set_major_locator(MaxNLocator(nbins = nbins))
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter(date_format))
        self.ax.tick_params(axis = "x", labelsize = font_size)
        
        # 3. Refresh canvas
        self.fig.autofmt_xdate()
        self.canvas.draw_idle()
#endregion

if __name__ == "__main__":
    app = App()
    app.mainloop()