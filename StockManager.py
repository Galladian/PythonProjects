#region imports
import customtkinter as ctk
import yfinance as yf
try:
    from ctypes import windll, byref, sizeof, c_int
except:
    pass
#endregion

class App(ctk.CTk):
    def __init__(self):
        # setup
        super().__init__(fg_color="#19233c")
        self.title("Stock Manager")
        self.geometry("500x450")
        self.minsize(500,200)
        self.ChangeTitleBar()

        # variables
        self.portfolio_total_string = ctk.StringVar(value = "Total: $0.00")

        # widgets
        self.CreateFrames()
    
    def CreateFrames(self):
        self.summary_frame = SummaryFrame(self, self.portfolio_total_string)
        self.summary_frame.place(relx = 0, rely = 0.85, relwidth = 1.0, relheight = 0.15)

        self.main_frame = MainFrame(self)
        self.main_frame.place(relx = 0, rely = 0, relwidth = 1.0, relheight = 0.85)

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
    def __init__(self, parent, portfolio_string: str, **kwargs):
        super().__init__(parent, fg_color="#121a2d", **kwargs)

        self.total_label = ctk.CTkLabel(
            self,
            textvariable = portfolio_string,
            font=("Helvetica", 14, "bold"),
            text_color="white"
        )
        self.total_label.pack(expand=True)
    
    def UpdateTotal(self, amount: float):
        '''Method to modify total'''
        self.total_label.configure(text = f"Total: ${amount:,.2f}")

#region MAINFRAME
class MainFrame(ctk.CTkScrollableFrame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color = "transparent", corner_radius = 0, **kwargs)

        self.grid_columnconfigure((1, 3), weight = 1) #ticker and quantity have expansive nature
        self.grid_columnconfigure((0, 2, 4, 5), weight = 0)
        self.rows_data = []
    
    def AddRow(self) -> None:
        '''Allows the addition of a new row'''
        index = len(self.rows_data)
        row_dictionary = {}

        row_dictionary["num"] = ctk.CTkLabel(self, text = f"{index + 1}", width = 30)
        row_dictionary["num"].grid(row = index, column = 0, padx = 5, pady = 5)

        row_dictionary["ticker"] = ctk.CTkEntry(self, placeholder_text = "Ticker", width = 120)
        row_dictionary["ticker"].grid(row=index, column = 1, padx = 5, pady = 5, sticky = "ew")

        row_dictionary["price"] = ctk.CTkLabel(self, text = "$0.00", width = 80)
        row_dictionary["price"].grid(row = index, column = 2, padx = 5, pady = 5)

        row_dictionary["amount"] = ctk.CTkEntry(self, placeholder_text="0.0", width=80)
        row_dictionary["amount"].grid(row = index, column = 3, padx = 5, pady = 5, sticky = "ew")

        row_dictionary["total"] = ctk.CTkLabel(self, text = "$0.00", width = 80)
        row_dictionary["total"].grid(row = index, column = 4, padx = 5, pady = 5)

        row_dictionary["delete"] = ctk.CTkButton(self, text="-", width=30, fg_color="#1f538d",
                                command = lambda: self.RemoveRow(row_dictionary))
        row_dictionary["delete"].grid(row = index, column = 5, padx = 5, pady = 5)

        # Store references so we can access their data later
        self.rows_data.append(row_dictionary)
    
    def RemoveRow(self, row_dict: dict) -> None:
        '''Allows the removal of an existing row'''
        for widget in self.row_dict.values():
            widget.destroy()
        self.rows_data.remove(row_dict)
        self.RefreshGrid()

    def RefreshGrid(self):
        '''Re-indexes and moves data after deletion'''
        for index, row in enumerate(self.rows_data): 
            row["num"].grid(row = index, column = 0)
            row["ticker"].grid(row = index, column = 1)
            row["price"].grid(row = index, column = 2)
            row["amount"].grid(row = index, column = 3)
            row["total"].grid(row = index, column = 4)
            row["delete"].grid(row = index, column = 5)
            
            # Update the visual row number
            row["num"].configure(text=str(index + 1))
#endregion

if __name__ == "__main__":
    app = App()
    app.mainloop()