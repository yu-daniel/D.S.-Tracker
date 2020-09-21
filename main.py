import requests
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import datetime
import webbrowser
import pandas as pd
import numpy as np
import mplfinance as mpf
import matplotlib.dates as mpldates
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class Application(tk.Frame):
    """
    The Application class acts as the parent/controller for the different interfaces of the app.
    """
    def __init__(self, root):
        super().__init__(root)
        self.root = root
        self.grid()

        self.main = Main(self)
        self.toolbar = Toolbar(self)
        root.config(menu=self.toolbar)

class Toolbar(tk.Menu):
    """
    The Toolbar class creates a menubar along with drop down menus.
    """
    def __init__(self, root):
        super().__init__(root)
        self.root = root

        self.radio_status = 1

        # initiates the main menu, which would contain submenus (file, view, and help)
        self.menubar = tk.Menu(self)

        # initiate the File submenu
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.add_cascade(label="File", menu=self.filemenu)
        self.filemenu.add_command(label="Import", command=self.import_csv)
        self.filemenu.add_command(label="Export", command=self.export_data)
        self.filemenu.add_separator()
        self.filemenu.add_command(label="Exit", command=root.quit)

        # initiate the View submenu
        self.viewmenu = tk.Menu(self.menubar, tearoff=0)
        self.add_cascade(label="View", menu=self.viewmenu)

        # the variable that controls the radio buttons is set to 1, which set Line graph as the default
        self.radio_var = tk.IntVar()
        self.radio_var.set(1)
        self.viewmenu.add_radiobutton(label='Line', variable=self.radio_var, value=1, command=self.get_radio)
        self.viewmenu.add_radiobutton(label='Area', variable=self.radio_var, value=2, command=self.get_radio)
        self.viewmenu.add_radiobutton(label='Candlesticks', variable=self.radio_var, value=3, command=self.get_radio)
        self.viewmenu.add_separator()
        self.viewmenu.add_command(label="Select All", command=self.select_all)

        # initiate the Help submenu
        self.helpmenu = tk.Menu(self.menubar, tearoff=0)
        self.add_cascade(label="Help", menu=self.helpmenu)
        self.helpmenu.add_command(label="Homepage", command=lambda:
            webbrowser.open("https://github.com/kuckikirukia?tab=repositories"))
        self.helpmenu.add_separator()
        self.helpmenu.add_command(label="About", command=self.create_window)

    def import_csv(self):
        """
        Allows the user to import a list of tickers from a .csv or .txt file using File->Import.
        The list should be have no headers and separated only by commas (i.e. aapl,msft,amzn).
        """
        path = tk.filedialog.askopenfile(initialdir="/", title="Select File",
                filetypes=(("Comma-separated values (.csv)", "*.csv"), ("Text Document (.txt)", "*.txt"),
                           ("All Files", "*.*")))

        items = []
        if path is not None:
            for ticker in path:
                items.append(ticker)
        else:
            return

        tickers = items[0].split(',')
        for ticker in tickers:
            self.root.main.get_quote(ticker)

    def export_data(self):
        """
        Allows the user to use File->Export to save the current stock data in the main table (i.e. Treeview).
        The supported file types are .csv or .txt.
        """
        stocks = {}
        headings = ['Security', 'Price', 'Change', 'Change %', '52 Week', 'Market Cap']

        for data in range(6):
            for items in self.root.main.treeview.get_children():
                values = self.root.main.treeview.item(items, 'values')
                if headings[data] not in stocks:
                    stocks[headings[data]] = []
                stocks.get(headings[data]).append(values[data])

        df = pd.DataFrame(stocks, columns=headings)
        path = tk.filedialog.asksaveasfilename(title='Save File As...',
                    filetypes=(("CComma-separated values (.csv)", "*.csv"), ("Text Document(.txt)", "*.txt")))

        if not path:
            return
        else:
            df.to_excel(path, index=False, header=True)

    def get_radio(self):
        """
        Controls the type of graph that should be displayed when the user toggles
        between the different radio buttons (View->Line/Area/or Candlestick).
        """
        current_status = self.radio_var.get()
        if current_status != self.radio_status:
            self.radio_var.set(current_status)
            self.radio_status = current_status
        self.root.main.generate_graph(self.root.main.ticker)
        self.root.main.remove_old_graphs()

    def select_all(self):
        """
        Highlights all of the items in the main table (Treeview) when the user uses View->Select All.
        This enables the user to use the Remove button to delete multiple line of stocks, rather than
        one at a time.
        """
        selected_stocks = self.root.main.treeview.get_children()
        self.root.main.treeview.selection_set(selected_stocks)

    def create_window(self):
        """
        Generates a new toplevel (About) window displaying version changes and other info
        from Help->About.
        """
        pos_x = self.root.winfo_x()
        pos_y = self.root.winfo_y()

        about_window = tk.Toplevel(self)
        about_window.geometry('380x345' + f"+{pos_x + 250}+{pos_y + 100}")
        about_window.iconbitmap('icon.ico')
        about_window.resizable(False, False)

        # creates an 'Ok' buttons that allow the user to closes the About window
        ok_btn = HoverButton(about_window, text="Ok", height=1, width=6, command=about_window.destroy)
        ok_btn.grid(row=3, column=0, sticky=tk.E, padx=10, pady=5)

        about_label = tk.Label(about_window, text="Version Changes:", )
        about_label.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)

        about_frame = tk.Frame(about_window)
        about_frame.grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)

        text_box = tk.Text(about_frame, height=17, width=46, font=("Calibri", 10))
        text_box.grid(row=2, column=0, sticky=tk.W, padx=5)
        changes = open("credit.txt").read()
        text_box.insert(tk.END, changes)

        # adds a scrollbar for easier navigation for quicker viewing of version changes
        scrollbar = tk.Scrollbar(about_frame, command=text_box.yview)
        text_box.config(yscrollcommand=scrollbar.set, state=tk.DISABLED)
        scrollbar.grid(row=2, column=1, sticky='ns')
        about_window.transient(self.root)

class Main(tk.Frame):
    """
    The Main class creates the main interface for the app, which includes
    a field for the user to enter new tickers, buttons such as Add/Remove,
    a table that contains the stocks and their details.

    Attributes:
        holdings - tracks the current stocks in the main table and prevents adding duplicate tickers
        ticker - tracks the current highlighted stock from the table
        df = the current stock's data stored in dataframe format
        price = the current stock's price history in numpy column
        date = the current graph's list of dates in numpy float format
        date_list = the current graph's list of dates in string format
        graph = the current graph that is being displayed
    """
    def __init__(self, root):
        super().__init__(root)
        self.root = root
        self.grid()

        self.holdings = {}
        self.ticker = None
        self.df = None
        self.price = None
        self.date = None
        self.date_list = None
        self.graph = None

        # initialize the main interface's Labels & Buttons
        self.addLabel = tk.Label(self, text="Ticker: ")
        self.addEntry = tk.Entry(self, width=15, fg="#3463ad")

        self.addButton = HoverButton(self, text='Add', command=self.add_ticker)
        self.removeButton = HoverButton(self, text='Remove', command=self.remove_pos)
        self.updateButton = HoverButton(self, text='Update', command=self.update_data)
        self.showGraph = HoverButton(self, text='Show Graph', command=self.show_data)

        # placement for the labels and buttons
        self.addLabel.grid(row=0, column=0, sticky=tk.W, padx=10, pady=7)
        self.addEntry.grid(row=0, column=0, sticky=tk.W, padx=50, pady=7)
        self.addButton.grid(row=1, column=0, sticky=tk.W, padx=10, pady=7)
        self.removeButton.grid(row=1, column=0, sticky=tk.W, padx=50, pady=7)
        self.updateButton.grid(row=1, column=0, sticky=tk.W, padx=110, pady=7)
        self.showGraph.grid(row=1, column=0, sticky=tk.W, padx=165, pady=7)

        # initiate the main table (i.e. a Treeview)
        columns = ("Ticker", "Price", "Change", "Change %", "52-week", "Market Cap")
        custom_width = (60, 60, 75, 75, 60, 100)
        self.treeview = ttk.Treeview(self, columns=columns, show="headings", height=15)
        self.treeview.grid(row=2, columnspan=10, padx=10, pady=16)

        for index, col in enumerate(columns):
            self.treeview.heading(col, text=col, command=lambda _col=col:
            self.treeview_sort_column(self.treeview, _col, False))
            self.treeview.column(col, width=custom_width[index], anchor='center')

        # add color/styling to the Treeview
        self.style = ttk.Style(root)
        self.style.theme_use("clam")
        self.style.map("Treeview", foreground=self.fixed_map("foreground"), background=self.fixed_map("background"))

        self.style.configure("Treeview", foreground="white")
        self.treeview.tag_configure('green', foreground='green')
        self.treeview.tag_configure('red', foreground='red')

    def add_ticker(self):
        """
        Retrieves the user's ticker entry and calls get_quote() to add the stock to the main table.
        """
        ticker = self.addEntry.get().upper()
        self.get_quote(ticker)

    def remove_pos(self):
        """
        Removes the line(s) of stocks in the main table that is highlighted by the user.
        """
        selected_items = self.treeview.selection()
        for items in selected_items:
            values = self.treeview.item(items, 'values')
            if values[0] in self.holdings:
                del self.holdings[values[0]]
            self.treeview.delete(items)
        return None

    def update_data(self):

        treeview = None
        tickers = [i for i in self.holdings.keys()]

        # find treeview from the frame's list of children and delete all its entries
        for widget in self.winfo_children():
            if widget.winfo_class() == 'Treeview':
                treeview = widget
            for ticker in treeview.get_children():
                self.treeview.delete(ticker)

        self.holdings = {}
        for ticker in tickers:
            self.get_quote(ticker)

    def show_data(self):
        """
        Grabs the line of stock highlighted in the main table and calls graph_data() to
        generate a graph that reflects the 52-week performance.
        """
        selected_items = self.treeview.selection()
        if len(selected_items) > 0:
            ticker = self.treeview.item(selected_items, 'values')[0]
            self.graph_data(ticker)
            self.ticker = ticker
        else:
            return None

    def graph_data(self, ticker):
        """
        From the user's ticker input is used to call the Alpha Vantage API to retrieve
        the stock's open, high, low, volume, and other information.
        """
        key = 'GLC0GTVKR51SY1V'

        url = 'https://www.alphavantage.co/query?function=TIME_SERIES_MONTHLY&symbol=IBM&apikey=demo'
        response = requests.get(url)
        string = response.json()

        ticker = string['Meta Data']['2. Symbol']
        dic = string['Monthly Time Series']
        keys = string['Monthly Time Series'].keys()
        key_list = list(keys)

        key_data = []
        date_list = []
        open_list = []
        high_list = []
        low_list = []
        close_list = []
        volume_list = []

        for x in range(len(key_list)-1, 0, -1):

            date = key_list[x]
            Open = dic[key_list[x]]['1. open']
            High = dic[key_list[x]]['2. high']
            Low = dic[key_list[x]]['3. low']
            Close = dic[key_list[x]]['4. close']
            Volume = dic[key_list[x]]['5. volume']

            entry = date + "," + Open
            key_data.append(entry)
            date_list.append(date)
            open_list.append(float(Open))
            high_list.append(float(High))
            low_list.append(float(Low))
            close_list.append(float(Close))
            volume_list.append(float(Volume))

        date, price = np.loadtxt(reversed(key_data), delimiter=',', unpack=True, converters={0: self.bytes_to_dates})

        # datelist_strs = []
        #
        # for date in date_list:
        #     new_date = datetime.datetime.strptime(date, "%Y-%m-%d")
        #     datelist_strs.append(new_date)

        date_objects = [datetime.datetime.strptime(date, '%Y-%m-%d') for date in date_list]

        dictionary = {'Date': date_objects, 'Open': open_list, 'High': high_list, 'Low': low_list, 'Close': close_list,
                      'Volume': volume_list}

        df = pd.DataFrame.from_dict(dictionary)
        df.set_index('Date', inplace=True)

        self.df = df
        self.date = date
        self.price = price
        self.date_list = date_list
        self.generate_graph(ticker)

    def generate_graph(self, ticker):
        """
        Initiates an extended window of the main interface to show the graph
        of the stock that the user selected.
        """
        mc = mpf.make_marketcolors(up='#00c223', down='#e83c3c', inherit=True, )
        custom_style = mpf.make_mpf_style(marketcolors=mc, gridstyle='dotted', rc={'font.size': 500})

        file_format = tk.IntVar()
        frame = tk.Frame(self)
        frame.grid(row=2, column=10, sticky=tk.W, padx=10, pady=5)

        saveGraph = HoverButton(frame, text='Save', command=self.save_image)
        saveGraph.grid(row=0, column=11, sticky=tk.W, pady=5, padx=0)

        closeGraph = HoverButton(frame, text='Close', command=frame.destroy)
        closeGraph.grid(row=0, column=11, sticky=tk.W, pady=5, padx=40)

        fig = plt.Figure(figsize=(4.5, 3.5), dpi=80)
        fig.suptitle(ticker, fontsize=12)

        # line graph
        if self.root.toolbar.radio_var.get() == 1:
            image = fig.add_subplot(111, xlabel='Year', ylabel='Price')
            image.plot_date(self.date, self.price, '-',)

        # area graph
        elif self.root.toolbar.radio_var.get() == 2:
            image = fig.add_subplot(111, xlabel='Year', ylabel='Price')
            image.plot_date(self.date, self.price,  '-',)

            image.fill_between(self.date, self.price, color='#539ecd')

        # candle graph
        elif self.root.toolbar.radio_var.get() == 3:
            image = fig.add_subplot(111)
            mpf.plot(self.df, ax=image, type='candle', style=custom_style, xrotation=15)

        chart = FigureCanvasTkAgg(fig, frame)
        chart.get_tk_widget().grid(row=1, column=11, sticky=tk.N+tk.W, pady=7)
        self.graph = fig

    def treeview_sort_column(self, treeview, column, reverse):
        """
        See Reference
        Allows the user to click on the columns in the main table, which would sort them
        in ascending or descending order.
        """
        data = [(treeview.set(ticker, column), ticker) for ticker in treeview.get_children('')]
        data.sort(reverse=reverse)

        # sort the stock(s)
        for index, (val, k) in enumerate(data):
            treeview.move(k, '', index)

        # reverse sort next time
        treeview.heading(column, command=lambda: self.treeview_sort_column(treeview, column, not reverse))

    def fixed_map(self, option):
        """
        Function fixes a styling bug in Tkinter - see References.
        Returns the style map for 'option' with any styles starting with
        ("!disabled", "!selected", ...) filtered out
        """
        return [elm for elm in self.style.map("Treeview", query_opt=option) if elm[:2] != ("!disabled", "!selected")]

    def get_quote(self, ticker):
        """
        Adds the current stock and its data to the main table (treeview).
        """
        key = 'GLC0GTVKR51SY1V'
        quote_url = 'https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=' + ticker.upper() + '&apikey=' + key
        key_metrics_url = 'https://www.alphavantage.co/query?function=OVERVIEW&symbol=' + ticker.upper() + '&apikey=' + key

        quote_response = requests.get(quote_url)
        string = quote_response.json()

        key_metrics_response= requests.get(key_metrics_url)
        metrics_str = key_metrics_response.json()
        color_tag = None

        if quote_response and 'Global Quote' in string:

            current_price = round(float(string['Global Quote']['05. price']), 2)
            change = round(float(string['Global Quote']['09. change']), 2)
            change_pct = string['Global Quote']['10. change percent'][:5] + "%"
            previous_price = round(float(string['Global Quote']['08. previous close']), 2)

            yearly_high = metrics_str['52WeekHigh']
            mark_cap = round(int(metrics_str['MarketCapitalization'])/10E8, 2)
            mark_cap_str = str(mark_cap) + "B"

            if ticker not in self.holdings:
                self.holdings[ticker] = current_price
                tuples = [ticker, current_price, change, change_pct, yearly_high, mark_cap_str]

                if current_price > previous_price:
                    color_tag = 'green'
                else:
                    color_tag = 'red'
                self.treeview.insert(parent='', index='end', values=tuples, tags=(color_tag,))
            return current_price
        else:
            return None

    def save_image(self):
        """
        Opens a window that enables the user to save the current graph as a .png or .pdf file.
        """
        filename = filedialog.asksaveasfilename(title='Save Image As...',
        filetypes=(("Portable Network Graphics (.png)", "*.png"), ("Portable Document Format(.pdf)", "*.pdf")))
        self.graph.savefig(filename, dpi=self.graph.dpi)

    def bytes_to_dates(self, date_str):
        """
        Dates obtained from the API are in string form, which is converted to Date objects
        using Date's datestr2num() method.
        """
        return mpldates.datestr2num(date_str.decode('utf-8'))

    def remove_old_graphs(self):
        """
        A new graph is generated when the user toggles between the Line, Area, and Candlesticks
        radio buttons. This function deletes the previous graph after the new one is
        generated.
        """
        widgets = self.winfo_children()
        graph_frames = []

        for widget in widgets:
            if type(widget) == tk.Frame:
                graph_frames.append(widget)

        for frame in range(len(graph_frames) - 1):
            graph_frames[frame].destroy()


class HoverButton(tk.Button):
    """
    The HoverButton class allow buttons to change color when hovered over/enters, and
    reverts to normal when exists.
    """
    def __init__(self, master, **kw):
        tk.Button.__init__(self, master=master, **kw)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.mouse_in)
        self.bind("<Leave>", self.mouse_out)

    def mouse_in(self, event):
        """
        Creates a subtle color change to the button that the users' mouse hovers over.
        """
        self['background'] = '#E5F3FF'

    def mouse_out(self, event):
        """
        When the user's mouse leaves a button, the color of the button reverts back to default.
        """
        self['background'] = self.defaultBackground


if __name__ == '__main__':
    root = tk.Tk()
    root.title("My Stock Tracker")
    root.iconbitmap('icon.ico')
    root.resizable(False, False)
    app = Application(root)
    root.mainloop()