from Tkinter import *
import xlrd
import tkFileDialog
from xlwt import Workbook
import datetime
import requests
from bs4 import BeautifulSoup as bs
import os.path

class product():    #The class for the basic properties of the product.
    def __init__(self, width, length, height, thickness, lid, seprator):

        self.width = float(width)
        self.length = float(length)
        self.height = float(height)
        self.thickness = float(thickness)
        self.lid = bool(lid)
        self.seprator = bool(seprator)

class input(product):   #The class for the calculation elemnts of the product.

    def __init__(self, try_usd, steel_price, surface_area, volume, weight, cost):
        product.__init__(self, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0)

        self.steel_price = float(steel_price)
        self.try_usd = float(try_usd)
        self.surface_area = float(surface_area)
        self.volume = float(volume)
        self.weight = float(weight)
        self.cost = float(cost)

class SteelBox(Frame, input):   #The class for GUI.
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.grid()
        self.parent = parent
        self.initUI()

    def initUI(self):
        #Title
        self.title_label = Label(self, text="SteelBox Inc. Calculator", font="Arial 20 bold")
        self.title_label.grid(row=0, column=0, columnspan=12, sticky="NEW", padx=200, pady=25)

        #Import Button
        self.import_button = Button(self, text="Import", bd=10, pady=10, padx=25, font="Arial 10 bold", command=self.open_file)
        self.import_button.grid(row=1, column=4, columnspan=2, pady=20, padx=25)

        #Width Entry
        self.width_default_text = StringVar()
        self.width_entry_label = Label(self, text="Width ", font="Arial 15")
        self.width_entry_label.grid(row=2, column=5, padx=25)
        self.width_entry = Entry(self, bd=10, textvariable=self.width_default_text)
        self.width_entry.grid(row=2, column=6, pady=20, padx=25)

        #Length Entry
        self.length_default_text = StringVar()
        self.length_entry_label = Label(self, text="Length", font="Arial 15")
        self.length_entry_label.grid(row=3, column=5, padx=25)
        self.length_entry = Entry(self, bd=10, textvariable=self.length_default_text)
        self.length_entry.grid(row=3, column=6, pady=20, padx=25)

        #Height Entry
        self.height_default_text = StringVar()
        self.height_entry_label = Label(self, text="Height", font="Arial 15")
        self.height_entry_label.grid(row=4, column=5, padx=25)
        self.height_entry = Entry(self, bd=10, textvariable=self.height_default_text)
        self.height_entry.grid(row=4, column=6, pady=20, padx=25)

        #Thickness Entry
        self.thickness_default_text = StringVar()
        self.thickness_entry_label = Label(self, text="Thickness", font="Arial 15")
        self.thickness_entry_label.grid(row=5, column=5, padx=25)
        self.thickness_entry = Entry(self, bd=10, textvariable=self.thickness_default_text)
        self.thickness_entry.grid(row=5, column=6, pady=20, padx=25)

        #placeholder
        self.plc = Label(self, text="------------------------------------------"
                                    "------------------------------------------"
                                    "------------------------------------------"
                                    "------------------------------------------")
        self.plc.grid(row=6, column=0, columnspan=12)

        self.plc2 = Label(self, text="|\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|"
                                     "\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|\n|")
        self.plc2.grid(row=1, column=7, rowspan=6)

        self.plc3 = Label(self, text="--------------------------------"
                                     "--------------------------------"
                                     "--------------------------------")
        self.plc3.grid(row=3, column=8, columnspan=4)

        #Calculate Button
        self.calculate_button = Button(self, text="Calculate", bd=10, command=self.calc, pady=10, padx=25, font="Arial 10 bold")
        self.calculate_button.grid(row=8, column=6, columnspan=1, pady=20)

        #Total Weight Entry
        self.total_weight_default = StringVar()
        self.total_weight_entry_label = Label(self, text="Total Weight", font="Arial 15")
        self.total_weight_entry_label.grid(row=7, column=8, padx=25)
        self.total_weight_entry = Entry(self, bd=10, textvariable=self.total_weight_default)
        self.total_weight_entry.grid(row=8, column=8, pady=20, padx=25)

        #Total Price Entry
        self.total_price_default = StringVar()
        self.total_price_entry_label = Label(self, text="Total Price", font="Arial 15")
        self.total_price_entry_label.grid(row=7, column=9, padx=25)
        self.total_price_entry = Entry(self, bd=10, textvariable=self.total_price_default)
        self.total_price_entry.grid(row=8, column=9, pady=20, padx=25)

        #Export Button
        self.export_button = Button(self, text="Export", bd=10, pady=10, padx=25, font="Arial 10 bold", command = self.export)
        self.export_button.grid(row=8, column=10, columnspan=3, pady=20, padx=25)

        #Current Steel Price Entry
        self.current_steel_default = StringVar()
        self.steel_price_entry_label = Label(self, text="Current\nSteel Price", font="Arial 15")
        self.steel_price_entry_label.grid(row=4, column=8, padx=25)
        self.steel_price_entry = Entry(self, bd=10, textvariable=self.current_steel_default)
        self.steel_price_entry.grid(row=4, column=9, pady=20, padx=25)

        #Exchange Rate Entry
        self.exchange_rate_default = StringVar()
        self.exchange_rate_entry_label = Label(self, text="TRY/USD\nExchange Rate", font="Arial 15")
        self.exchange_rate_entry_label.grid(row=5, column=8, padx=25)
        self.exchange_rate_entry = Entry(self, bd=10, textvariable=self.exchange_rate_default)
        self.exchange_rate_entry.grid(row=5, column=9, pady=20, padx=25)

        #Get Button
        self.export_button = Button(self, text="Get", bd=10, pady=10, padx=25, font="Arial 10 bold", command=self.get)
        self.export_button.grid(row=4, column=10, columnspan=3, rowspan=2, pady=20, padx=25)

        #Lid Checkbutton
        self.lid_var = BooleanVar()
        self.exchange_rate_entry_label = Label(self, text="Lid?", font="Arial 15")
        self.exchange_rate_entry_label.grid(row=1, column=8, padx=25)
        self.lid_checkbutton = Checkbutton(self, bg="lightgray", variable=self.lid_var)
        self.lid_checkbutton.grid(row=2, column=8, padx=25, sticky="NSEW")

        #Seperator Checkbutton
        self.sep_var = BooleanVar()
        self.seperator_entry_label = Label(self, text="Seperator?", font="Arial 15")
        self.seperator_entry_label.grid(row=1, column=9, padx=25)
        self.seperator_checkbutton = Checkbutton(self, bg="lightgray", variable=self.sep_var)
        self.seperator_checkbutton.grid(row=2, column=9, padx=25, sticky="NSEW")

    def open_file(self):    # Import button functionality. User selects the pre created excel file for otomatic slot filling.
        file_path = tkFileDialog.askopenfilename()
        wb = xlrd.open_workbook(file_path)
        sheet = wb.sheet_by_index(0)

        total_row = sheet.nrows

        row_number = 0
        input_list = []

        while row_number < total_row:
            value = sheet.cell_value(row_number, 1)
            input_list.append(value)
            row_number += 1

        product.width = input_list[0]
        product.length = input_list[1]
        product.height = input_list[2]
        product.thickness = input_list[3]

        if input_list[4] == 0:
            product.lid = False
        else:
            product.lid = True
        if input_list[5] == 0:
            product.seprator = False
        else:
            product.seprator = True

        input.steel_price = input_list[6]
        input.try_usd = input_list[7]

        if product.width > product.length:
            self.width_default_text.set(product.length)
            self.length_default_text.set(product.width)

        else:
            self.width_default_text.set(product.width)
            self.length_default_text.set(product.length)

        self.height_default_text.set(product.height)
        self.thickness_default_text.set(product.thickness)

        if product.lid == True: self.lid_checkbutton.select()
        if product.seprator == True: self.seperator_checkbutton.select()

        self.current_steel_default.set(input.steel_price)
        self.exchange_rate_default.set(input.try_usd)

    def calc(self): # The calculation button functionality. Gets the information from all GUI slots, makes the operation and fills the total weight and price slots.

        if self.width_entry.get() == "": self.width_default_text.set(0)                 #
        if self.length_entry.get() == "": self.length_default_text.set(0)               #
        if self.height_entry.get() == "": self.height_default_text.set(0)               #
        if self.thickness_entry.get() == "": self.thickness_default_text.set(0)         # Sets values to zero if calculate button gets pressed wihout any value entered.
                                                                                        #
        if self.steel_price_entry.get() == "": self.current_steel_default.set(0)        #
        if self.exchange_rate_entry.get() == "": self.exchange_rate_default.set(0)      #

        product.width = float(self.width_entry.get())
        product.length = float(self.length_entry.get())
        product.height = float(self.height_entry.get())
        product.thickness = float(self.thickness_entry.get())

        input.steel_price = float(self.steel_price_entry.get())
        input.try_usd = float(self.exchange_rate_entry.get())

        if product.width > product.length:                          #
            product.width = float(self.width_entry.get())           #
            product.length = float(self.length_entry.get())         #
                                                                    #
            self.width_entry.delete(0, END)                         # If widt is larger than length changes their places.
            self.length_entry.delete(0, END)                        #
                                                                    #
            self.width_default_text.set(product.length)             #
            self.length_default_text.set(product.width)             #

        # From this point to end of this function, calculation process occurs.

        if self.lid_var.get() == True and self.sep_var.get() == False:
            input.surface_area = 0.0
            input.surface_area = 2.0 * ((product.width * product.height) + (product.length * product.height)) + (product.width * product.length) + (product.width * product.length)
        elif self.sep_var.get() == True and self.lid_var.get() == False:
            input.surface_area = 0.0
            input.surface_area = 2.0 * ((product.width * product.height) + (product.length * product.height)) + (product.width * product.length) + (product.width * product.height)
        elif self.sep_var.get() == True and self.lid_var.get() == True:
            input.surface_area = 0.0
            input.surface_area = 2.0 * ((product.width * product.height) + (product.length * product.height)) + (product.width * product.length) + (product.width * product.length) + (product.width * product.height)
        else:
            input.surface_area = 0.0
            input.surface_area = 2.0 * ((product.width * product.height) + (product.length * product.height)) + (product.width * product.length)

        input.volume = input.surface_area * product.thickness
        input.weight = (input.volume * 7.8) / 1000
        input.cost = input.weight * (input.steel_price / 1000) * input.try_usd

        self.total_weight_default.set(input.weight)     # Puts the results to
        self.total_price_default.set(input.cost)        # buttom two cells.

    def export(self):   # The export button functionality.
        currentDT = datetime.datetime.now()

        filename = 'SteelBox Inc. Results.xls'
        if os.path.exists(filename):
            os.remove(filename)

        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')

        sheet1.write(0, 0, "Date")
        sheet1.write(0, 1, currentDT.strftime("%d/%m/%Y"))

        sheet1.write(1, 0, "Time")
        sheet1.write(1, 1, currentDT.strftime("%H:%M"))

        sheet1.write(2, 0, "Width")
        sheet1.write(2, 1, float(self.width_entry.get()))

        sheet1.write(3, 0, "Length")
        sheet1.write(3, 1, float(self.length_entry.get()))

        sheet1.write(4, 0, "Height")
        sheet1.write(4, 1, float(self.height_entry.get()))

        sheet1.write(5, 0, "Thickness")
        sheet1.write(5, 1, float(self.thickness_entry.get()))

        sheet1.write(6, 0, "Lid Exist?")
        sheet1.write(6, 1, float(self.lid_var.get()))

        sheet1.write(7, 0, "Seprator Exist?")
        sheet1.write(7, 1, float(self.sep_var.get()))

        sheet1.write(8, 0, "Steels' Price in USD (per ton)")
        sheet1.write(8, 1, float(self.steel_price_entry.get()))

        sheet1.write(9, 0, "USD/TRY")
        sheet1.write(9, 1, float(self.exchange_rate_entry.get()))

        sheet1.write(10, 0, "Surface Area")
        sheet1.write(10, 1, float(self.surface_area))

        sheet1.write(11, 0, "Volume")
        sheet1.write(11, 1, float(self.volume))

        sheet1.write(12, 0, "Weight")
        sheet1.write(12, 1, float(self.weight))

        sheet1.write(13, 0, "Total Cost")
        sheet1.write(13, 1, float(self.total_price_entry.get()))

        wb.save(filename)


    def get(self):  # The get button functionality. Gets the currency and steel price data from internet and puts them into their places (or replaces the previous values if exists.)

        if self.exchange_rate_entry.get() == "":
            self.exchange_rate_entry.delete(0, END)

        if self.steel_price_entry.get() == "":
            self.steel_price_entry.delete(0, END)

        site = 'https://kur.doviz.com/'
        site2 = 'https://iscrapapp.com/metals/1-steel/'

        steel_fetch = []

        headers = {
            'User-Agent': (
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) '
                'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'
            )
        }

        r = requests.get(site, headers=headers)
        r2 = requests.get(site2, headers=headers)

        if r.status_code and r2.status_code != 200:
            print('Could Not Fetch Data From Website!')
        else:
            soup = bs(r.content, 'html.parser')
            soup2 = bs(r2.content, 'html.parser')

            value = soup.find_all(lambda tag: tag.name == 'span' and tag.get('class') == ['value'])
            currency = value[1].text

            for a in soup2.find_all('strong'):
                steel_fetch.append(a.text)

            price = float(steel_fetch[0][1:-4])
            money = ""

            i = 0
            while i < len(currency):

                if currency[i] != ",":
                    money += currency[i]
                    i += 1
                else:
                    money += "."
                    i += 1
            money = float(money)

            self.exchange_rate_default.set(money)
            self.current_steel_default.set(price)


root = Tk()
root.title("SteelBox Inc. Calculator")
app = SteelBox(root)
root.mainloop()