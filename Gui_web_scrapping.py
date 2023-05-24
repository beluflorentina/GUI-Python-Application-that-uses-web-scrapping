import tkinter
from tkinter import filedialog
import matplotlib.pyplot as plt
import pandas
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager



def retrieve_data():
    driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.implicitly_wait(30)
    driver.get("https://www.cabelas.com/l/trolling-rods#numberOfResults=32")


    global products
    products = []
    global prices
    prices = []

    elementHTML=driver.find_element("id","main")
    children_element=elementHTML.find_elements("class name", "styles_ItemDetails__MtDO0" )

    for child_element in children_element:

        #products name
        title=child_element.find_element("class name","styles_ResultItemTitleLink___CgWd").get_attribute("innerText")
        products.append(title)

        #product prices
        title=child_element.find_element("class name", "styles_ResultItemPrice__vsgH5").get_attribute("innerText")

        #formatting prices
        dash_position=title.find("-")
        title=title[dash_position+1:]
        title=title.replace("$", "")
        title=title.replace(",", "")
        prices.append(float(title))

    print(products, prices)


def display_chart():
    prd_array=pandas.Series(products)
    prc_array=pandas.Series(prices)

    prod_prc={"Product_description":prd_array, "Product_price":prc_array}
    prd_DF=pandas.DataFrame(prod_prc)

    prd_DF.plot.barh(x="Product_description")

    plt.show()

def display_matrix():
    prd_array = pandas.Series(products)
    prc_array = pandas.Series(prices)

    prod_prc = {"Product_description": prd_array, "Product_price": prc_array}
    prd_DF = pandas.DataFrame(prod_prc)
    print(prd_DF)

def save_to_excel():
    name=text_name.get()

    prd_array = pandas.Series(products)
    prc_array = pandas.Series(prices)

    prod_prc = {"Product_description": prd_array, "Product_price": prc_array}
    prd_DF = pandas.DataFrame(prod_prc)
    prd_DF.to_excel(filedialog.asksaveasfilename( initialfile=name), sheet_name=text_name.get())


form=tkinter.Tk()
label_name=tkinter.Label(form, text="Write the name of the excel file to save with .xlsx", fg="green")

text_name=tkinter.Entry(form, bg="lightgreen")
label_name.pack()
text_name.pack()

ButtonRetrieveData=tkinter.Button(form, text="Retrieve data", command=retrieve_data, bg="#56450e", fg="white")
ButtonCreateTheGraph=tkinter.Button(form, text="Create the graph", command=display_chart, bg="#56450e", fg="white")
ButtonDisplayTheMatrix=tkinter.Button(form, text="Display the matrix", command=display_matrix, bg="#56450e", fg="white")
ButtonSaveToExcelFile=tkinter.Button(form, text="Save to Excel File", command=save_to_excel, bg="#56450e", fg="white")

ButtonRetrieveData.pack()
ButtonCreateTheGraph.pack()
ButtonDisplayTheMatrix.pack()
ButtonSaveToExcelFile.pack()

form.title("Product and price")
form.geometry("600x200")
form.mainloop()