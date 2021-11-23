import os
import sys
import typing

import pandas as pd


class Product:
    name: str = ""
    price: float = 0.0
    amount: int = 0
    total: float = 0.0


class ProductParser:
    @staticmethod
    def fromSplittedString(rawData: [str]) -> [Product]:
        products: [Product] = []
        product = Product()

        state = "initial"

        for element in rawData:

            if "," in element:  # float locale
                element = element.replace(",", ".")

            try:  # if element is a number
                number = float(element)

                if state == "initial":
                    product = Product()
                    product.amount = number
                    state = "productName"

                elif state == "totalsBegin":
                    product.price = number
                    if product.amount > 1:
                        state = "totalsEnd"
                    else:
                        product.total = product.price * product.amount
                        products.append(product)
                        state = "initial"

                elif state == "totalsEnd":
                    product.total = product.price * product.amount
                    products.append(product)
                    state = "initial"

            except ValueError:  # element is a string
                if state == "productName" or state == "totalsBegin":
                    if product.name == "":
                        product.name = element
                    else:
                        product.name += " " + element
                    state = "totalsBegin"
        return products


class Bill:
    date: str = ""
    time: str = ""
    id: int = 0
    products: [Product] = []
    total: float = 0.0

    def toDataframeRows(self) -> [typing.Dict]:

        dfRows = []

        for idx, product in enumerate(self.products):
            dfRow = {"ID": self.id,
                     "Datum": self.date,
                     "Produkt": product.name,
                     "Anzahl": product.amount,
                     "Einzelpreis": product.price,
                     "Gesamtpreis": product.total}

            dfRows.append(dfRow)

        return dfRows

    def fromSplittedString(self, rawData: [str]):
        self.id = int(rawData[2])
        self.date = rawData[4]
        self.time = rawData[6]

        # find all idxs that contain `'___________________________________________'`
        seperatorIdxs = [i for i, x in enumerate(rawData) if x == '___________________________________________']

        for i in range(0, len(seperatorIdxs), 2):
            product: Product = Product()
            self.products = ProductParser.fromSplittedString(rawData[seperatorIdxs[i] + 1:seperatorIdxs[i + 1]])


def parseTxtFile(filepath: str) -> [Bill]:
    parsedBills: [Bill] = []

    # open file
    with open(filepath, 'r') as fileHandler:
        fileContent = fileHandler.read()  # read entire file content

        bills = fileContent.split("\n\n   \n")  # split bills by separator
        bills.pop(0)  # remove header

        for bill in bills:
            bill: str = bill
            billFragments = bill.split()  # split by any separator
            bill: Bill = Bill()
            bill.fromSplittedString(billFragments)
            parsedBills.append(bill)

    return parsedBills


def transformBillsToTable(bills: [Bill], targetPath: str, save: bool = True) -> pd.DataFrame:
    dicts = []

    for bill in bills:
        bill: Bill = bill
        dicts += bill.toDataframeRows()
    # https://stackoverflow.com/a/47561390/11466033
    df: pd.DataFrame = pd.DataFrame(dicts)

    if save:
        targetFile = os.path.splitext(targetPath)[0] + ".xlsx"
        df.to_excel(targetFile,
                    sheet_name="Zusammenfassung",
                    index_label="ID",
                    index=False,
                    freeze_panes=(1, 0))

        print(f"Excel Datei gespeichert unter {targetFile}")

    return df


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Keine Datei angegeben!")
    else:
        if os.path.isfile(sys.argv[1]):
            bills = parseTxtFile(sys.argv[1])
            table = transformBillsToTable(bills, sys.argv[1])
        else:
            print("Angegebener Pfad ungültig!")
