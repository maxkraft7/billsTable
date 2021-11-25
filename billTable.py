import os
import shutil
import sys
import traceback
import typing

import pandas as pd
from openpyxl import load_workbook


class Product:
    name: str = ""
    price: float = 0.0
    amount: int = 0
    total: float = 0.0


class ProductParser:

    @staticmethod
    def inferType(x: str) -> typing.Union[float, int, str]:

        if x.__contains__(','):  # could be a float
            try:
                x = x.replace(',', '.')
                return float(x)
            except ValueError:  # not a float
                return x
        else:
            try:
                return int(x)
            except ValueError:
                return x

    @staticmethod
    def fromSplittedString(rawData: [str]) -> [Product]:
        products: [Product] = []
        product = Product()

        state = "initial"

        for element in rawData:
            typedElement = ProductParser.inferType(element)

            if type(typedElement) is float:
                number = typedElement

                if state == "totalsBegin":
                    product.price = number
                    if product.amount != 0:
                        product.total = product.price * product.amount
                        products.append(product)
                        state = "initial"
                    else:
                        product.total = product.price * product.amount
                        products.append(product)
                        state = "initial"

                elif state == "totalsEnd":
                    product.total = product.price * product.amount
                    products.append(product)
                    state = "initial"
                pass
            elif type(typedElement) is int:

                if state == "initial":
                    product = Product()
                    product.amount = typedElement
                    state = "productName"
                pass
            else:
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
        # get first occurrence of Nr.
        nrLocus = rawData.index('Nr.')
        try:
            self.id = int(rawData[nrLocus + 1])
            self.date = rawData[nrLocus + 3]
            self.time = rawData[nrLocus + 5]

            # find all idxs that contain `'___________________________________________'`
            seperatorIdxs = [i for i, x in enumerate(rawData) if x == '___________________________________________']

            for i in range(0, len(seperatorIdxs), 2):
                self.products = ProductParser.fromSplittedString(rawData[seperatorIdxs[i] + 1:seperatorIdxs[i + 1]])
        except ValueError:
            print("Fehler beim Rechnung parsen!")


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


def injectFormulas(parsedData: pd.DataFrame) -> pd.DataFrame:
    # !important use english formula names here!
    formulas: pd.DataFrame = pd.DataFrame({
        "": [""],  # placeholder
        # "summe": ["=sum(A:A)"]  # sum of all ids (test)
    })

    merged = pd.concat(
        [parsedData, formulas],
        axis=1,
        join="outer",
        ignore_index=False,
        keys=None,
        levels=None,
        names=None,
        verify_integrity=False,
        copy=True,
    )

    return merged


def addImportedDataToTemplate(targetPath: str, imported: pd.DataFrame):
    book = load_workbook(targetPath)  # already contains `Berechnung` Sheet
    # https://stackoverflow.com/a/61364633/11466033
    writer = pd.ExcelWriter(targetPath,
                            engine_kwargs={'options': {'strings_to_formulas': False}}
                            )
    writer.book = book

    imported.to_excel(writer,  # add `Import` to file with `Berechnung` already in it
                      sheet_name="Import",
                      index_label="ID",
                      index=False,
                      freeze_panes=(1, 0))

    writer.save()
    writer.close()

    print(f"Excel Datei gespeichert unter {targetFile}")


def transformBillsToTable(bills: [Bill], targetPath: str, save: bool = True) -> pd.DataFrame:
    dicts = []

    for bill in bills:
        bill: Bill = bill
        dicts += bill.toDataframeRows()
    # https://stackoverflow.com/a/47561390/11466033
    df: pd.DataFrame = pd.DataFrame(dicts)

    return df


def createTargetFile(templatePath: str, destinationPath: str) -> str:
    destinationPath = os.path.splitext(destinationPath)[0] + ".xlsx"

    shutil.copy2(templatePath, destinationPath)

    return destinationPath


if __name__ == '__main__':
    try:
        if len(sys.argv) < 2:
            print("Keine Datei angegeben!")
        else:
            if os.path.isfile(sys.argv[1]):
                bills = parseTxtFile(sys.argv[1])
                table = transformBillsToTable(bills, sys.argv[1])
                # table = injectFormulas(table) # can be used to directly add formulas to import-df
                targetFile: str = createTargetFile(
                    os.path.dirname(os.path.realpath(__file__)) + "\\template.xlsx",
                    os.path.realpath(sys.argv[1])
                )
                addImportedDataToTemplate(targetFile, table)
            else:
                print("Angegebener Pfad ungültig!")
    except Exception as e:
        print(traceback.format_exc())

    os.system('pause')
