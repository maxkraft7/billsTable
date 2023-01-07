import os
import shutil
import sys
import typing
# todo show warnings
import warnings

import pandas as pd
from openpyxl import load_workbook

from enum import Enum


# class syntax

class ParserState(Enum):
    ITEM_AMOUNT = 1
    PRODUCT_NAME = 2
    TOTALS_BEGIN = 3
    TOTALS_END = 4


class Product:
    name: str = ""
    price: float = 0.0
    amount: int = 0
    totalPrice: float = 0.0


class ProductParser:

    @staticmethod
    def inferType(x: str) -> typing.Union[float, int, str]:

        # todo remove try/catch blocks
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
    def appendToProductName(element, product):
        if product.name == "":
            product.name = element
        else:
            product.name += " " + element

        return product

    @staticmethod
    def fromSplittedString(rawData: [str]) -> [Product]:
        products: [Product] = []
        product = Product()

        state: ParserState = ParserState.ITEM_AMOUNT

        for i, element in enumerate(rawData):
            typedElement = ProductParser.inferType(element)

            if type(typedElement) is float:
                number = typedElement

                if state == ParserState.TOTALS_BEGIN or state == ParserState.PRODUCT_NAME:
                    product.price = number
                    if product.amount != 0:
                        product.totalPrice = product.price * product.amount
                        products.append(product)
                        state = ParserState.ITEM_AMOUNT
                    else:
                        product.totalPrice = product.price * product.amount
                        products.append(product)
                        state = ParserState.ITEM_AMOUNT

                elif state == ParserState.TOTALS_END:
                    product.totalPrice = product.price * product.amount
                    products.append(product)
                    state = ParserState.ITEM_AMOUNT
                pass
            elif type(typedElement) is int:

                if state == ParserState.PRODUCT_NAME:
                    product = ProductParser.appendToProductName(element, product)
                    state = ParserState.PRODUCT_NAME

                if state == ParserState.ITEM_AMOUNT:
                    product = Product()
                    product.amount = typedElement
                    state = ParserState.PRODUCT_NAME
            else:
                # if the product name still gets concatenated and it's not yet the last element in line append it to
                #  the product name
                if state == ParserState.TOTALS_BEGIN and i != len(rawData) - 1:
                    product = ProductParser.appendToProductName(element, product)

                if state == ParserState.PRODUCT_NAME or state == ParserState.TOTALS_BEGIN:
                    product = ProductParser.appendToProductName(element, product)
                    state = ParserState.PRODUCT_NAME

        return products


class Bill:
    id: int = None
    description: str = None  # describes the type of bill
    billNr: int = None
    billId: int = None
    date: str = None
    products: [Product] = []
    time: str = None
    total: float = None

    # static
    csvcounter: int = 0

    def __init__(self):
        self.id = Bill.csvcounter
        Bill.csvcounter += 1

    def toDataframeRows(self) -> [typing.Dict]:

        dfRows = []

        for idx, product in enumerate(self.products):
            dfRow = {"CSV Zeilenzähler": self.id,
                     "Beschreibung": self.description,
                     "Nr.": self.billNr,
                     "ID-Nr.": self.billId,
                     "Datum": self.date,  # todo parse from bill
                     "Produkt": product.name,
                     "Anzahl": product.amount,
                     "Einzelpreis": product.price,
                     "Gesamtpreis": product.totalPrice}

            dfRows.append(dfRow)

        return dfRows

    def fromSplittedString(self, rawData: [str]) -> bool:

        if len(rawData) == 0:
            return False

        # get first occurrence of Nr.
        nrLocus = rawData.index('Nr.')

        self.billNr = int(rawData[nrLocus + 1])
        self.billId = int(rawData[nrLocus + 8])
        self.date = rawData[nrLocus + 3]
        self.time = rawData[nrLocus + 5]
        self.description = rawData[0].split(" ")[0]

        # find all idxs that contain `'___________________________________________'`
        seperatorIdxs = [i for i, x in enumerate(rawData) if x == '___________________________________________']

        for i in range(0, len(seperatorIdxs), 2):
            self.products = ProductParser.fromSplittedString(rawData[seperatorIdxs[i] + 1:seperatorIdxs[i + 1]])

        return True


def parseTxtFile(filepath: str) -> [Bill]:
    parsedBills: [Bill] = []

    # open file
    with open(filepath, 'r') as fileHandler:
        fileContent = fileHandler.read()  # read entire file content

        bills = fileContent.split("\n\n   \n")  # split bills by separator
        bills.pop(0)  # remove header

        for rawBillString in bills:
            rawBillString: str = rawBillString
            billFragments = rawBillString.split()  # split by any separator
            parsedBillObject: Bill = Bill()
            if parsedBillObject.fromSplittedString(billFragments):
                parsedBills.append(parsedBillObject)

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


def addImportedDataToTemplate(targetPath: str, extension: str, imported: pd.DataFrame, extensionExcel: str = ".xlsx"):
    if extension == extensionExcel:
        # excel mode
        book = load_workbook(targetPath)  # already contains `Berechnung` Sheet
        # https://stackoverflow.com/a/61364633/11466033
        writer = pd.ExcelWriter(targetPath,
                                mode='w')  # if_sheet_exists='overlay',  ,engine_kwargs={'options': {'strings_to_formulas': False}}
        writer.book = book

        imported.to_excel(writer,  # add Import to file with `Berechnung` already in it
                          sheet_name="Import",
                          index_label="ID",
                          index=False,
                          freeze_panes=(1, 0))

        writer.save()
        writer.close()

        print(f"Excel Datei gespeichert unter {targetFile}")
    else:
        # csv mode
        imported.to_csv(targetPath, index=False)


def transformBillsToDictList(bills: [Bill]) -> [{}]:
    dicts = []

    for bill in bills:
        bill: Bill = bill
        dicts += bill.toDataframeRows()

    return dicts


def transformBillsToTable(bills: [Bill]) -> pd.DataFrame:
    # https://stackoverflow.com/a/47561390/11466033
    df: pd.DataFrame = pd.DataFrame(transformBillsToDictList(bills))

    return df


def createTargetFile(templatePath: str, destinationPath: str, extensionExcel: str = ".xlsx",
                     extensionCsv: str = ".csv") -> typing.Tuple[str, str]:
    # varations: .xlsx, .csv
    if os.path.exists(templatePath + extensionExcel):
        # excel-mode
        destinationPath = os.path.splitext(destinationPath)[0] + extensionExcel

        shutil.copy2(templatePath + extensionExcel, destinationPath)

        return destinationPath, extensionExcel
    else:
        # csv mode
        destinationPath = os.path.splitext(destinationPath)[0] + extensionCsv

        # shutil.copy2(templatePath, destinationPath)

        return destinationPath, extensionCsv


if __name__ == '__main__':

    # ignore pandas warnings for now
    # todo fix deprecation in future versions
    warnings.filterwarnings('ignore')

    if len(sys.argv) < 2:
        print("Keine Datei angegeben!")
    else:
        templateFileName = "template"

        if os.path.isfile(sys.argv[1]):
            bills = parseTxtFile(sys.argv[1])
            table: pd.DataFrame = transformBillsToTable(bills)
            # table = injectFormulas(table) # can be used to directly add formulas to import-df
            targetFile, extension = createTargetFile(
                os.path.dirname(os.path.realpath(__file__)) + "\\" + templateFileName,
                os.path.realpath(sys.argv[1]),
                ".xlsm"
            )
            addImportedDataToTemplate(targetFile, extension, table, ".xlsm")
        else:
            print("Angegebener Pfad ungültig!")

os.system('pause')
