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

    def calculateTotalPrice(self):
        self.totalPrice = self.price * self.amount


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
    def fromSplittedString(rawData: [str], billLines: [str], currentIdx: int) -> typing.Optional[Product]:

        # overflow line with just the price
        if len(rawData) < 3:
            return None

        product: Product = Product()
        product.amount = int(rawData[0])

        priceCandidate: str = rawData[-1].replace(',', '.')

        if priceCandidate.replace('.', '').isdigit():
            # last arg is price
            pass
        else:
            # price is only arg in next line
            priceCandidate = billLines[currentIdx + 1].replace(',','.')

        product.price = float(priceCandidate)

        for s in rawData[1:-1]:
            product = ProductParser.appendToProductName(s, product)

        product.calculateTotalPrice()

        return product


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

    def fromSplittedString(self, billFragments: [str], billLines: [str]) -> bool:

        if len(billFragments) == 0:
            return False

        # get first occurrence of Nr.
        nrLocus = billFragments.index('Nr.')

        self.billNr = int(billFragments[nrLocus + 1])
        self.billId = int(billFragments[nrLocus + 8])
        self.date = billFragments[nrLocus + 3]
        self.time = billFragments[nrLocus + 5]
        self.description = billFragments[0].split(" ")[0]

        # find all idxs that contain `'___________________________________________'`
        seperatorIdxs = [i for i, x in enumerate(billLines) if x == '___________________________________________']

        for i in range(seperatorIdxs[0] + 1, seperatorIdxs[1]):

            optionalProduct: typing.Optional[Product] = ProductParser.fromSplittedString(billLines[i].split(), billLines, i)

            if optionalProduct is not None:
                self.products.append(optionalProduct)

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
            if parsedBillObject.fromSplittedString(billFragments, rawBillString.splitlines()):
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
