from src.billTable import *

testfile: str = "2016_1.txt"


def test_parseBills():
    bills = parseTxtFile(testfile)

    assert len(bills) == 11


def test_normalProductListNoNumericValue():
    bills = parseTxtFile(testfile)

    expectedDict = {
        'Beschreibung': 'RECHNUNG',
        'Nr.': 2,
        'ID-Nr.': 2,
        'Datum': '03.05.2016',
        'Produkt': 'Fuï¿½pflege_2023',
        'Anzahl': 1,
        'Einzelpreis': 28.0,
        'Gesamtpreis': 28.0
    }

    dictList: [{}] = transformBillsToDictList(bills)

    assert dictList[1]["Produkt"] == expectedDict["Produkt"]

def test_spaceAndNumericValueAtEndInProductList():
    bills = parseTxtFile(testfile)

    expectedDict = {
        'Beschreibung': 'RECHNUNG',
        'Nr.': 1,
        'ID-Nr.': 1,
        'Datum': '03.05.2016',
        'Produkt': 'Fuï¿½pflege 2023',
        'Anzahl': 1,
        'Einzelpreis': 28.0,
        'Gesamtpreis': 28.0
    }

    dictList: [{}] = transformBillsToDictList(bills)

    assert dictList[0]["Produkt"] == expectedDict["Produkt"]


def test_spaceAndNumericValueAtBeginInProductList():
    bills = parseTxtFile(testfile)

    expectedDict = {
        'Beschreibung': 'RECHNUNG',
        'Nr.': 1,
        'ID-Nr.': 1,
        'Datum': '03.05.2016',
        'Produkt': '2023 Fuï¿½pflege+Lack',
        'Anzahl': 1,
        'Einzelpreis': 29.0,
        'Gesamtpreis': 29.0
    }

    dictList: [{}] = transformBillsToDictList(bills)

    assert dictList[2]["Produkt"] == expectedDict["Produkt"]
