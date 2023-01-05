from src.billTable import *

testfile: str = "2016_1.txt"


def test_parseBills():
    bills = parseTxtFile(testfile)

    assert len(bills) == 11


def test_spaceAndNumericValueInProductList():
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