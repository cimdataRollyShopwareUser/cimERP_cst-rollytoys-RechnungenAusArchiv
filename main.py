import re
import os
import base64
import PyPDF2
import openpyxl
import datetime
import requests
import win32clipboard
import xml.etree.ElementTree as ET

myPid = os.getpid()


def main():
    merger = PyPDF2.PdfMerger()
    # rgNums = getRgNrFromExcel("PayPal Zahlungseingang 16.03.2026.xlsx")
    rgNums = getRgNrFromClipboard()
    pdfPaths = getPfdsFromArchive(rgNums)
    printLog("Merging Files...")
    for pdfPath in pdfPaths:
        merger.append(pdfPath)
    merger.write("out.pdf")
    merger.close()
    printLog("Done!")


def getRgNrFromClipboard():
    win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()

    return data.split("\r\n")


def getRgNrFromExcel(pathToExcel):
    wb = openpyxl.load_workbook(pathToExcel)
    ws = wb.active

    rgNums = []
    for row in ws.iter_rows(min_col=4, max_col=4, min_row=2, values_only=True):
        rgNums.append(row[0])

    return rgNums


def logonEasy(url, user, pw):
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    payload = f"""<?xml version="1.0" encoding="UTF-8"?>
                    <REQUEST XMLID="XMLID">
                        <LOGIN REQUESTID="0">
                            <USERNAME>{user}</USERNAME>
                            <PASSWORD CRYPT="NONE">"{base64.b64encode(pw.encode("utf-8")).decode()}"</PASSWORD>
						</LOGIN>
					</REQUEST>"""
    response = requests.post(url, data=payload, headers=headers, verify=False)
    contextid = re.search(r'CONTEXTID="([^"]+)"', response.text).group(1)
    printLog(f"Context ID: {contextid}")
    return contextid


def getEasyDocIdFromDatabase(db, rgNum):
    with db.cursor() as cur:
        cur.execute(f"SELECT easy_docid FROM easy_schnittstelle WHERE belegnummer = {rgNum}")
        return cur.fetchall()[0][0]


def getDocumentfromEasy(url, contextid, docEasyLink, fileName):
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    payload = f"""<?xml version="1.0" encoding="UTF-8"?><REQUEST XMLID="XMLID" CONTEXTID="{contextid}"><DOCUMENT REQUESTID="0" EASYDOCREF="{docEasyLink}" BLOBDATA="1" BLOBID="" INTFIELDS="0" FIELDID="" RENDER="0" IFRCCODEB64="0"/></REQUEST>"""
    response = requests.post(url, data=payload, headers=headers, verify=False)
    root = ET.fromstring(response.text)
    blob_field = root.find('.//FIELD[@TYPE="BLOB"]')
    if blob_field is None:
        raise ValueError("No BLOB field found")
    data = blob_field.findtext("DATA")
    if not data:
        raise ValueError("No DATA found in BLOB field")
    with open(fileName, "wb") as f:
        f.write(base64.b64decode(data.strip()))
    printLog(f"{fileName} saved!")


def searchForBelegNr(url, contextid, belegnr):
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    payload = f"""<REQUEST XMLID="XMLID" CONTEXTID="{contextid}"><QUERY REQUESTID="0" HITPOSITION="1" MAXHITCOUNT="2000"><EQL>SELECT * FROM /CD2000 WHERE CD2000.Belegnummer = '{belegnr}'</EQL></QUERY></REQUEST>"""
    response = requests.post(url, data=payload, headers=headers, verify=False)
    root = ET.fromstring(response.text)
    hitline = root.find(".//HITLINE")
    return hitline.get("EASYDOCREF")


def logoffEasy(url, contextid):
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    payload = f"""<?xml version="1.0" encoding="UTF-8"?><REQUEST XMLID="XMLID" CONTEXTID="{contextid}"><LOGOUT REQUESTID="0"/></REQUEST>"""
    response = requests.post(url, data=payload, headers=headers, verify=False)


def getPfdsFromArchive(rgNums):
    contextid = logonEasy("http://snarchiv-1:9090/eex-xmlserver/eex-xmlserver", "cd2000", "cd2000")
    exportedDocuments = []
    # db = pypyodbc.connect("DSN=betrieb01")
    for num, rgNum in enumerate(rgNums):
        # docEasyLink = getEasyDocIdFromDatabase(db, rgNum)
        docEasyLink = searchForBelegNr("http://snarchiv-1:9090/eex-xmlserver/eex-xmlserver", contextid, rgNum)
        getDocumentfromEasy("http://snarchiv-1:9090/eex-xmlserver/eex-xmlserver", contextid, docEasyLink, f"{num:010}.pdf")
        exportedDocuments.append(f"{num:010}.pdf")
    # db.close()
    logoffEasy("http://snarchiv-1:9090/eex-xmlserver/eex-xmlserver", contextid)
    return exportedDocuments


def printLog(logtext: str) -> None:
    print(datetime.datetime.now().astimezone().isoformat() + " (" + str(myPid) + ") => " + logtext, flush=True)


if __name__ == '__main__':
    main()
