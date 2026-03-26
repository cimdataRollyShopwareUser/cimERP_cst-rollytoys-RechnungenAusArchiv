import re
import os
import base64
import PyPDF2
import datetime
import requests
import win32clipboard
import xml.etree.ElementTree as ET

myPid = os.getpid()
easyUser = "cd2000"
easyPass = "cd2000"
easyCd2000Archivname = "CD2000"
easyCd2000Schemaname = "CD2000"
easyURL = "http://snarchiv-1:9090/eex-xmlserver/eex-xmlserver"


def main():
    outFileName = "out.pdf"
    merger = PyPDF2.PdfMerger()
    rgNums = getRgNrFromClipboard()
    pdfPaths = getPfdsFromArchive(rgNums)
    printLog("Merging Files...")
    for pdfPath in pdfPaths:
        merger.append(pdfPath)
    merger.write(outFileName)
    merger.close()
    printLog(f"{outFileName} written. Done!")


def getRgNrFromClipboard():
    win32clipboard.OpenClipboard()
    try:
        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()
    return data.split("\r\n")[:-1]


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
    printLog(f"Easy Login Context ID: {contextid}")
    return contextid


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
    payload = f"""<REQUEST XMLID="XMLID" CONTEXTID="{contextid}"><QUERY REQUESTID="0" HITPOSITION="1" MAXHITCOUNT="2000"><EQL>SELECT * FROM /{easyCd2000Archivname} WHERE {easyCd2000Schemaname}.Belegnummer = '{belegnr}'</EQL></QUERY></REQUEST>"""
    response = requests.post(url, data=payload, headers=headers, verify=False)
    root = ET.fromstring(response.text)
    hitline = root.find(".//HITLINE")
    return hitline.get("EASYDOCREF")


def logoffEasy(url, contextid):
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    payload = f"""<?xml version="1.0" encoding="UTF-8"?><REQUEST XMLID="XMLID" CONTEXTID="{contextid}"><LOGOUT REQUESTID="0"/></REQUEST>"""
    response = requests.post(url, data=payload, headers=headers, verify=False)


def getPfdsFromArchive(rgNums):
    contextid = logonEasy(easyURL, easyUser, easyPass)
    exportedDocuments = []
    for num, rgNum in enumerate(rgNums):
        fileName = f"{rgNum}.pdf"
        printLog(f"Downloading {rgNum} as {fileName} ({num + 1}/{len(rgNums)})...")
        docEasyLink = searchForBelegNr(easyURL, contextid, rgNum)
        getDocumentfromEasy(easyURL, contextid, docEasyLink, fileName)
        exportedDocuments.append(fileName)
    logoffEasy(easyURL, contextid)
    return exportedDocuments


def printLog(logtext: str) -> None:
    print(datetime.datetime.now().astimezone().isoformat() + " (" + str(myPid) + ") => " + logtext, flush=True)


if __name__ == '__main__':
    main()
