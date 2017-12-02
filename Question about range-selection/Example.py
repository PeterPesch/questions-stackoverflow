"""Reading a Time Schedule - Read Per Category."""

from __future__ import print_function

import os
import win32com.client
import win32gui

import time         # Just added for the code example

def areas2range(Excel, Areas):
    """Convert an Areas to a Range"""
    if not Areas:
        return None
    result = Areas.Item(1)
    for i in range(2, Areas.Count+1):
        result = Excel.Union(result, Areas.Item(i))
    return result

def chronoloog():
    """Reading a Track and Field Time Schedule.
    
    This version lets the user choose an input file.
    After that, the user can select a combination of ranges which contains
    the Time Schedule per (age/gender) Category.
    In the next step, the user can specify several parts of the Time Schedule
    which contain certain types of information ...
    (Rest of the program has been left out).
    """
    invoerFolder = r"C:\temp"
    macroBestand = r"C:\temp\Chronoloog_Hulpfuncties.xlsm"
    vraagRange = "Chronoloog_Hulpfuncties.xlsm!vraagRange"
    vraagFilename = "Chronoloog_Hulpfuncties.xlsm!vraagFilename"

    # Start a new Excel instance
    xlApp = win32com.client.DispatchEx("Excel.Application")
    # xlApp = win32com.client.Dispatch("Excel.Application")
    win32gui.SetForegroundWindow(xlApp.HWND)
    xlApp.Visible = 1

    # Open the Macro file
    try:
        mcr = xlApp.Workbooks.Open(macroBestand, ReadOnly=1)
        mcr.Windows(1).Visible = False     #   Can probably be left out after the first run ...
    except:
        print('Error while trying to open "' + macroBestand + '"')
        raise

    invoerBestand = xlApp.Run(vraagFilename, invoerFolder,
                              "Select the Time Schedule file")

    if not invoerBestand:
        print("No File chosen")
    else:
        # Open an Excel file
        wb = xlApp.Workbooks.Open(invoerBestand, ReadOnly=1)
        print(wb.Name)
        print('')
        # Choose Areas for the categorie-blokjes in the workbook (Compulsary)
        try:
            catBlokjes = xlApp.Run(vraagRange, xlApp,
                                "Schedule per Category", "Choose a Range")
        except:
            mcr.close
            mcr = None
            xlApp.Quit()
            xlApp is None
            raise
        else:
            xlApp.Goto(areas2range(xlApp, catBlokjes))
        finally:
            pass
        if catBlokjes is not None:
            # Highlight Selection, Hide rest of sheet
            ws = areas2range(xlApp, catBlokjes).Worksheet
            ws.UsedRange.Interior.ColorIndex                        = -4142     # ColorIndex None
            ws.UsedRange.Font.Color                                 = 0xFFFFFF  # wit
            areas2range(xlApp, catBlokjes).Interior.Color           = 0xFFFFCC  # lichtblauw
            areas2range(xlApp, catBlokjes).Font.Color               = 0x000000  # zwart
        # Now let the User specify all parts of the selection:
        try:
            catCategorie = xlApp.Run(vraagRange, xlApp,
                                "Specifying Schedule per Category",
                                "Select all cells which contain a Category")
        finally:
            pass
        if catCategorie is not None:
            # Remove Highlight from specified cells
            areas2range(xlApp, catCategorie).Interior.ColorIndex    = -4142     # ColorIndex None
        try:
            catTijden = xlApp.Run(vraagRange, xlApp,
                                "Specifying Schedule per Category",
                                "Select all cells which contain a Start Time")
        finally:
            pass
        if catTijden is not None:
            areas2range(xlApp, catTijden).Interior.ColorIndex       = -4142     # ColorIndex None
        try:
            catOnderdelen = xlApp.Run(vraagRange, xlApp,
                                "Specifying Schedule per Category",
                                "Select all cells which contain an Event")
        finally:
            pass
        # etcetera ...
        # ...
        # Close the excel file
        wb.Saved = True
        wb.close
        wb = None

    # Close the macro file
    mcr.close
    mcr = None

    # close the excel application
    xlApp.Quit()
    xlApp = None


if __name__ == "__main__":
    try:
        chronoloog()
    except Exception:
        print("Please note that this example code is not ripe for production ...")
    print("End of example run.")
    time.sleep(3)
