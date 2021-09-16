import win32com.client
import pythoncom

wincad = win32com.client.Dispatch("AutoCAD.Application")
doc = wincad.ActiveDocument
enter = chr(13)
# msp = doc.ModelSpace
print("現在使用的檔案: ", doc.Name)
"""
(LOAD \"C:/lisp/pass10.lsp\")
(LOAD \"C:/lisp/DCLTUBE3.lsp\")
test123
tube_arr3
"""
doc.SendCommand("(LOAD \"C:/lisp/pass10.lsp\")" + enter)
doc.SendCommand("(LOAD \"C:/lisp/DCLHEAT.lsp\")" + enter)
#doc.SendCommand("HEAT" + enter+"1"+enter)
