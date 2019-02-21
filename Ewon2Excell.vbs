'
' Ewon2Excell.vbs 0.1 Adapt EWON files to Excell
' (c) Iker De Echaniz for Grupo Elektra
' 19/02/2019 under GPL license
'
' Requires Windows Scripting Host and Microsoft Excell
' It won't work under LibreOffice
'
Help="Ewon2Excell.vbs 0.1 Adaptador de ficheros EWON a Excell"+vbcrlf+vbcrlf
Help=Help+ "¡Necesito un fichero CSV o de texto como parametro!"+vbcrlf+vbcrlf
Help=Help+ "(c) Iker De Echaniz para Grupo Elektra"+vbcrlf 

WScript.Interactive = True

If Wscript.Arguments.Count = 0 Then
  Wscript.Echo Help
  WScript.Quit 1 'Exit with errorlevel 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False

' Documentation:
'https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.opentext
'https://docs.microsoft.com/en-us/office/vba/api/excel.xlcolumndatatype

'expression. OpenText( _Filename_ , _Origin_ , _StartRow_ , _DataType_ , 
                       '_TextQualifier_ , _ConsecutiveDelimiter_ , _Tab_ , _Semicolon_ , _Comma_ , _Space_ ,
'                      _Other_ , _OtherChar_ , _FieldInfo_ , _TextVisualLayout_ , 
'                      _DecimalSeparator_ , _ThousandsSeparator_ , _TrailingMinusNumbers_ , _Local_ )

Const xlDelimited                =  1
Const xlTextQualifierDoubleQuote =  1

xl.Workbooks.OpenText WScript.Arguments.Item(0), , , xlDelimited _
  , xlTextQualifierDoubleQuote, True, False, True, False, False, False, _
  , Array(Array(1,1), Array(2,2))


'Get the filename to be saved
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFile=FSO.GetFile(WScript.Arguments.Item(0))
sPath= Left(objFile.Path, Len(objFile.Path)-Len(objFile.Name))
newFileFullPath=sPath+FSO.GetBaseName(WScript.Arguments.Item(0))+".xlsx"
Set objFile = Nothing
Set FSO = Nothing

' If you have an error here the document is probably already open in use.
Set wb = xl.ActiveWorkbook
wb.SaveAs newFileFullPath, 51, , , , False

wb.Close
xl.Quit
set xl = Nothing
