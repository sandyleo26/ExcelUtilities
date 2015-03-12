Function SelectNcopy() As String
' select first visible cell in ISBN column
Range([A2], Cells(Rows.Count, "A")).SpecialCells(xlCellTypeVisible)(1).Select
firstrow = ActiveCell.Row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' check the order type Local/Air or Sea
Dim tmprng As Range, c As Range, isSeaOrder As Boolean, isLocalOrAir As Boolean
isSeaOrder = True
isLocalOrAir = True

Set Ret = Range("A:Z").Find("ORDER", SearchOrder:=xlByRows, LookAt:=xlWhole)
If firstrow = lastrow Then
    Set tmprng = Union(Cells(firstrow, Ret.Column), Cells(1, Ret.Column))
Else
    Set tmprng = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))
End If
For Each c In tmprng.SpecialCells(xlCellTypeVisible)
    If c = "" Then
        isLocalOrAir = False
    End If
Next

Set Ret = Range("A:Z").Find("SEA", SearchOrder:=xlByRows, LookAt:=xlWhole)
If firstrow = lastrow Then
    Set tmprng = Union(Cells(firstrow, Ret.Column), Cells(1, Ret.Column))
Else
    Set tmprng = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))
End If
For Each c In tmprng.SpecialCells(xlCellTypeVisible)
    If c = "" Then
        isSeaOrder = False
    End If
Next

If isSeaOrder And isLocalOrAir Then
    MsgBox "Cannot decide it is Local/Air or Sea order. Choose Manually."
ElseIf Not isSeaOrder And Not isLocalOrAir Then
    MsgBox "This is neither a Local/Air nor Sea order. Please check."
ElseIf isSeaOrder Then
    t = "SEA"
ElseIf isLocalOrAir Then
    t = "ORDER"
End If

' Select and copy ISBN, Title, Quantity, Status and Date
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
firstrow = ActiveCell.Row
Dim rng1 As Range, rng2 As Range, rng3 As Range, rng4 As Range, rng5 As Range, rngUnion As Range

If t <> "ORDER" And t <> "SEA" And t <> "HK" Then
    MsgBox "Unknow Order Type: " & t
End If

' Select ISBN
Set Ret = Range("A:Z").Find("ISBN", SearchOrder:=xlByRows, LookAt:=xlWhole)
Set rng1 = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))

'Title
Set Ret = Range("A:Z").Find("Title", SearchOrder:=xlByRows, LookAt:=xlWhole)
Set rng2 = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))

' Quantity: ORDER/SEA/HK
Set Ret = Range("A:Z").Find(t, SearchOrder:=xlByRows, LookAt:=xlWhole)
Set rng3 = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))

'Status
Set Ret = Range("A:Z").Find("STATUS", SearchOrder:=xlByRows, LookAt:=xlWhole)
Set rng4 = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))

'Date
Set Ret = Range("A:Z").Find("DATE", SearchOrder:=xlByRows, LookAt:=xlWhole)
Set rng5 = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))

Set rngUnion = Union(rng1, rng2, rng3, rng4, rng5)
rngUnion.Select
Selection.copy

SelectNcopy = t

End Function

Sub copy_from_order_to_template()

Workbooks.Open filename:="J:\NewSouth Books\Inventory\Templates\Purchase Order Template.xltx"
'Windows("Purchase Order Template1").Activate
    Range("A23").Select
    'ActiveWindow.SmallScroll Down:=12
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                
End Sub
Sub make_total_line()

lastrow_plus = Cells(Rows.Count, 1).End(xlUp).Row + 1
Range(Cells(lastrow_plus, 1), Cells(lastrow_plus, 6)).Select
Call total_line

End Sub
Sub MakeOrderStep1()

'todo: automatically check website about the status, release date and title of an ISBN so that filling in orders will be easier
'todo: HK order?
' get publiser
Dim pub As String, ABK As String, country As String, freight As String, tld As Range

Set Ret = Range("A:Z").Find("PUB", SearchOrder:=xlByRows, LookAt:=xlWhole)
pub_column = Ret.Column
lastrow = Cells(Rows.Count, pub_column).End(xlUp).Row
pub = Cells(lastrow, pub_column)
pub = Replace(pub, " ", "")

t = SelectNcopy()

If pub = "UNB" Then
    pub = "SOS"
End If
If pub = "PLG" Then
    pub = "BIR"
End If

Call copy_from_order_to_template

Call make_total_line

' fill A6 the ABK publications
Set tld = Range("I7:I207").Find(pub, SearchOrder:=xlByRows)
ABK = Cells(tld.Row, tld.Column - 1)
If ABK = "" Then
  Exit Sub
End If

Cells(6, 1) = ABK
''freight = Cells(tld.Row, tld.Column + 2)
country = Cells(tld.Row, tld.Column + 1)
country = Replace(country, " ", "")

' call USA air or UK air
If t = "ORDER" Then
    If country = "USA" Or country = "CAN" Then
        Call USA_Air
    ElseIf country = "UK" Or country = "FRA" Then
        Call UK_Air2
    ElseIf country = "SING" Then
        MsgBox "Singpore has sea freight only!"
    End If
ElseIf t = "SEA" Then
    If country = "USA" Or country = "CAN" Then
        Call USAseafreight
    ElseIf country = "UK" Or country = "FRA" Then
        Call UK_SEA
    ElseIf country = "SING" Then
        Call SINGsea
    ElseIf country = "AUS" Then
        MsgBox "Local publisher is using Sea order"
    End If
ElseIf t = "HK" Then
    Call HKsea
End If

End Sub
Sub main_sea()
Call MakeOrderStep1
End Sub
Sub main_local_air()
Call MakeOrderStep1
End Sub
Sub main_hk()
Call MakeOrderStep1
End Sub
Sub main_post()
Dim pub As String, country As String, freight As String, tld As Range

HK_DELIVERY = "Global Connect Cargo Logistics Ltd (Infinity Cargo Express)"
pub = Cells(2, 2)
Set tld = Range("I7:I207").Find(pub, SearchOrder:=xlByRows)
country = Cells(tld.Row, tld.Column + 1)

fname = Cells(2, 2) & Cells(1, 2)
If country = "USA" Or country = "CAN" Or country = "UK" Or country = "SING" Or country = "FRA" Then
    Call OS_PO_Format
ElseIf country = "AUS" Then
    Call Aust_P_O
End If

' set print area
lastrow = Cells(Rows.Count, 1).End(xlUp).Row + 1
Range(Cells(1, 1), Cells(lastrow, 6)).Select
ActiveSheet.PageSetup.PrintArea = Selection.Address
Cells(7, 1).Select
' save file
savepath = ""
allorderwbname = ""
For Each book In Workbooks
    If book.Name Like "[0-9]*.xlsx" Or book.Name Like "Source.xlsx" Then
        savepath = book.Path & "\"
        allorderwbname = book.Name
    End If
    
Next book

delivery_to = Range("A11")
If delivery_to = HK_DELIVERY Then
    fname = fname & "-HK Drop Ship"
End If

ActiveWorkbook.SaveAs filename:=savepath & fname

'TODO: automatically add PO number on OSA sheet
AIR_FREIGHT = "Air Freight"
SEA_FREIGHT = "SEA FREIGHT"
LOCAL_FREIGHT = "Local"
freight = Cells(8, 1)
orderwbname = ActiveWorkbook.Name
tmpoffset = 0
' todo: fix the name issue
If freight = SEA_FREIGHT Or freight = AIR_FREIGHT Then
    tmpoffset = 7
    Workbooks("MyOrders.xlsx").Sheets("OSA").Activate
Else
    tmpoffset = 4
    Workbooks("MyOrders.xlsx").Sheets("LOA").Activate
End If
If pub = "SOS" Then
    pub = "UNB"
End If

Set Ret = Range("C3:C100").Find(pub, SearchOrder:=xlByRows, LookAt:=xlWhole)
If Ret = "" Then
    MsgBox "cannot find pub:(" & pub & ")"
End If

If Cells(Ret.Row, Ret.Column + tmpoffset) = "" Then
    Cells(Ret.Row, Ret.Column + tmpoffset) = fname
    Cells(Ret.Row, Ret.Column + tmpoffset).Font.Name = "Times New Roman"
    Cells(Ret.Row, Ret.Column + tmpoffset).Font.Size = 9
'    If freight = AIR_FREIGHT Or freight = LOCAL_FREIGHT Then
'        Cells(Ret.Row, Ret.Column + tmpoffset).Font.Color = RGB(0, 123, 234)
'    Else
'        Cells(Ret.Row, Ret.Column + tmpoffset).Font.Color = RGB(94, 162, 38)
'    End If
Else
    Cells(Ret.Row, Ret.Column + tmpoffset) = Cells(Ret.Row, Ret.Column + tmpoffset) & ", " & fname
    'commapos = InStr(Cells(Ret.Row, Ret.Column + tmpoffset), ",")
'    commapos = InStrRev(Cells(Ret.Row, Ret.Column + tmpoffset), ",")
'    If freight = AIR_FREIGHT Or freight = LOCAL_FREIGHT Then
'        Cells(Ret.Row, Ret.Column + tmpoffset).Characters(commapos).Font.Color = RGB(0, 123, 234)
'    Else
'        Cells(Ret.Row, Ret.Column + tmpoffset).Characters(commapos).Font.Color = RGB(94, 162, 38)
'    End If
End If

Workbooks(orderwbname).Activate
Workbooks(allorderwbname).Activate
Workbooks(orderwbname).Activate

End Sub
Sub main_post_n_send()
Dim pub As String, country As String, freight As String, tld As Range
Dim specialworkbook As String, folder As String, fullname As String
Dim book As Workbook

'todo: 3. make it
HK_DELIVERY = "Global Connect Cargo Logistics Ltd (Infinity Cargo Express)"
pub = Cells(2, 2)
Set tld = Range("I7:I207").Find(pub, SearchOrder:=xlByRows)
country = Cells(tld.Row, tld.Column + 1)


fname = Cells(2, 2) & Cells(1, 2)
If country = "USA" Or country = "CAN" Or country = "UK" Or country = "SING" Or country = "FRA" Then
    Call OS_PO_Format
ElseIf country = "AUS" Then
    Call Aust_P_O
End If

' set print area
lastrow = Cells(Rows.Count, 1).End(xlUp).Row + 1
Range(Cells(1, 1), Cells(lastrow, 6)).Select
ActiveSheet.PageSetup.PrintArea = Selection.Address
Cells(7, 1).Select
' save file
savepath = ""
For Each book In Workbooks
    If book.Name Like "[0-9]*.xlsx" Or book.Name Like "Source.xlsx" Then
        savepath = book.Path & "\"
    End If
    
Next book

delivery_to = Range("A11")
If delivery_to = HK_DELIVERY Then
    fname = fname & "-HK Drop Ship"
End If

ActiveWorkbook.SaveAs filename:=savepath & fname
' todo: for SOS, just save it;
'TODO: automatically add PO number on OSA sheet

' find specialsheetname
specialworkbook = ""
For Each book In Workbooks
    If book.Worksheets(1).Name = "get_all_orders" Then
        specialworkbook = book.Name
    End If
Next book

' convert to pdf
' order is already opened in single mode
folder = Application.ActiveWorkbook.Path & "\"
fullname = Application.ActiveWorkbook.Name

Call convert_to_pdf_function(specialworkbook, folder, fullname)

Call send_order_function(specialworkbook, folder, fullname)

End Sub
Sub get_all_orders()
Dim pub As String, folder As String, country As String, freight As String, tld As Range
Dim file As String, metafile As String, newbook As Workbook

'todo: add \ if not there
folder = Cells(1, 1)

Set newbook = Workbooks.Add
specialsheetname = "get_all_orders"

book1wbname = ActiveWorkbook.Name
ActiveSheet.Name = specialsheetname
Cells(1, 1) = folder

file = Dir(folder & "*.xlsx")
Count = 2
While (file <> "")
    'MsgBox file
    testcheck = file Like "[0-9]*.xlsx"
    pub = Left(file, 3)
    If testcheck Then
        MsgBox "This is the meta file"
    ElseIf pub = "SOS" Then
        MsgBox "SOS: You should use POD template"
    'ElseIf pub = "CUR" Or pub = "OBR" Or pub = "NHN" Then
    '    MsgBox "Put " & file & " in Brooke_to_send"
    Else
        Cells(Count, 1) = file
        Count = Count + 1
    End If
    file = Dir
Wend

' copy the ABK, Publication information from template
Workbooks.Open filename:="J:\NewSouth Books\Inventory\Templates\Purchase Order Template.xltx"
templatewbname = ActiveWorkbook.Name
Range("H7:K106").Select
Selection.copy

Workbooks(book1wbname).Activate
Range("C2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

Workbooks(templatewbname).Activate
ActiveWorkbook.Close savechanges:=False

' copy the EMAILS sheet from order file
Workbooks.Open filename:="J:\NewSouth Books\Inventory\Keyed Orders\Order List.xlsx"
orderlistwbname = ActiveWorkbook.Name
Sheets("EMAILS").copy after:=Workbooks(book1wbname).Sheets(1)

Workbooks(orderlistwbname).Activate
ActiveWorkbook.Sheets("EMAILS").Activate
ActiveWorkbook.Close savechanges:=False

Workbooks(book1wbname).Activate
ActiveWorkbook.Sheets(specialsheetname).Activate


End Sub
Sub send_order_function(specialworkbookname As String, folder As String, fullname As String)
Dim pub As String, country As String, freight As String, ABK As String, tld As Range, orderfmt As Range
Dim fname As String, wrdArray() As String, emailto As String, copy As String, copy2 As String, copy3 As String, copy4 As String, copy5 As String
Dim greetstr As String

'TODO: 1. handle  Perseus air and sea order;
'todo: 2. if current active is get_all_orders, then do as usual; if current is an order, then send seperate email


'folder = "J:\NewSouth Books\Inventory\Keyed Orders\221214\"
'folder = "D:\Users\z3507168\Desktop\test\"
'folder = Cells(1, 1)
' some constant
local_addr = "NewSouth Books c/- TL Distribution Pty Ltd"
US_ADDR = "Aeronet"
UK_SEA_ADDR = "JAS Forwarding UK Limited"
UK_AIR_ADDR = "JAS Forwarding UK Limited"
AIR_FREIGHT = "Air Freight"
SEA_FREIGHT = "SEA FREIGHT"
LOCAL_FREIGHT = "Local"
SING_ADDR = "C.T.FREIGHT PTE LTD"

orderwbname = ActiveWorkbook.Name
book1wbname = specialworkbookname
'Workbooks(book1wbname).Activate

'' open order file
'fullname = ActiveCell
'wrdArray = Split(fullname, ".")
'fname = wrdArray(0)
'fullpathname = folder & fullname
'xlsfullpathname = fullpathname
'pdffullpathname = folder & fname & ".pdf"
'Workbooks.Open FileName:=folder & fullname
'freight = Cells(8, 1)
'Address = Cells(11, 1)
'ABK = Cells(6, 1)
'ActiveWorkbook.Close


wrdArray = Split(fullname, ".")
fname = wrdArray(0)
xlsfullpathname = folder & fullname
pdffullpathname = folder & fname & ".pdf"
freight = Cells(8, 1)
Address = Cells(11, 1)
ABK = Cells(6, 1)


Workbooks(book1wbname).Activate
Set tld = Range("C1:C200").Find(ABK, SearchOrder:=xlByRows)
pub = Cells(tld.Row, tld.Column + 1)

If pub = "CRW" Then
    pub = "CRW UK"
    MsgBox "Use CRW UK by default."
End If

'todo check everywhere this affects
If pub = "CS1" Then
    pub = "CSO"
End If

' search in the order list and get pdf, email information
ActiveWorkbook.Sheets("EMAILS").Activate
Set orderfmt = Range("A1: A200").Find(pub, SearchOrder:=xlByRows)
pdf_or_xls = Cells(orderfmt.Row, orderfmt.Column + 1)

emailto = Cells(orderfmt.Row, orderfmt.Column + 3)
copy = Cells(orderfmt.Row, orderfmt.Column + 4)
copy2 = Cells(orderfmt.Row, orderfmt.Column + 5)
copy3 = Cells(orderfmt.Row, orderfmt.Column + 6)
copy4 = Cells(orderfmt.Row, orderfmt.Column + 7)
copy5 = Cells(orderfmt.Row, orderfmt.Column + 8)
Workbooks(book1wbname).Sheets("get_all_orders").Activate
If isEmail(emailto) = False Then
    MsgBox "Email is not valid: " & Email
    Exit Sub
End If
If isEmail(copy) = False Then
    copy = ""
End If
If isEmail(copy2) = False Then
    copy2 = ""
End If
If isEmail(copy3) = False Then
    copy3 = ""
End If
If isEmail(copy4) = False Then
    copy4 = ""
End If
If isEmail(copy5) = False Then
    copy5 = ""
End If

CC = "i.chavez@newsouthbooks.com.au;" & "brooke@newsouthbooks.com.au;" & _
    copy & ";" & copy2 & ";" & copy3 & ";" & copy4 & ";" & copy5 & ";"

' draft email
Dim objOutlook As Object
Dim objMailMessage As Outlook.MailItem
Dim emlBody, sendTo As String
Dim wkbook As String
Application.ScreenUpdating = False
Set objOutlook = CreateObject("Outlook.Application")
Set objMailMessage = objOutlook.CreateItem(0)
Subject = "Order " & fname

If Address = US_ADDR Then
    fdt_to = "Aeronet"
ElseIf Address = UK_SEA_ADDR Then
    fdt_to = "JAS Forwarding"
ElseIf Address = UK_AIR_ADDR Then
    fdt_to = "JAS Forwarding"
ElseIf Address = local_addr Then
    fdt_to = "TL Distribution."
ElseIf Address = SING_ADDR Then
    fdt_to = "C.T. Freight Singapore."
End If

If Address = SING_ADDR Then
    fdt_freight = ""
Else
    If freight = AIR_FREIGHT Then
        fdt_freight = " for air freight."
    ElseIf freight = SEA_FREIGHT Then
        fdt_freight = " for sea freight."
    Else
        fdt_freight = ""
    End If
End If

greetstr = "<p>Hello,</p>"
If pub = "CSO" Then
    greetstr = "<p>Hi Dalene,</p>"
ElseIf pub = "BER" Then
    greetstr = "<p>Hi Raudah,</p>"
End If

strbody = greetstr & _
            "<p>Please find attached " & Subject & " for delivery to " & fdt_to & fdt_freight & _
            "</p>"

If pub = "SOI" Then
    strbody = strbody & _
        "<p>Please responde to nearest quantity.</p>"
ElseIf pub = "OBR" Then
    strbody = strbody & "<p>Can you please let me know if anything on this order cannot be supplied in full?</p>"
ElseIf pub = "NHN" Then
    strbody = strbody & "<p>Please advise of any out of stocks, status code or pub date changes.</p>"
ElseIf pub = "TPG" Then
    Subject = Subject & " UK Delivery"
End If

notestr1 = "We have decided to consolidate our freight arrangements to JAS Forwarding UK, meaning all air and sea orders are now delivered to the same address. Please update your system so that all air freight orders are sent to the following address:" & _
"<p>JAS Forwarding UK Limited" & "<br>Cargopoint Bedfont Road" & "<br>Stanwell" & "<br>Middlesex" & "<br>TW19  7NZU UK" & _
"<br>Contact: Scott Caldwell" & "<br>scaldwell@ jasuk.com" & "<br>+44 01784 229 004</p>" & "<p>If you have any questions, please let me know.</p>"

If Address = UK_AIR_ADDR And freight = AIR_FREIGHT Then
    strbody = strbody & notestr1
End If

Dim signature As String
'signature = objMailMessage.HTMLBody
signature = Environ("appdata") & _
                "\Microsoft\Signatures\newsouthbooks.htm"
                
If pdf_or_xls = "XLS" Then
    attachfile = xlsfullpathname
ElseIf pdf_or_xls = "PDF" Then
    attachfile = pdffullpathname
End If

If Dir(signature) <> "" Then
    signature = GetBoiler(signature)
Else
    signature = ""
End If

With objMailMessage
    .To = emailto
    .CC = CC
    .HTMLBody = strbody & "<br>" & signature
    .Subject = Subject
    .Attachments.Add attachfile
    .Display
    '.Save
End With


Workbooks(orderwbname).Activate
'workbooks(s
'With Selection.Interior
'    .Pattern = xlSolid
'    .PatternColorIndex = xlAutomatic
'    'yellow : 65535, green: 5296274
'    .Color = 65535
'    .TintAndShade = 0
'    .PatternTintAndShade = 0
'End With
'Cells(ActiveCell.Row + 1, ActiveCell.Column).Select


End Sub
Sub send_order()
Dim pub As String, folder As String, country As String, freight As String, ABK As String, tld As Range, orderfmt As Range
Dim fname As String, fullname As String, wrdArray() As String, emailto As String, copy As String, copy2 As String, copy3 As String, copy4 As String, copy5 As String
Dim greetstr As String

'TODO: 1. handle  Perseus air and sea order;
'todo: 2. if current active is get_all_orders, then do as usual; if current is an order, then send seperate email


'folder = "J:\NewSouth Books\Inventory\Keyed Orders\221214\"
folder = "D:\Users\z3507168\Desktop\test\"
folder = Cells(1, 1)
' some constant
local_addr = "NewSouth Books c/- TL Distribution Pty Ltd"
US_ADDR = "Aeronet"
UK_SEA_ADDR = "JAS Forwarding UK Limited"
UK_AIR_ADDR = "JAS Forwarding UK Limited"
AIR_FREIGHT = "Air Freight"
SEA_FREIGHT = "SEA FREIGHT"
LOCAL_FREIGHT = "Local"
SING_ADDR = "C.T.FREIGHT PTE LTD"
book1wbname = ActiveWorkbook.Name
Workbooks(book1wbname).Activate

' open order file
fullname = ActiveCell
wrdArray = Split(fullname, ".")
fname = wrdArray(0)
fullpathname = folder & fullname
xlsfullpathname = fullpathname
pdffullpathname = folder & fname & ".pdf"
Workbooks.Open filename:=folder & fullname
freight = Cells(8, 1)
Address = Cells(11, 1)
ABK = Cells(6, 1)
ActiveWorkbook.Close

Workbooks(book1wbname).Activate
Set tld = Range("C1:C200").Find(ABK, SearchOrder:=xlByRows)
pub = Cells(tld.Row, tld.Column + 1)

If pub = "CRW" Then
    pub = "CRW UK"
    MsgBox "Use CRW UK by default."
End If

'todo check everywhere this affects
If pub = "CS1" Then
    pub = "CSO"
End If

' search in the order list and get pdf, email information
ActiveWorkbook.Sheets("EMAILS").Activate
Set orderfmt = Range("A1: A200").Find(pub, SearchOrder:=xlByRows)
pdf_or_xls = Cells(orderfmt.Row, orderfmt.Column + 1)

emailto = Cells(orderfmt.Row, orderfmt.Column + 3)
copy = Cells(orderfmt.Row, orderfmt.Column + 4)
copy2 = Cells(orderfmt.Row, orderfmt.Column + 5)
copy3 = Cells(orderfmt.Row, orderfmt.Column + 6)
copy4 = Cells(orderfmt.Row, orderfmt.Column + 7)
copy5 = Cells(orderfmt.Row, orderfmt.Column + 8)
If isEmail(emailto) = False Then
    MsgBox "Email is not valid: " & Email
    Exit Sub
End If
If isEmail(copy) = False Then
    copy = ""
End If
If isEmail(copy2) = False Then
    copy2 = ""
End If
If isEmail(copy3) = False Then
    copy3 = ""
End If
If isEmail(copy4) = False Then
    copy4 = ""
End If
If isEmail(copy5) = False Then
    copy5 = ""
End If

CC = "i.chavez@newsouthbooks.com.au;" & "brooke@newsouthbooks.com.au;" & _
    copy & ";" & copy2 & ";" & copy3 & ";" & copy4 & ";" & copy5 & ";"

' draft email
Dim objOutlook As Object
Dim objMailMessage As Outlook.MailItem
Dim emlBody, sendTo As String
Dim wkbook As String
Application.ScreenUpdating = False
Set objOutlook = CreateObject("Outlook.Application")
Set objMailMessage = objOutlook.CreateItem(0)
Subject = "Order " & fname

If Address = US_ADDR Then
    fdt_to = "Aeronet"
ElseIf Address = UK_SEA_ADDR Then
    fdt_to = "JAS Forwarding"
ElseIf Address = UK_AIR_ADDR Then
    fdt_to = "JAS Forwarding"
ElseIf Address = local_addr Then
    fdt_to = "TL Distribution."
ElseIf Address = SING_ADDR Then
    fdt_to = "C.T. Freight Singapore."
End If

If Address = SING_ADDR Then
    fdt_freight = ""
Else
    If freight = AIR_FREIGHT Then
        fdt_freight = " for air freight."
    ElseIf freight = SEA_FREIGHT Then
        fdt_freight = " for sea freight."
    Else
        fdt_freight = ""
    End If
End If

greetstr = "<p>Hello,</p>"
If pub = "CSO" Then
    greetstr = "<p>Hi Dalene,</p>"
ElseIf pub = "BER" Then
    greetstr = "<p>Hi Raudah,</p>"
End If

strbody = greetstr & _
            "<p>Please find attached " & Subject & " for delivery to " & fdt_to & fdt_freight & _
            "</p>"

If pub = "SOI" Then
    strbody = strbody & _
        "<p>Please responde to nearest quantity.</p>"
ElseIf pub = "OBR" Then
    strbody = strbody & "<p>Can you please let me know if anything on this order cannot be supplied in full?</p>"
ElseIf pub = "NHN" Then
    strbody = strbody & "<p>Please advise of any out of stocks, status code or pub date changes.</p>"
ElseIf pub = "TPG" Then
    Subject = Subject & " UK Delivery"
End If

notestr1 = "We have decided to consolidate our freight arrangements to JAS Forwarding UK, meaning all air and sea orders are now delivered to the same address. Please update your system so that all air freight orders are sent to the following address:" & _
"<p>JAS Forwarding UK Limited" & "<br>Cargopoint Bedfont Road" & "<br>Stanwell" & "<br>Middlesex" & "<br>TW19  7NZU UK" & _
"<br>Contact: Scott Caldwell" & "<br>scaldwell@ jasuk.com" & "<br>+44 01784 229 004</p>" & "<p>If you have any questions, please let me know.</p>"

If Address = UK_AIR_ADDR And freight = AIR_FREIGHT Then
    strbody = strbody & notestr1
End If

Dim signature As String
'signature = objMailMessage.HTMLBody
signature = Environ("appdata") & _
                "\Microsoft\Signatures\newsouthbooks.htm"
                
If pdf_or_xls = "XLS" Then
    attachfile = xlsfullpathname
ElseIf pdf_or_xls = "PDF" Then
    attachfile = pdffullpathname
End If

If Dir(signature) <> "" Then
    signature = GetBoiler(signature)
Else
    signature = ""
End If

With objMailMessage
    .To = emailto
    .CC = CC
    .HTMLBody = strbody & "<br>" & signature
    .Subject = Subject
    .Attachments.Add attachfile
    .Display
    '.Save
End With

Workbooks(book1wbname).Sheets("get_all_orders").Activate
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    'yellow : 65535, green: 5296274
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Cells(ActiveCell.Row + 1, ActiveCell.Column).Select


End Sub
Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function
Sub convert_to_pdf()
Dim pub As String, folder As String, country As String, freight As String, ABK As String, tld As Range, orderfmt As Range

Dim fname As String, fullname As String, wrdArray() As String

'folder = "J:\NewSouth Books\Inventory\Keyed Orders\221214\"
folder = Cells(1, 1)
book1wbname = ActiveWorkbook.Name
Workbooks(book1wbname).Activate

' open order file
fullname = ActiveCell

wrdArray = Split(fullname, ".")
fname = wrdArray(0)
fullpathname = folder & fullname
Workbooks.Open filename:=folder & fullname
orderwbname = ActiveWorkbook.Name
ABK = Cells(6, 1)

' find the TLD
Workbooks(book1wbname).Activate
Set tld = Range("C1:C200").Find(ABK, SearchOrder:=xlByRows)
pub = Cells(tld.Row, tld.Column + 1)

' todo: switch between UK/HK according to freight type
If pub = "CRW" Then
    pub = "CRW UK"
    MsgBox "Use CRW UK by default."
End If

'todo check everywhere this affects
If pub = "CS1" Then
    pub = "CSO"
End If

' search in the order list and get pdf, email information
ActiveWorkbook.Sheets("EMAILS").Activate

Set orderfmt = Range("A1: A200").Find(pub, SearchOrder:=xlByRows)
pdf_or_xls = Cells(orderfmt.Row, orderfmt.Column + 1)

Workbooks(orderwbname).Activate

If pdf_or_xls = "PDF" Then
    Dim ws As Worksheet
    Dim strPath As String
    Dim myFile As Variant
    Dim strFile As String
    On Error GoTo errHandler
    
    Set ws = ActiveSheet
    
    strFile = folder & fname & ".pdf"
    
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    If myFile <> "False" Then
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    
        'MsgBox "PDF file has been created."
    End If
    
errHandler:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & " just occured."
        MsgBox "Could not create PDF file"
    End If
                
ElseIf pdf_or_xls = "XLS" Then
    MsgBox "no need."
Else
    MsgBox "unkown format: " & pdf_or_xls & "!!!!"
End If

ActiveWorkbook.Close

Workbooks(book1wbname).Activate
ActiveWorkbook.Sheets("get_all_orders").Activate
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    'yellow : .Color = 65535
    .Color = 5296274
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Cells(ActiveCell.Row + 1, ActiveCell.Column).Select

End Sub
Sub convert_to_pdf_function(specialworkbookname As String, folder As String, fullname As String)
Dim pub As String, country As String, freight As String, ABK As String, tld As Range, orderfmt As Range

Dim fname As String, wrdArray() As String

orderwbname = ActiveWorkbook.Name
book1wbname = specialworkbookname
' suppose upon enter this function, the order workbook is already opened and activated
' open order file
'fullname = ActiveCell


wrdArray = Split(fullname, ".")
fname = wrdArray(0)
'fullpathname = folder & fullname
'Workbooks.Open FileName:=folder & fullname
'orderwbname = ActiveWorkbook.Name
'Workbooks(fullname).Activate

ABK = Cells(6, 1)

If ABK = "" Then
    MsgBox "ABK is empty. Check"
    Exit Sub
End If

' find the TLD
Workbooks(book1wbname).Activate
Set tld = Range("C1:C200").Find(ABK, SearchOrder:=xlByRows)
pub = Cells(tld.Row, tld.Column + 1)

If pub = "CRW" Then
    pub = "CRW UK"
    MsgBox "Use CRW UK by default."
End If

'todo check everywhere this affects
If pub = "CS1" Then
    pub = "CSO"
End If

' search in the order list and get pdf, email information
ActiveWorkbook.Sheets("EMAILS").Activate

Set orderfmt = Range("A1: A200").Find(pub, SearchOrder:=xlByRows)
pdf_or_xls = Cells(orderfmt.Row, orderfmt.Column + 1)
ActiveWorkbook.Sheets("get_all_orders").Activate

Workbooks(orderwbname).Activate

If pdf_or_xls = "PDF" Then
    Dim ws As Worksheet
    Dim strPath As String
    Dim myFile As Variant
    Dim strFile As String
    On Error GoTo errHandler
    
    Set ws = ActiveSheet
    
    strFile = folder & fname & ".pdf"
    
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    If myFile <> "False" Then
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    
        'MsgBox "PDF file has been created."
    End If
    
errHandler:
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & " just occured."
        MsgBox "Could not create PDF file"
    End If
                
ElseIf pdf_or_xls = "XLS" Then
    MsgBox "no need."
Else
    MsgBox "unkown format: " & pdf_or_xls & "!!!!"
End If


'ActiveWorkbook.Close

Workbooks(orderwbname).Activate
'ActiveWorkbook.Sheets("Sheet1").Activate
'With Selection.Interior
'    .Pattern = xlSolid
'    .PatternColorIndex = xlAutomatic
'    'yellow : .Color = 65535
'    .Color = 5296274
'    .TintAndShade = 0
'    .PatternTintAndShade = 0
'End With
'Cells(ActiveCell.Row + 1, ActiveCell.Column).Select

End Sub
Sub FixDateFromMac2PC()
'fix date issue when copying from mac to pc

    col = ActiveCell.Column
    Row = ActiveCell.Row
    Cells(23, 10) = 1462
    
    Range("J23").Select
    Selection.copy
    Cells(Row, col).Select
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "m/d/yyyy"
    Range("J23").Select
    Selection.ClearContents
    Cells(Row + 1, col).Select
    
End Sub

Sub test_multi_selection()
Dim fname As String, fullname As String, wrdArray() As String

Dim rng As Range
Set rng = Application.Selection
Dim cell As Range
For Each cell In rng
    MsgBox cell
Next

End Sub
Function isEmail_old(str As String) As Boolean
trimstr = Trim(WorksheetFunction.Substitute(str, Chr(160), " "))


isEmail = trimstr Like "[0-9a-zA-Z._]*@[0-9a-zA-Z-_]*.[0-9a-zA-Z-_.]*"

End Function

Function isEmail(str As String) As Boolean
trimstr = Trim(WorksheetFunction.Substitute(str, Chr(160), " "))
'todo: test thoroughly about this new pattern

'Refer to http://www.geeksengine.com/article/validate-email-vba.html
Set objRegExp_1 = CreateObject("vbscript.regexp")
objRegExp_1.Global = True
objRegExp_1.IgnoreCase = True
objRegExp_1.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
strToSearch = trimstr
Set regExp_Matches = objRegExp_1.Execute(strToSearch)
If regExp_Matches.Count = 1 Then
    isEmail = True
Else
    isEmail = False
End If

End Function


Sub Macro1984()
'
' Macro1984 Macro
'
    col = ActiveCell.Column
    Row = ActiveCell.Row
    Cells(23, 10) = 1462
    
    Range("J23").Select
    Selection.copy
    Cells(Row, col).Select
    
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "m/d/yyyy"
    Range("J23").Select
    Selection.ClearContents
    Cells(Row + 1, col).Select
    
End Sub

Sub ClearISBN()

Selection.ClearFormats
Selection.NumberFormat = "0"
Dim rng As Range
Set rng = Application.Selection
Dim cell As Range
For Each cell In rng
    cell = Trim(WorksheetFunction.Substitute(cell, Chr(160), " "))
Next
End Sub

Sub SamsDSR425WRK()
'
' SamsDSR425WRK

Call DSR425WRK

' Remove leading A000
    Cells.Replace What:="A000", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
' insert new PO column, concatenate PO#
    lastrow = Range("B:B").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Columns("B:B").Select

    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "PO#"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[6],RC[1])"
    Selection.AutoFill Destination:=Range(Cells(2, 2), Cells(lastrow, 2))
    
' reorder ISBN
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
' Hide Order
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
' reorder ship method
    Columns("E:E").Select
    Selection.Cut
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
' Hide L1, PUB, IMP
    Columns("F:H").Select
    Selection.EntireColumn.Hidden = True
' reorder STATUS
    Columns("I:I").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("N:N").Select
    Selection.Insert Shift:=xlToRight
' reorder DUE DATE
    Columns("J:J").Select
    Selection.Cut
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
' add my commet
Range("O1").Select
ActiveCell.FormulaR1C1 = "Commet"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
Range("O2").Select
ActiveCell.FormulaR1C1 = "IF(ISNA(VLOOKUP(C2,working.xlsx!working,9,FALSE)),"""",IF(VLOOKUP(C2,working.xlsx!working,9,FALSE)="""","""",VLOOKUP(C2,working.xlsx!working,9,FALSE)))"
'Selection.AutoFill Destination:=Range(Cells(2, 11), Cells(lastrow, 11))
End Sub


Sub test()
Dim rng3 As Range, isSeaOrder As Boolean, isLocalOrAir As Boolean

isSeaOrder = True
isLocalOrAir = True

Range([A2], Cells(Rows.Count, "A")).SpecialCells(xlCellTypeVisible)(1).Select

firstrow = ActiveCell.Row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Set Ret = Range("A:Z").Find("ORDER", SearchOrder:=xlByRows, LookAt:=xlWhole)

If firstrow = lastrow Then
    Set rng3 = Union(Cells(11519, 2), Cells(11519, 2))
    'MsgBox rng3
Else
'Set rng3 = Range(Cells(firstrow, Ret.Column), Cells(lastrow, Ret.Column))
'Set rng3 = Range(Cells(11519, 1), Cells(11519, 1))
End If

Dim c As Range
For Each c In rng3.SpecialCells(xlCellTypeVisible)
    If c = "" Then
        'MsgBox "This is not LocalOrAir Order"
        isLocalOrAir = False
    End If
Next
    
End Sub


Function IsWorkBookOpen(filename As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open filename For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    'Case Else: Error ErrNo
    End Select
End Function
