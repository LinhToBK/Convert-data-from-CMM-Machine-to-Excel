VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   OleObjectBlob   =   "Linh RD Filter Rev_14_Jan_2021.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'=================btn_chaydi_Click() ==============
Private Sub btn_chaydi_Click()

'================================================
'==============        Nap file TXT   ===========
'================================================
    Dim fileToOpen As Variant
    Dim demA As Integer
    Dim i As Integer
    Dim listname(250)
    Dim filenow As String
    filenow = ThisWorkbook.Name

    

    fileToOpen = Application.GetOpenFilename(MultiSelect:=True)
    ' "fileToOpen : la 1 bien luu cac file txt vua nap
    demA = UBound(fileToOpen)
    
    Dim wbMaster As Worksheet
    Dim wbtextImport As Workbook
    
    ' ====== Vong For cho lap chay nhieu doi tuong ========
    For i = 1 To demA
        Workbooks.OpenText fileToOpen(i), DataType:=xlDelimited, Space:=True
        listname(i) = Sheets(1).Name
        ' thuc hien nap du lieu vao sheet
        Set wbtextImport = ActiveWorkbook
        Sheets(1).Copy Before:=Workbooks(filenow).Sheets(1)
        wbtextImport.Close False
    Next i

'==================================================
'=============Finish Nap file TXT =================
'==================================================

'==================================================
'=============Lay du lieu tu user form ============
'==================================================
    Dim cute As String
    Dim dem As Integer
    Dim dem2 As Integer
    dem = demA + 1
    Sheets(listname(1)).Select ' Chon de xoa chuc nang chon toan bo

    '==========TAO SHEET MOI DE COPY DU LIEU================
    Dim status_Doiten As Boolean
    Dim status_Doimau As Boolean
    Dim status_Tensheet As String
    Dim sheetmoi As Worksheet
'    status_Doiten = ckc_Cotten.Value
    status_Doimau = ckc_Doimau.Value
    status_Tensheet = txt_Tensheet.Value
    
'========= Check xem file da ton tai hay chua ======
    Dim dem4 As Integer
    For dem4 = 1 To dem - 1
        If status_Tensheet = listname(dem4) Then
            MsgBox ("ten sheet nay trung roi, kiem tra lai di !!!")
            Unload Me 'neu co sheet roi thi thoat chuong trinh
        End If
    Next dem4
'====================================================
    If status_Tensheet = "" Then
        status_Tensheet = "Linhcute"
    End If
'    MsgBox (dem)
    Set sheetmoi = Worksheets.Add
    sheetmoi.Name = status_Tensheet
'==================================================
'=============Finish nap du lieu tu user form =====
'==================================================
    
' ======== LOC DU LIEU DOI VOI SHEET =====
            Dim dem3 As Integer
            Dim so_cot As Integer
            Dim dong_cuoi As Integer
            Dim demPar As Integer

        For dem2 = 1 To (dem - 1) ' dem2 => la bien so luong file
            Sheets(listname(dem2)).Select ' Chon de xoa chuc nang chon toan bo
            dong_cuoi = Range("A" & Rows.Count).End(xlUp).Row
            'MsgBox ("Dong cuoi cua listname la : " & dong_cuoi)
            so_cot = 1
'            '***********************************
            ' Version 3 : check bang dong lenh DIM
            Dim Nominal_value As Double
            Dim Up_tolerance As Double
            Dim down_tolerance As Double
            '****************************
            For dem3 = 1 To dong_cuoi
                If Cells(dem3, 1) = "DIM" Then
                    'MsgBox ("Da tim thay dong Dim: " & dem3)
                    
                    ' ======================================
                    ' =======truong hop la {"POINT"}========
                    ' ======================================
                    If Cells(dem3, 5) = "POINT" Then
                        'MsgBox (" Day la tinh POINT ")
                        '___________________________________
                        ' __ xoa dong PAR __ neu co ______
                        For demPar = 1 To 4
                            If Cells(dem3 + demPar, 1) = "PART" Then
                                Rows(dem3 + demPar).EntireRow.Delete
                            End If
                        Next demPar ' end __ xoa dong PAR __ neu co
                        '___________________________________
                        ' __ check doi mau diem NG __
                        If status_Doimau = True Then
                        ' so sanh voi gia tri goc
                            Nominal_value = Cells(dem3 + 2, 2)
                            Up_tolerance = Cells(dem3 + 2, 3)
                            down_tolerance = Cells(dem3 + 2, 4)
                        
                            If (Cells(dem3 + 2, 5) > Nominal_value + Up_tolerance) Or (Cells(dem3 + 2, 5) < Nominal_value - down_tolerance) Then
                                Cells(dem3 + 2, 5).Interior.Color = RGB(250, 0, 0) ' color red
                            End If ' so sanh gia tri goc
                        End If '  end __ check doi mau diem NG __
                        '___________________________________
                       ' __ copy sang  sheet moi ___
                       
                       Range(Cells(dem3, 5), Cells(dem3 + 3, 5)).Copy Destination:=sheetmoi.Cells(2 + (dem2 - 1) * 5, so_cot + 1)
                       Cells(dem3, 2).Copy Destination:=sheetmoi.Cells(1 + (dem2 - 1) * 5, so_cot + 1)
                    so_cot = so_cot + 1
                    End If ' // end truong hop la {"POINT"}
                    
                    ' ======================================
                    ' =====truong hop la {"CIRCLE"} ========
                    ' ======================================
                    If Cells(dem3, 5) = "CIRCLE" Then
                        '*******************************
                        For demPar = 1 To 4
                            If Cells(dem3 + demPar, 1) = "PART" Then
                                Rows(dem3 + demPar).EntireRow.Delete
                            End If
                        Next demPar ' end delete cells("PAR")

                    '*******************************
                    If status_Doimau = True Then
                        ' so sanh voi gia tri goc
                        If Cells(dem3 + 4, 6) > Cells(dem3 + 4, 3) Then
                            Cells(dem3 + 4, 6).Interior.Color = RGB(250, 0, 0)

                        End If ' so sanh gia tri goc
                    End If ' doi mau thanh phan NG
                    '********************************
                    
                    Cells(dem3, 6).Copy Destination:=sheetmoi.Cells(1 + (dem2 - 1) * 5, so_cot + 1)
                    Range(Cells((dem3 + 1), 6), Cells(dem3 + 4, 6)).Copy Destination:=sheetmoi.Cells(2 + (dem2 - 1) * 5, so_cot + 1)
                    
                    so_cot = so_cot + 1

                    End If '// end truong hop la {"CIRCLE"}
                    
                End If ' // check cells ("DIM ) //
                
            Next dem3

        '    MsgBox ("So luong diem check la :" & (soluong - 1))
            sheetmoi.Cells(1 + (dem2 - 1) * 5, 1) = listname(dem2)
        Next dem2






'*****************************************************
'*****************************************************
'*****************************************************

'If status_Doiten = False Then
'    For dem2 = 1 To (dem - 1)
'            Sheets(listname(dem2)).Select ' Chon de xoa chuc nang chon toan bo
'            'Dim dong_cuoi2 As Integer
'            dong_cuoi = Range("A" & Rows.Count).End(xlUp).Row
'        '    MsgBox ("Dong cuoi cua listname la : " & dong_cuoi)
'
'        ' ===== Lap tim cac measure va gan vao cac hang o ben sheetmoi  ===
''            Dim dem3 As Integer
''            Dim soluong As Integer
'            soluong = 1
'
'            For dem3 = 1 To dong_cuoi
'                If Cells(dem3, 6) = "MEAS" Then
'                    '**********************************
'                    For demPar = 1 To 4
'                        If Cells(dem3 + demPar, 1) = "PART" Then
'                            Rows(dem3 + demPar).EntireRow.Delete
'                        End If
'                    Next demPar
'
'                    '*******************************
'                    If status_Doimau = True Then
'                        ' so sanh voi gia tri goc
'                        If Cells(dem3 + 4, 6) > Cells(dem3 + 4, 3) Then
'                            Cells(dem3 + 4, 6).Interior.Color = RGB(250, 0, 0)
'
'                        End If ' so sanh gia tri goc
'                    End If ' doi mau thanh phan NG
'                    '***********************************
'                    Range(Cells((dem3 + 1), 6), Cells(dem3 + 4, 6)).Copy Destination:=sheetmoi.Cells(2 + (dem2 - 1) * 5, soluong + 1)
'                    Cells((dem3 - 1), 4).Copy Destination:=sheetmoi.Cells(1 + (dem2 - 1) * 5, soluong + 1)
'                    soluong = soluong + 1
'                End If
'
'            Next dem3
'        '    MsgBox ("So luong diem check la :" & (soluong - 1))
'            sheetmoi.Cells(1 + (dem2 - 1) * 5, 1) = listname(dem2)
'
'        Next dem2 ' end for dem2
'
'
'End If ' == status_Doiten = true
'*************************************************************



MsgBox ("Chay xong roi nhe !!!")
Unload Me


End Sub
'================= finish sub btn_chaydi_Click ==============














