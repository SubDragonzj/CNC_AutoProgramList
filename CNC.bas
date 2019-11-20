Attribute VB_Name = "CNC"
Dim ncrow As Long

Sub foundT(txtfile, trow)

    Dim fs, f
    Dim sRow, D1, T1, S1 As String
    Dim Z1, Snum As Single
    Dim r, h, k, n, m As Long
    Dim ym06 As Long

    Z1 = 200
    h = 100000
    Snum = 1
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(txtfile)
    
    If f.readline <> "%" Then Exit Sub                                   '判定是否NC程式
    
    Do While Left(sRow, 1) <> "%"
        r = r + 1
        sRow = f.readline
        If ym06 = True And InStr(sRow, "M06") > 0 Then                      '判定是否換刀異常
            MsgBox "注意：程式有2處M06換刀，請更正程式！！！", vbExclamation
            End
        End If
        If D1 = "" Then
            If Left(sRow, 1) = "(" And Right(sRow, 1) = ")" Then             '抓取刀具
                k = InStr(sRow, "=") + 1
                For n = 2 To 8
                    If Mid(sRow, k, n) = Trim(Mid(sRow, k, n + 1)) Then
                        D1 = Mid(sRow, k, n)
                        Exit For
                    End If
                Next
            End If
        End If
        
        If T1 = "" Then
            If Left(sRow, 1) = "T" And Right(sRow, 3) = "M06" Then           '抓取刀號
                If Left(sRow, 3) = "T00" Then                                '判定刀號是否異常
                    MsgBox "注意：  此行 T00 刀號異常，請更正程式！！！", vbExclamation
                    End
                End If
                ym06 = True
                T1 = Trim(Left(sRow, Len(sRow) - 3))
                h = r + 1
            End If
        End If

            k = InStr(sRow, "S")                                          '抓取最大轉速
            If k > 0 Then
                For n = 1 To Len(sRow) - k
                    If n = Len(sRow) - k Then
                        If Snum < CSng(Mid(sRow, k + 1, n)) Then Snum = CSng(Mid(sRow, k + 1, n))
                            If Snum>30000 Or Snum<10 Then
                                MsgBox "注意： " & D1 & "轉速" & Snum & "轉異常，請更正程式！！！", vbExclamation
                                End
                            End if
                            S1 = CStr(Snum)
                    Else
                        If Mid(sRow, k + 1, n) = Trim(Mid(sRow, k + 1, n + 1)) Then
                            If Snum < CSng(Mid(sRow, k + 1, n)) Then
                                Snum = CSng(Mid(sRow, k + 1, n))
                                If Snum>30000 Or Snum<10 Then
                                    MsgBox "注意： " & D1 & "轉速" & Snum & "轉異常，請更正程式！！！", vbExclamation
                                    End
                                End if
                                S1 = CStr(Snum)
                            End If
                            Exit For
                        End If
                    End If
                Next
            End If
        
        If h < r And InStr(sRow, "G91") = 0 Then
            m = InStr(sRow, "F")
            If m > 1 Then                                                    '判斷F值是否正確
                For n = 1 To Len(sRow) - m
                    If n = Len(sRow) - m Then
                        If CLng(Mid(sRow, m + 1, n)) < 10 Or CLng(Mid(sRow, m + 1, n)) > 10000 Then
                            MsgBox "注意： " & T1 & "進給速度F" & CLng(Mid(sRow, m + 1, n)) & " 異常，請更正程式！！！", vbExclamation
                            End
                        End If
                    Else
                        If Mid(sRow, m, n) = Trim(Mid(sRow, m, n + 1)) Then
                            If CLng(Mid(sRow, m + 1, n)) < 10 Or CLng(Mid(sRow, m + 1, n)) > 10000 Then
                                MsgBox "注意： " & T1 & "進給速度F" & CLng(Mid(sRow, m + 1, n)) & " 異常，請更正程式！！！", vbExclamation
                                End
                            End If
                            Exit For
                        End If
                    End If
                Next
            End If
        
            k = InStr(sRow, "Z")                                          '抓取最大加工深度
            If k > 0 Then
                For n = 1 To Len(sRow) - k
                    If n = Len(sRow) - k Then
                        If Z1 > CSng(Mid(sRow, k + 1, n)) Then Z1 = CSng(Mid(sRow, k + 1, n))
                    Else
                        If Mid(sRow, k + 1, n) = Trim(Mid(sRow, k + 1, n + 1)) Then
                            If Z1 > CSng(Mid(sRow, k + 1, n)) Then
                                Z1 = CSng(Mid(sRow, k + 1, n))
                            End If
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    Loop

    f.Close
    Set f = Nothing
    Set fs = Nothing
    
    If D1 = "" Or T1 = "" Or S1 = "" Or Z1 = 200 Then
        MsgBox txtfile & "程式文件有誤！請確認！", vbExclamation
        End
    End If
    
    Cells(trow, 2).Value = D1
    Cells(trow, 3).Value = T1
    Cells(trow, 4).Value = S1
    Cells(trow, 5).Value = CStr(Round(Z1, 3))

End Sub

Sub main()
    Dim fs, f, f1, fc, ft
    Dim fd As FileDialog
    Dim vrtSelectedItem As Variant
    Dim fPath As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then
        Call Ncr                                                            '查找程序起始行
        Cells(4, 5).Value = Now()
        If Cells(3, 11).Value = "" Then
            MsgBox "加工件相關資料未填寫，請務必注意！", vbExclamation
        Else
            Call Rew
        End If
        Range("A" & ncrow & ":N100").ClearContents                          '刪除舊數據
        For Each vrtSelectedItem In fd.SelectedItems
            fPath = vrtSelectedItem
        Next vrtSelectedItem

        Set f = fs.GetFolder(fPath)
        Set fc = f.Files
        For Each f1 In fc
            If f1.Type = "文字文件" Then
                ft = fPath & "/" & f1.Name
                Call foundT(ft, ncrow)
                If Cells(ncrow, 2).Value <> "" Then
                    Cells(ncrow, 1).Value = Left(f1.Name, Len(f1.Name) - 4)
                    Cells(ncrow, 1).HorizontalAlignment = xlCenter
                    ncrow = ncrow + 1
                    Rows(ncrow & ":" & ncrow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                End If
            End If
        Next
        If Cells(ncrow - 1, 1).Value = "程式名" Then
            MsgBox "此資料夾無程式文檔，請確認！", vbExclamation
        Else
            Cells(5, 2).Value = f
            Rows(ncrow & ":" & ncrow + 2).Delete Shift:=xlUp
            Cells(ncrow, 1).HorizontalAlignment = xlLeft
            Cells(ncrow, 1).Value = "此程式單上資料僅供參考，零件尺寸公差及特殊要求等以圖紙為準，零件加工完成后請做好自檢。"
        End If
    End If

    Set f = Nothing
    Set fs = Nothing
    Set fd = Nothing
    'Shell "F:\Screen.exe", vbNormalNoFocus                             '打開截圖軟件
End Sub

Sub Rew()
    Cells(4, 2).Value = Cells(3, 9).Value
    Range("I3:J3").ClearContents
    
    Rows("3:3").Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Bold = False
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .WrapText = True
        .Rows.AutoFit
    End With
    
    Range("I3:J3").Merge
    Cells(3, 9).Value = Cells(3, 12).Value
    Range("K1:N50").Clear
End Sub

Sub Ncr()
    Dim i As Long
    ncrow = 0
    For i = 6 To 50
        If Trim(Cells(i, 1).Value) = "程式名" Then
            ncrow = i + 1
            Rows(ncrow + 2 & ":" & ncrow + 50).Delete Shift:=xlUp
            Exit Sub
        End If
    Next
    If ncrow = 0 Then
        MsgBox "程式單第一列未能找到 程式名，無法確認程式數據起始行！", vbExclamation
        End
    End If
End Sub

