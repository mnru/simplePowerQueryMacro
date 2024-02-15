Attribute VB_Name = "Module1"
Sub testQry()
    fml = tblToQry("èjì˙qry", Array(2021, 5))
    Call mkQueryTable(fml, "èjì˙ï\")
End Sub


Function tblToQry(tbln, Optional params, Optional ret = "")
    Dim qry, sps, first, p
    Dim rNum
    sps = String(4, " ")
    qry = "let" & vbCrLf
    first = True
    p = 0
    rNum = Range(tbln).Rows.Count
    For r = 1 To rNum
        expr = Range(tbln & "[éÆ]")(r)
        If expr = "?" Then
            If IsArray(params) Then
                If p <= UBound(params) Then
                    expr = params(p)
                    p = p + 1
                End If
            End If
        Else
            expr = Replace(expr, vbLf, vbCrLf & sps)
        End If
        qry = qry & sps & IIf(first, "", ",") & Range(tbln & "[ïœêî]")(r) & "=" & expr & vbCrLf
        first = False
    Next
    qry = qry & "in" & vbCrLf
    If ret = "" Then ret = Range(tbln & "[ïœêî]")(rNum)
    qry = qry & sps & ret & vbCrLf
    tblToQry = qry
End Function


Sub mkQueryTable(fml, tbln, Optional qryn = "", Optional shtn = "", Optional r = 1, Optional c = 1, Optional dbg = True, Optional deltmp = True)
    t0 = Time
    Dim formula As String
    If deltmp Then
        Call delTmpQrys
        Call delTmpSheets
    End If
    If dbg Then Debug.Print fml
    If qryn = "" Then qryn = tbln
    If shtn = "" Then shtn = ThisWorkbook.Sheets.Add.Name
    Set qry = ThisWorkbook.Queries.Add(Name:=qryn, formula:=fml)
    Set sht = ThisWorkbook.Sheets(shtn)
    Set lo = sht.ListObjects.Add( _
    SourceType:=xlSrcExternal, _
    Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
    "Location=" & qryn, _
    Destination:=sht.Cells(r, c))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qry.Name & "]")
        .ListObject.DisplayName = tbln
        .Refresh BackgroundQuery:=False
    End With
    t1 = Time
    Debug.Print Format(t1 - t0, "hh:mm:ss")
End Sub

Sub delTmpQrys()
    Dim qry, cn
    On Error Resume Next
    For Each qry In ThisWorkbook.Queries
        qry.Delete
    Next
    For Each cn In ThisWorkbook.Connections
        'Debug.Print cn.Name
        If cn.Name Like "ê⁄ë±*" Or cn.Name Like "WorkSheetConnection_*" Then cn.Delete
    Next
    On Error GoTo 0
End Sub


Sub delTmpSheets()
    On Error Resume Next
    Application.DisplayAlerts = False
    For Each sht In ThisWorkbook.Sheets
        If sht.Name Like "Sheet*" Then
            sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

