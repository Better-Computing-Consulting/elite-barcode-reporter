Imports System.IO
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Module Module1
    Sub Main()
        Dim manyfolders As New List(Of EliteFolder)()
        Dim filepath As String = ""
        If My.Application.CommandLineArgs.Count > 0 Then
            filepath = My.Application.CommandLineArgs(0)
        Else
            Console.WriteLine("Drag and drop onto the utility a .txt file containing only the list of barcodes.")
            Console.ReadLine()
            Exit Sub
        End If
        'Dim filepath As String = "C:\Users\fcanton\Documents\Outlook Files\Real Estate Inventory 8-15-18.txt"
        Try
            Using sr As StreamReader = New StreamReader(filepath)

                Dim aBarcode As String = ""
                Do
                    aBarcode = sr.ReadLine
                    If aBarcode <> "" Then
                        manyfolders.Add(New EliteFolder(aBarcode))
                    End If
                Loop Until aBarcode Is Nothing
                sr.Close()
            End Using
        Catch ex As Exception
            Console.WriteLine("The file could not be read:")
            Console.WriteLine(ex.Message)
        End Try
        Dim reportpath As String = CreateExcelReport(manyfolders)
        If Not String.IsNullOrEmpty(reportpath) Then Process.Start(reportpath)
        'Console.ReadLine()
    End Sub
    Private Function CreateExcelReport(BarCodeList As List(Of EliteFolder)) As String
        Dim aExcelApp As New Excel.Application
        Dim aExcelWrkbook As Excel.Workbook = aExcelApp.Workbooks.Add
        Dim xlsWSheet As Excel.Worksheet = aExcelWrkbook.Worksheets(1)
        Dim c As Integer = 1
        With xlsWSheet
            .Name = "FolderFromBarcodes"
            .Range("A1").Value = "findex"
            .Range("B1").Value = "fname"
            .Range("C1").Value = "fdesc1"
            .Range("C1").ColumnWidth = 40
            .Range("D1").Value = "mclient"
            .Range("E1").Value = "bindex"
            .Range("F1").Value = "bdesc1"
            .Range("F1").ColumnWidth = 35
            .Range("G1").Value = "bdnarr"
            .Range("G1").ColumnWidth = 55
            .Range("H1").Value = "mmatter"
            .Range("H1").ColumnWidth = 10
            .Range("I1").Value = "fbarcode"
            .Range("I1").ColumnWidth = 10
            .Range("J1").Value = "fdesc2"
            .Range("K1").Value = "ftmkpr"
            .Range("L1").Value = "fopen"
            .Range("L1").ColumnWidth = 15
            .Range("M1").Value = "fstatus"
            .Range("N1").Value = "fcrop"
            .Range("O1").Value = "freview"
            .Range("O1").ColumnWidth = 15
            .Range("P1").Value = "fstore"
            .Range("P1").ColumnWidth = 15
            .Range("Q1").Value = "fdestroy"
            .Range("Q1").ColumnWidth = 15
            .Range("R1").Value = "flocation"
            .Range("S1").Value = "fbox"
            .Range("T1").Value = "fallow"
            .Range("U1").Value = "finout"
            .Range("V1").Value = "ftype"
            .Range("W1").Value = "fclose"
            .Range("X1").Value = "fcrtime"
            .Range("Y1").Value = "fvolume"
            .Range("Z1").Value = "faddlloc"
            .Range("AA1").Value = "ffromdate"
            .Range("AA1").ColumnWidth = 15
            .Range("AB1").Value = "fthrudate"
            .Range("AC1").Value = "fmatthk"
            .Range("AD1").Value = "fdesc3"
            .Range("AE1").Value = "mediatype"
            .Range("AF1").Value = "ftkauth"
            .Range("AG1").Value = "fvital"
            .Range("AH1").Value = "factdestroy"
            .Range("AI1").Value = "fdestreason"
            .Range("AH1").Value = "fdocumentcount"
            .Range("AI1").Value = "finsertcount"
            .Range("AJ1").Value = "fcurrloc"
            .Range("AK1").Value = "fpleadings"
            .Range("AL1").Value = "fmatter"
            .Range("AM1").Value = "fsubnumber"
            .Range("AN1").Value = "clname1"
            .Range("AN1").ColumnWidth = 15
            For Each i As EliteFolder In BarCodeList
                c += 1
                .Range("A" & c).Value = i.findex
                .Range("B" & c).Value = i.fname
                .Range("C" & c).Value = i.fdesc1
                .Range("D" & c).Value = i.mclient
                '.Range("E" & c).Value = i.bindex
                '.Range("F" & c).Value = i.bdesc1
                'If Not String.IsNullOrEmpty(i.bdnarr.Trim) Then
                '    .Range("G" & c).Value = i.bdnarr.Remove(i.bdnarr.Length - 1)
                'End If
                .Range("H" & c).Value = i.mmatter
                .Range("I" & c).Value = i.fbarcode
                .Range("J" & c).Value = i.fdesc2
                .Range("K" & c).Value = i.ftmkpr
                .Range("L" & c).Value = i.fopen
                .Range("M" & c).Value = i.fstatus
                .Range("N" & c).Value = i.fcrop
                .Range("O" & c).Value = i.freview
                .Range("P" & c).Value = i.fstore
                .Range("Q" & c).Value = i.fdestroy
                .Range("R" & c).Value = i.flocation
                .Range("S" & c).Value = i.fbox
                .Range("T" & c).Value = i.fallow
                .Range("U" & c).Value = i.finout
                .Range("V" & c).Value = i.ftype
                .Range("W" & c).Value = i.fclose
                .Range("X" & c).Value = i.fcrtime
                .Range("Y" & c).Value = i.fvolume
                .Range("Z" & c).Value = i.faddlloc
                .Range("AA" & c).Value = i.ffromdate
                .Range("AB" & c).Value = i.fthrudate
                .Range("AC" & c).Value = i.fmatthk
                .Range("AD" & c).Value = i.fdesc3
                .Range("AE" & c).Value = i.mediatype
                .Range("AF" & c).Value = i.ftkauth
                .Range("AG" & c).Value = i.fvital
                .Range("AH" & c).Value = i.factdestroy
                .Range("AI" & c).Value = i.fdestreason
                .Range("AH" & c).Value = i.fdocumentcount
                .Range("AI" & c).Value = i.finsertcount
                .Range("AJ" & c).Value = i.fcurrloc
                .Range("AK" & c).Value = i.fpleadings
                .Range("AL" & c).Value = i.fmatter
                .Range("AM" & c).Value = i.fsubnumber
                .Range("AN" & c).Value = i.clname1
                .Rows(c & ":" & c).RowHeight = 45
            Next
        End With
        aExcelWrkbook.Sheets.Item("Sheet2").delete()
        aExcelWrkbook.Sheets.Item("Sheet3").delete()
        Dim filepath As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\FoldersBarcodesReport." & Now.ToString("yyyyMMddHHmmss") & ".xlsx"
        If File.Exists(filepath) Then
            Try
                File.Move(filepath, filepath.Replace(".xlsx", "." & Now.Ticks & ".xlsx"))
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Report save error")
                Return ""
            End Try
        End If
        aExcelWrkbook.SaveAs(filepath)
        aExcelWrkbook.Close()
        aExcelApp.Quit()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(aExcelWrkbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(aExcelApp)
        aExcelWrkbook = Nothing
        aExcelApp = Nothing
        GC.Collect()
        Return filepath
    End Function
End Module
Public Class EliteFolder
#Region "Class Variables"
    'Public bindex As String = ""
    'Public bdesc1 As String = ""
    Public findex As String = ""
    Public fname As String = ""
    Public fdesc1 As String = ""
    Public mclient As String = ""
    Public mmatter As String = ""
    Public fbarcode As String = ""
    Public fdesc2 As String = ""
    Public ftmkpr As String = ""
    Public fopen As String = ""
    Public fstatus As String = ""
    Public fcrop As String = ""
    Public freview As String = ""
    Public fstore As String = ""
    Public fdestroy As String = ""
    Public flocation As String = ""
    Public fbox As String = ""
    Public fallow As String = ""
    Public finout As String = ""
    Public ftype As String = ""
    Public fclose As String = ""
    Public fcrtime As String = ""
    Public fvolume As String = ""
    Public faddlloc As String = ""
    Public ffromdate As String = ""
    Public fthrudate As String = ""
    Public fmatthk As String = ""
    Public fdesc3 As String = ""
    Public mediatype As String = ""
    Public ftkauth As String = ""
    Public fvital As String = ""
    Public factdestroy As String = ""
    Public fdestreason As String = ""
    Public fdocumentcount As String = ""
    Public finsertcount As String = ""
    Public fcurrloc As String = ""
    Public fpleadings As String = ""
    Public fmatter As String = ""
    Public fsubnumber As String = ""
    Public clname1 As String = ""
#End Region
    Sub New(aBarCode As String)
        fbarcode = aBarCode
        Using conn As New SqlConnection("Data Source=wrselite;Initial Catalog=son_db;Integrated Security=SSPI")
            Dim queryString As String = "SELECT DISTINCT folder.findex,folder.fname,folder.fdesc1,matter.mclient,clname1,matter.mmatter,folder.fbarcode,folder.fdesc2,folder.ftmkpr,folder.fopen,folder.fstatus,folder.fcrop,folder.freview,folder.fstore,folder.fdestroy,folder.flocation,folder.fbox,folder.fallow,folder.finout,folder.ftype,folder.fclose,folder.fcrtime,folder.fvolume,folder.faddlloc,folder.ffromdate,folder.fthrudate,folder.fmatthk,folder.fdesc3,folder.mediatype,folder.ftkauth,folder.fvital,folder.factdestroy,folder.fdestreason,folder.fdocumentcount,folder.finsertcount,folder.fcurrloc,folder.fpleadings,folder.fmatter,folder.fsubnumber FROM folder,matter,client WHERE folder.fmatter = matter.mmatter AND clnum = mclient AND folder.fbarcode = '" & fbarcode & "' ORDER BY folder.fname,folder.findex"
            Dim cmd As New SqlCommand(queryString, conn)
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader()
            If r.HasRows Then
                Console.WriteLine(fbarcode & " found")
                Try
                    While r.Read
                        If Not IsDBNull(r.Item("findex")) Then findex = Trim(r.Item("findex"))
                        If Not IsDBNull(r.Item("fname")) Then fname = Trim(r.Item("fname"))
                        If Not IsDBNull(r.Item("fdesc1")) Then fdesc1 = Trim(r.Item("fdesc1"))
                        If Not IsDBNull(r.Item("mclient")) Then mclient = Trim(r.Item("mclient"))
                        If Not IsDBNull(r.Item("clname1")) Then clname1 = Trim(r.Item("clname1"))
                        If Not IsDBNull(r.Item("mmatter")) Then mmatter = Trim(r.Item("mmatter"))
                        If Not IsDBNull(r.Item("fbarcode")) Then fbarcode = Trim(r.Item("fbarcode"))
                        If Not IsDBNull(r.Item("fdesc2")) Then fdesc2 = Trim(r.Item("fdesc2"))
                        If Not IsDBNull(r.Item("ftmkpr")) Then ftmkpr = Trim(r.Item("ftmkpr"))
                        If Not IsDBNull(r.Item("fopen")) Then fopen = Trim(r.Item("fopen"))
                        If Not IsDBNull(r.Item("fstatus")) Then fstatus = Trim(r.Item("fstatus"))
                        If Not IsDBNull(r.Item("fcrop")) Then fcrop = Trim(r.Item("fcrop"))
                        If Not IsDBNull(r.Item("freview")) Then freview = Trim(r.Item("freview"))
                        If Not IsDBNull(r.Item("fstore")) Then fstore = Trim(r.Item("fstore"))
                        If Not IsDBNull(r.Item("fdestroy")) Then fdestroy = Trim(r.Item("fdestroy"))
                        If Not IsDBNull(r.Item("flocation")) Then flocation = Trim(r.Item("flocation"))
                        If Not IsDBNull(r.Item("fbox")) Then fbox = Trim(r.Item("fbox"))
                        If Not IsDBNull(r.Item("fallow")) Then fallow = Trim(r.Item("fallow"))
                        If Not IsDBNull(r.Item("finout")) Then finout = Trim(r.Item("finout"))
                        If Not IsDBNull(r.Item("ftype")) Then ftype = Trim(r.Item("ftype"))
                        If Not IsDBNull(r.Item("fclose")) Then fclose = Trim(r.Item("fclose"))
                        If Not IsDBNull(r.Item("fcrtime")) Then fcrtime = Trim(r.Item("fcrtime"))
                        If Not IsDBNull(r.Item("fvolume")) Then fvolume = Trim(r.Item("fvolume"))
                        If Not IsDBNull(r.Item("faddlloc")) Then faddlloc = Trim(r.Item("faddlloc"))
                        If Not IsDBNull(r.Item("ffromdate")) Then ffromdate = Trim(r.Item("ffromdate"))
                        If Not IsDBNull(r.Item("fthrudate")) Then fthrudate = Trim(r.Item("fthrudate"))
                        If Not IsDBNull(r.Item("fmatthk")) Then fmatthk = Trim(r.Item("fmatthk"))
                        If Not IsDBNull(r.Item("fdesc3")) Then fdesc3 = Trim(r.Item("fdesc3"))
                        If Not IsDBNull(r.Item("mediatype")) Then mediatype = Trim(r.Item("mediatype"))
                        If Not IsDBNull(r.Item("ftkauth")) Then ftkauth = Trim(r.Item("ftkauth"))
                        If Not IsDBNull(r.Item("fvital")) Then fvital = Trim(r.Item("fvital"))
                        If Not IsDBNull(r.Item("factdestroy")) Then factdestroy = Trim(r.Item("factdestroy"))
                        If Not IsDBNull(r.Item("fdestreason")) Then fdestreason = Trim(r.Item("fdestreason"))
                        If Not IsDBNull(r.Item("fdocumentcount")) Then fdocumentcount = Trim(r.Item("fdocumentcount"))
                        If Not IsDBNull(r.Item("finsertcount")) Then finsertcount = Trim(r.Item("finsertcount"))
                        If Not IsDBNull(r.Item("fcurrloc")) Then fcurrloc = Trim(r.Item("fcurrloc"))
                        If Not IsDBNull(r.Item("fpleadings")) Then fpleadings = Trim(r.Item("fpleadings"))
                        If Not IsDBNull(r.Item("fmatter")) Then fmatter = Trim(r.Item("fmatter"))
                        If Not IsDBNull(r.Item("fsubnumber")) Then fsubnumber = Trim(r.Item("fsubnumber"))
                    End While
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SQL Error on all")
                End Try
            Else
                Console.WriteLine(fbarcode & " not found")
            End If
        End Using
    End Sub
End Class