'导入命名空间，导入前需添加引用Microsoft excel 14.0 object library(excel 2010)
Imports Microsoft.Office.Interop.Excel
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swcommands
Imports SolidWorks.Interop.swconst
Imports SolidWorksTools
Imports SolidWorks.Interop.swpublished
Imports FolderBrowserDialog

Public Class 批量读取文件名

    '定义变量 

    '文件名
    Dim StrFilename As String
    '文件名数组
    Dim strFolderName As String()
    '已过滤文件名数组
    Dim Sfnaf(1) As String
    '数组循环开始标志
    Dim i As String = 0
    'Excel应用程序对象
    Dim vbexcel As Microsoft.Office.Interop.Excel.Application
    'Excel工作薄
    Dim vbworkbook As Workbook
    'Excel工作表
    Dim vbworksheet As Worksheet

    Dim FileNames As New List(Of String)
    Dim br1 As FolderBrowserDialog.FolderBrowserDialog = New FolderBrowserDialog.FolderBrowserDialog()
   



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "pdf"
    End Sub
    '定义一个过滤器函数，检查文件后缀名
    Public Function GetFileType(filename As String) As String
        Dim k
        Dim result As String
        k = Split(filename, ".")
        result = k(UBound(k))
        Return result
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        i = 0
        ReDim Sfnaf(i)

        ' br1.ShowDialog(Me)
        If br1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            '获取目标路径下所有文件路径，保存到数组StrFolderName
            strFolderName = System.IO.Directory.GetFiles(br1.DirectoryPath)
            Label3.Text = "当前目录：" & br1.DirectoryPath
            TextBox1.Text = ""
            For Each Me.StrFilename In strFolderName
                If GetFileType(StrFilename) = ComboBox1.Text Or GetFileType(StrFilename) = ComboBox1.Text.ToUpper() Then
                    'i = 0
                    ReDim Preserve Sfnaf(i)
                    Sfnaf(i) = System.IO.Path.GetFileNameWithoutExtension(StrFilename)

                    FileNames.Add(StrFilename)
                    TextBox1.Text = TextBox1.Text & System.IO.Path.GetFileNameWithoutExtension(StrFilename) & vbNewLine
                    i = i + 1
                    'Try
                    '    vbexcel = GetObject(, "excel.application")
                    '    vbexcel.Visible = True
                    'Catch ex As Exception
                    '    vbexcel = CreateObject("excel.application")
                    '    vbexcel.Visible = True
                    'End Try



                    'With vbexcel
                    '    If vbexcel.ActiveWorkbook Is Nothing Then
                    '        vbworkbook = .Workbooks.Add()

                    '    Else : vbworkbook = vbexcel.ActiveWorkbook
                    '    End If
                    '    With vbworkbook
                    '        If vbexcel.ActiveSheet Is Nothing Then
                    '            vbworksheet = .Worksheets.Add()
                    '            vbworksheet.Activate()
                    '        Else : vbworksheet = vbexcel.ActiveSheet
                    '        End If
                    '        With vbworksheet
                    '            .Cells(i, 1) = System.IO.Path.GetFileNameWithoutExtension(StrFilename)
                    '        End With
                    '    End With
                    'End With

                    'i = i + 1
                End If

            Next
        End If
        Label2.Text = "文件数量：" & (i)

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            vbexcel = GetObject(, "excel.application")

            vbexcel.Visible = True
        Catch ex As Exception
            vbexcel = CreateObject("excel.application")
            vbexcel.Visible = True
        End Try



        With vbexcel

            'If vbexcel.ActiveWorkbook IsNot Nothing Then
            '    vbworkbook = .Workbooks.Add()

            'Else : vbworkbook = vbexcel.ActiveWorkbook

            'End If
            'With vbworkbook
            '    If vbexcel.ActiveSheet Is Nothing Then
            '        'vbworksheet = .Worksheets.Add()
            '        vbworksheet = .Worksheets.Add()
            '        vbworksheet.Activate()
            '    Else : vbworksheet = vbexcel.ActiveSheet
            '    End If
            vbworkbook = .Workbooks.Add()
            With vbworkbook
                vbworksheet = .Worksheets.Add()
                vbworksheet.Activate()

                With vbworksheet
                    .Cells(2, 1) = "代号"
                    .Cells(2, 2) = "名称"
                    .Cells(2, 3) = "每台数量"
                    .Cells(2, 4) = "材质"
                    .Cells(2, 5) = "设备型号"
                    .Cells(2, 6) = "表面处理"
                    .Cells(2, 7) = "加工数量"
                    .Cells(2, 8) = "交货日期"
                    .Cells(2, 9) = "备注"
                    .Range("A1:I1").MergeCells = True
                    .Cells(1, 1) = New System.IO.DirectoryInfo(br1.DirectoryPath).Name
                    ' FolderBrowserDialog1.SelectedPath.g()
                    For j = 0 To Sfnaf.Length - 1
                        .Cells(j + 3, 1) = Sfnaf(j)
                        If ComboBox1.Text = "sldprt" Then
                            Dim customnames As ArrayList
                            customnames = New ArrayList()
                            customnames.Add("零件名称")
                            customnames.Add("数量")
                            customnames.Add("材料")


                            Dim res As New ArrayList()
                            res = GetCustomPropertys(FileNames.Item(j), customnames)
                            .Cells(j + 3, 2) = res.Item(0)
                            .Cells(j + 3, 3) = res.Item(1)
                            .Cells(j + 3, 4) = res.Item(2)
                        End If
                    Next j
                End With
            End With
        End With
        vbworksheet.Cells.Cells.EntireColumn.AutoFit()
        vbworksheet.Cells.Cells.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter
        ' Sfnaf(i) = System.IO.Path.GetFileNameWithoutExtension(StrFilename)

        '保存excel
        'Dim dirName() As String = FolderBrowserDialog1.SelectedPath.Split('\\')
        'Dim temp As String
        'temp = FolderBrowserDialog1.SelectedPath
        'Dim temp1 As String = Split(temp, "\")(UBound(Split(temp, "\")))
        vbworkbook.SaveAs(GetSavePath(br1.DirectoryPath))

    End Sub
    Private Function GetSavePath(ByVal path As String)
        Dim savepath As String
        savepath = path + "\" + Split(path, "\")(UBound(Split(path, "\")))
        Return savepath
    End Function
    Private Function GetCustomPropertys(ByVal filename As String, ByVal custompropertynames As ArrayList)
        Dim swApp As SldWorks
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swCustProp As CustomPropertyManager
        'val变量存储 value
        Dim val As String
        'valout变量存储evaluated value
        Dim valout As String
        Dim custompropertyname As String
        Dim evaluatedValues As ArrayList
        swApp = New SldWorks()
        swApp.Visible = False
        swModel = swApp.OpenDoc(filename, 1)
        swModelDocExt = swModel.Extension
        swCustProp = swModelDocExt.CustomPropertyManager("")
        evaluatedValues = New ArrayList()
        For Each custompropertyname In custompropertynames
            swCustProp.Get2(custompropertyname, val, valout)

            evaluatedValues.Add(valout)
        Next

        swApp.CloseAllDocuments(True)
        Return evaluatedValues

    End Function

End Class
