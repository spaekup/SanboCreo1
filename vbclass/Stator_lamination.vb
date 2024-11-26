Imports System.Math
Imports System.Runtime.InteropServices
Imports MySql.Data.MySqlClient
Imports pfcls
Imports System.IO
Imports System.Configuration
Imports System.Collections.Generic




Public Class Stator_lamination

    Dim AC As IpfcAsyncConnection = Nothing
        Dim CcAC As New CCpfcAsyncConnection
        Dim iparamValue As IpfcParamValue
        Dim iParameterOwner As IpfcParameterOwner
        Dim iParameter As IpfcParameter
        Dim window As IpfcWindow
        Dim session As IpfcSession
        Dim smodel, smodel_1, smodel_2, md, md1, md2, md3, md4, md5, md6, md7, md8, md9, solidDesc As IpfcModelDescriptor
        Dim drw, model, models_1, m, m1, m2, m3, m4, m5, m6, m7, m8, m9, solid As IpfcModel
        Dim models As IpfcModels
        Dim drawing, drawing1 As IpfcDrawing
        Dim drwFormat As IpfcDrawingFormat
        Dim assembly, assembly_1 As IpfcAssembly
        Dim part, ComponentModel, componentModel_1, componentModel_2, componentModel_3, s, s1, s2, s3, s4, s5, s6, s7, s8, s9 As IpfcSolid
        Dim asmcomp, asmcomp_1, asmcomp_2, asmcomp_3, asmcomp_4, asmcomp_5, asmcomp_6, asmcomp_7, asmcomp_8, asmcomp_9 As IpfcComponentFeat
        Dim constraints, constraints1 As IpfcComponentConstraints
        Dim asmItem_1, asmItem_2, asmItem_3, compItem_1, compItem_2, compItem_3, compItem_4, compItem_5, compItem_6 As IpfcModelItem
        Dim ids As Cintseq
        Dim path As IpfcComponentPath
        Dim compSelect_1, compSelect_2, compSelect_3， compSelect_4, compSelect_5， compSelect_6 As IpfcSelection
        Dim constraint， constraint1, constraint2 As IpfcComponentConstraint
        Dim matrix, matrix_1, matrix_2, matrix_3 As New CpfcMatrix3D
        Dim transform3D_1, transform3D_2 As IpfcTransform3D
        Dim M3d As IpfcMatrix3D
        Dim outline As IpfcOutline3D
        Dim p3d As IpfcPoint3D
        Dim view As IpfcView
        Dim viewer As IpfcViewOwner
        Dim excludes As IpfcModelItemTypes
        Dim Scale As Double
        Dim se As IpfcSession
        Dim row, col As Integer
        Dim view2D As IpfcView2D = Nothing
        Dim name_view As IpfcView
        Dim instrs As IpfcGeneralViewCreateInstructions
        Dim drawingOptions As New CpfcDrawingCreateOptions
        Dim viewdisplay As IpfcViewDisplay
        Dim view2Ds As IpfcView2Ds
        Dim i As Integer
        Dim viewName As String
        Dim displayStyle As String
        Dim sheetNo As Integer


        ' 数据库
        Dim mysqlcon, Dmysqlcon, connect1, connect2, connect3, connect4 As MySqlConnection
        Dim mysqlcom, Dmysqlcom, cursor1, cursor2, cursor3, cursor4 As MySqlCommand
        Public read, read1, read2, Dread, Dread1, Dread2 As MySqlDataReader
        Dim fff As MySqlParameter
        Dim ffff As MySqlParameterCollection
        Dim database_name1, database_name2, database_name3, database_name4, database_name5,
            table_name1, table_name2, table_name3, table_name4, table_name5,
            updata1, updata2, updata3,
            updata4, updata5, updata6, updata7, updata8, updata9, updata10, updata11, updata12, updata13,
            updata14, updata15, updata16, updata17, updata18 As String
        Dim values1, values2, values3, values4, values5 As Object()

        Dim ID As Integer = 1

        Public Sub mysql_new2(localhost$, database$, tabledase$, code$)
            mysqlcon = New MySqlConnection("server=" + localhost + ";userid=root" & ";password=123456" & ";database=" + database + ";pooling=false")
            mysqlcon.Open()
            mysqlcom = New MySqlCommand("select * from " + tabledase, mysqlcon)
            read = mysqlcom.ExecuteReader()
            read.Read()
            Do Until read("code") = code$
                read.Read()
            Loop
        End Sub
    Public Sub mysql_new3(localhost$, database$, tabledase$, ID$)
        mysqlcon = New MySqlConnection("server=" + localhost + ";userid=root" & ";password=123456" & ";database=" + database + ";pooling=false")
        mysqlcon.Open()
        mysqlcom = New MySqlCommand("select * from " + tabledase, mysqlcon)
        read = mysqlcom.ExecuteReader()
        read.Read()
        Do Until read("ID") = ID$
            read.Read()
        Loop
    End Sub
    Public Sub openAPP(VersionNumber#)
            If VersionNumber = 2.0 Then
                AC = CcAC.Start("C:\Users\Public\Desktop\Creo Parametric 2.0", ".")
            ElseIf VersionNumber = 7.0 Then
                AC = CcAC.Start("C:\Users\Public\Desktop\parametric.exe", ".")
            End If
            AC.Session.LoadConfigFile("D:\Creo\trail_dir\config.pro")
            AC.Session.ChangeDirectory("D:\Creo\trail_dir")
        End Sub
        Public Sub endApp()

            If Not AC Is Nothing Then
                If AC.IsRunning Then
                    AC.End()
                End If
            End If

        End Sub
        Public Sub setWorkDirectory()

        End Sub
        Public Sub listModelTeatures()

        End Sub
        Public Function activateModel(partName$, modelType#) As IpfcModel

            If Not AC Is Nothing And AC.IsRunning Then
                session = AC.Session
                model = session.GetModel(partName, modelType)
                window = session.CreateModelWindow(model)
                model.Display()
                window.Activate()
                model.Regenerate(Nothing)
                AC.Session.CurrentWindow.Refresh()
            End If
            activateModel = model

        End Function
        Public Function activate(partName$, modelType#) As IpfcModel
            If Not AC Is Nothing And AC.IsRunning Then
                session = AC.Session
                model = session.GetModel(partName, modelType)
                window = session.CreateModelWindow(model)
                model.Display()
                window.Activate()
            End If
            activate = model

        End Function
        Public Function retrieveModel(modelType$, modelPath$) As IpfcModel
            If modelType = "asm" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_ASSEMBLY, modelPath, Nothing)
            ElseIf modelType = "prt" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_PART, modelPath, Nothing)
            ElseIf modelType = "drw" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DRAWING, modelPath, Nothing)
            ElseIf modelType = "frm" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DWG_FORMAT, modelPath, Nothing)

            End If
            model = AC.Session.RetrieveModel(smodel) '载入模型

        End Function
        ''' <summary>
        ''' 添加参数
        ''' </summary>
        ''' <param name="model">模块</param>
        ''' <param name="paramName$">参数名</param>
        ''' <param name="paramValue$">参数值</param>
        ''' <param name="paramType$">参数类型</param>
        Public Sub addParam(model As IpfcModel, paramName$, paramValue$， paramType$)

            If model IsNot Nothing Then
                If (paramType = "浮点型") Then
                    iparamValue = (New CMpfcModelItem).CreateDoubleParamValue(Double.Parse(paramValue))
                ElseIf (paramType = "整型") Then
                    iparamValue = (New CMpfcModelItem).CreateIntParamValue(Int32.Parse(paramValue))
                ElseIf (paramType = "字符串") Then
                    iparamValue = (New CMpfcModelItem).CreateStringParamValue(paramValue)
                ElseIf (paramType = "布尔型") Then
                    iparamValue = (New CMpfcModelItem).CreateBoolParamValue(Boolean.Parse(paramValue))
                Else
                    iparamValue = (New CMpfcModelItem).CreateNoteParamValue(Long.Parse(paramValue))
                End If

                iParameterOwner = CType(model, IpfcParameterOwner)
                iParameterOwner.CreateParam(paramName, iparamValue)
            End If

        End Sub
        ''' <summary>
        ''' 修改参数
        ''' </summary>
        ''' <param name="model"></param>
        ''' <param name="paramName$"></param>
        ''' <param name="paramValue$"></param>
        ''' <param name="paramType$"></param>
        Public Sub setParamValue(model As IpfcModel, paramName$, paramValue$， paramType$， Optional angle# = 0)

            If model IsNot Nothing Then
                iParameterOwner = CType(model, IpfcParameterOwner)
                iParameter = iParameterOwner.GetParam(paramName)
                iparamValue = iParameter.GetScaledValue
                If paramType = "浮点型" Then
                    iparamValue.DoubleValue = Double.Parse(paramValue)
                ElseIf paramType = "字符串" Then
                    iparamValue.StringValue = paramValue
                ElseIf paramType = "布尔型" Then
                    iparamValue.BoolValue = Boolean.Parse(paramValue)
                ElseIf paramType = "角度" Then
                    iparamValue.DoubleValue = angle
                Else
                    iparamValue.BoolValue = Long.Parse(paramValue)
                End If
                iParameter.SetScaledValue(iparamValue, Nothing)
                'AC.Session.RunMacro("imi  ~Command `ProCmdRegenPart` ;")

            End If

        End Sub
        Public Sub partValue(model As IpfcModel)
            Dim paramName As String()
            Dim paramValue As String()
            paramName = {"SOURCE", "MATERIAL", "REMARK", "SHEET_SIZE",
                "DESCRIPTION", "OPT_LEVEL", "MARK", "MARK_A", "MARK_B",
                 "PTC_MATERIAL_NAME", "VERSION", "MATERIAL_CODE", "MT_LEVEL"}
            paramValue = {"外购件", "45号钢", "备注", "A3",
                "描述", "优选", "A", "标记A", "标记B",
                 "TBD", "版本", "物料编码", "基地级"}

            If model IsNot Nothing Then
                For i = 0 To paramName.Count - 1

                    iParameterOwner = CType(model, IpfcParameterOwner)
                    iParameter = iParameterOwner.GetParam(paramName(i))
                    iparamValue = iParameter.GetScaledValue
                    iparamValue.StringValue = paramValue(i)
                    iParameter.SetScaledValue(iparamValue, Nothing)

                    'AC.Session.RunMacro("imi  ~Command `ProCmdRegenPart` ;")
                Next
            End If

        End Sub
        Public Sub partValue_new(model As IpfcModel)
            Dim paramName As String()
            Dim paramValue As String()
            paramName = {"SOURCE", "MATERIAL", "REMARK", "SHEET_SIZE",
                "DESCRIPTION", "OPT_LEVEL", "MARK", "MARK_A", "MARK_B",
                 "VERSION", "MATERIAL_CODE", "MT_LEVEL"}
            paramValue = {read("SOURCE"), read("MATERIAL"), read("REMARK"), read("SHEET_SIZE"),
                read("DESCRIPTION"), read("OPT_LEVEL"), read("MARK"), read("MARK_A"),
                read("MARK_B")， read("VERSION"), read("MATERIAL_CODE"), read("MT_LEVEL")}

            If model IsNot Nothing Then
                For i = 0 To paramName.Count - 1

                    iParameterOwner = CType(model, IpfcParameterOwner)
                    iParameter = iParameterOwner.GetParam(paramName(i))
                    iparamValue = iParameter.GetScaledValue
                    iparamValue.StringValue = paramValue(i)
                    iParameter.SetScaledValue(iparamValue, Nothing)

                    'AC.Session.RunMacro("imi  ~Command `ProCmdRegenPart` ;")
                Next
            End If

        End Sub
        Public Sub deleteParam(model As IpfcModel, paramName$)

            If model IsNot Nothing Then
                iParameterOwner = CType(model, IpfcParameterOwner)
                iParameter = iParameterOwner.GetParam(paramName)
                iParameter.Delete()
            End If
        End Sub
        Public Sub regenerate() '激活模块

            AC.Session.RunMacro("mapkey 1 ~ Command `ProCmdEnvShadedEdges` ;")

        End Sub
        '删除文件夹下面的所有内容：包括文件，文件夹。
        Public Sub DeleteFoldeSubFF(ByVal fpath As String)
            Try
                For Each fd As String In Directory.GetDirectories(fpath)
                    DeleteFolder(fd)
                Next
                For Each fi As String In Directory.GetFiles(fpath)
                    DeleteFile(fi)
                Next
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information)
            End Try
            'Try
            '    For Each fd As DirectoryInfo In GetFolderS(fpath)
            '        DeleteFolder(fd.FullName)
            '    Next
            '    For Each fi As FileInfo In GetFileS(fpath)
            '        DeleteFile(fi.FullName)
            '    Next
            'Catch ex As Exception
            '    MsgBox(ex.Message, MsgBoxStyle.Information)
            'End Try
        End Sub
        '删除文件。
        Public Sub DeleteFile(ByVal fpath As String)
            If IO.File.Exists(fpath) Then
                '删除文件file的方法1:删除到回收站里面。
                My.Computer.FileSystem.DeleteFile(fpath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                '删除文件file的方法2:直接从硬盘上删除。
                'IO.File.Delete(fpath)
            End If
        End Sub
        '删除文件夹。
        Public Sub DeleteFolder(ByVal folder As String)
            If IO.Directory.Exists(folder) Then
                '删除文件夹folder的方法1:删除到回收站里面。
                My.Computer.FileSystem.DeleteDirectory(folder, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin, FileIO.UICancelOption.DoNothing)
                '删除文件夹folder的方法2:直接从硬盘上删除。
                'IO.Directory.Delete(folder, True)
            End If
        End Sub
        Public Sub fileBackup(modelType$, modelPath$)

            AC = CcAC.Start("D:\Creo\Creo 2.0\Parametric\bin\parametric.exe", ".")
            m1 = retrieveModel(modelType, modelPath)
            If modelType = "asm" Then
                m1 = activateModel(Right(modelPath, 17), 0)
            ElseIf modelType = "prt" Then
                m1 = activateModel(Right(modelPath, 17), 1)
            ElseIf modelType = "drw" Then
                m1 = activateModel(Right(modelPath, 17), 2)
            End If
            m1.Rename("transition", True)
            smodel = m1.Descr
            smodel.Path = "E:\works\Creo\Process file"
            m1.Backup(smodel)
            endApp()

        End Sub
        Public Function openModel(modelType$, modelPath$) As IpfcModel

            AC = CcAC.Start("D:\Creo\Creo 2.0\Parametric\bin\parametric.exe", ".")
        AC.Session.LoadConfigFile("D:\PTC\proe_stds\config.pro")
        'AC.Session.ChangeDirectory("E:\works\Creo\Process file")
        m1 = retrieveModel(modelType, modelPath)


        End Function
        Public Sub deleteFileChangeModel(modelType$, modelPath$)

        DeleteFoldeSubFF("E:\works\creo\Process file\")
        fileBackup(modelType, modelPath)
        m1 = openModel(modelType, "E:\works\Creo\Process file\transition." + modelType + ".1")

    End Sub
        Public Function openDrawing(drwFormat$, modelPath$) As IpfcModel

            If drwFormat = "a0" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DRAWING, modelPath, Nothing)
            ElseIf drwFormat = "a1" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DRAWING, modelPath, Nothing)
            ElseIf drwFormat = "a2" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DRAWING, modelPath, Nothing)
            ElseIf drwFormat = "a3" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DWG_FORMAT, modelPath, Nothing)
            ElseIf drwFormat = "a4" Then
                smodel = (New CCpfcModelDescriptor).Create(EpfcModelType.EpfcMDL_DRAWING, modelPath, Nothing)
            End If
            model = AC.Session.RetrieveModel(smodel) '载入模型
            model.Display()
            AC.Session.CurrentWindow.Activate() '激活模型
            openDrawing = model

        End Function
        Public Function CreatePoint(x#, y#, Optional z# = 0) As IpfcPoint3D

            p3d = New CpfcPoint3D
            p3d.Set(0, x)
            p3d.Set(1, y)
            p3d.Set(2, z)
            Return p3d

        End Function
        Public Function CreateTransfromMatrix() As IpfcTransform3D

            matrix = New CpfcMatrix3D
            For i = 0 To 3
                For j = 0 To 3
                    If i = j Then
                        matrix.Set(i, j, 1.0)
                    Else
                        matrix.Set(i, j, 0.0)
                    End If
                Next
            Next
            transform3D_1 = (New CCpfcTransform3D).Create(matrix)
            Return transform3D_1

        End Function
        Public Function matrixNormalize(matrix As IpfcMatrix3D) As IpfcMatrix3D
            Dim scale As Double
            Dim row, col As Integer

            matrix.Set(3, 0, 0.0)
            matrix.Set(3, 1, 0.0)
            matrix.Set(3, 2, 0.0)

            scale = Math.Sqrt(matrix.Item(0, 0) * matrix.Item(0, 0) + matrix.Item(0, 1) *
                              matrix.Item(0, 1) + matrix.Item(0, 2) * matrix.Item(0, 2))

            For row = 0 To 2
                For col = 0 To 2
                    matrix.Set(row, col, matrix.Item(row, col) / scale)
                Next
            Next

            matrixNormalize = matrix

        End Function
        Public Function transformNormalize(transform3D_1 As IpfcTransform3D) As IpfcTransform3D

            matrix = transform3D_1.Matrix
            transform3D_1 = (New CCpfcTransform3D).Create(matrixNormalize(matrix))
            Return transform3D_1
            ' Return (New CCpfcTransform3D).Create(matrixNormalize(matrix))

        End Function
        Public Function createDrawingFromTemplate(template$) As IpfcDrawing

            drawingOptions.Insert(0, EpfcDrawingCreateOption.EpfcDRAWINGCREATE_DISPLAY_DRAWING)
            drawingOptions.Insert(1, EpfcDrawingCreateOption.EpfcDRAWINGCREATE_SHOW_ERROR_DIALOG)
            session = AC.Session
            model = session.CurrentModel
            drawing = session.CreateDrawingFromTemlate(model.FullName, template, model.Descr, drawingOptions)
            createDrawingFromTemplate = drawing

        End Function
        Public Function listViews(drawing As IpfcDrawing) As IpfcView2Ds


            view2Ds = drawing.List2DViews
            For i = 0 To view2Ds.Count - 1

                view2D = view2Ds.Item(i)
                viewName = view2D.Name
                sheetNo = view2D.GetSheetNumber
                solid = view2D.GetModel
                solidDesc = solid.Descr

                outline = view2D.Outline
                Scale = view2D.Scale
                viewdisplay = view2D.Display
                displayStyle = "unknown"

                Select Case viewdisplay.Style
                    Case EpfcDisplayStyle.EpfcDISPSTYLE_DEFAULT
                        displayStyle = "default"
                    Case EpfcDisplayStyle.EpfcDISPSTYLE_HIDDEN_LINE
                        displayStyle = "hidden line"
                    Case EpfcDisplayStyle.EpfcDISPSTYLE_NO_HIDDEN
                        displayStyle = "no hidden"
                    Case EpfcDisplayStyle.EpfcDISPSTYLE_SHADED
                        displayStyle = "shaded"
                    Case EpfcDisplayStyle.EpfcDISPSTYLE_WIREFRAME
                        displayStyle = "wireframe"
                End Select

            Next
        End Function
        Public Function Hd(a) As Double
            Return a * PI / 180   ' 函数具有返回值

        End Function

    Public Sub creo_stator_lamination(localhost$, ID As Integer, path$)

        '数据库
        mysql_new3(localhost$, "creo", "stator_lamination", ID)
        ' 零件图
        deleteFileChangeModel("prt", "E:\works\creo\creo_stator_lamination\" + read("oldcode") + ".prt.1")
        m6 = activateModel("transition", 1)
        partValue_new(m6)
        m6.Regenerate(Nothing)
        If read("type") = 1 Then
            setParamValue(m6, "d5", read("outer_diameter"), "浮点型") '外径
            setParamValue(m6, "d4", read("thickness"), "浮点型") '厚度
            setParamValue(m6, "d7", read("inside_diameter"), "浮点型") '内径
            setParamValue(m6, "d9", "", "角度"， Hd(read("slot_angle"))) '槽斜角度
            setParamValue(m6, "d10", read("slot_wide_mouth"), "浮点型") '槽口宽
            setParamValue(m6, "d11", read("slot_wide_neck"), "浮点型") '槽肩宽
            setParamValue(m6, "d12", read("slot_diameter"), "浮点型") '槽圆角直径
            setParamValue(m6, "d13", read("slot_high_mouth"), "浮点型") '槽口高
            setParamValue(m6, "d14", read("slot_high_shoulder"), "浮点型") '槽肩高
            setParamValue(m6, "d15", read("slot_high_neck"), "浮点型") '槽颈高
            setParamValue(m6, "p19", read("slot_number"), "浮点型") '槽数
            setParamValue(m6, "d22", "", "角度"， Hd(read("tag_angle"))) '标记与Y轴偏角
            setParamValue(m6, "d23", read("tag_high"), "浮点型") '标记高
            setParamValue(m6, "d24", read("tag_wide"), "浮点型") '标记宽
        ElseIf read("type") = 2 Then
            setParamValue(m6, "d5", read("outer_diameter"), "浮点型") '外径
            setParamValue(m6, "d4", read("thickness"), "浮点型") '厚度
            setParamValue(m6, "d7", read("inside_diameter"), "浮点型") '内径
            setParamValue(m6, "d9", "", "角度"， Hd(read("slot_angle"))) '槽斜角度
            setParamValue(m6, "d10", read("slot_wide_mouth"), "浮点型") '槽口宽
            setParamValue(m6, "d12", read("slot_wide_neck"), "浮点型") '槽肩宽
            setParamValue(m6, "d13", read("slot_diameter"), "浮点型") '槽圆角直径
            setParamValue(m6, "d11", read("slot_high_mouth"), "浮点型") '槽口高
            setParamValue(m6, "d15", read("slot_high_shoulder"), "浮点型") '槽肩高
            setParamValue(m6, "d14", read("slot_high_neck"), "浮点型") '槽颈高
            setParamValue(m6, "p19", read("slot_number"), "浮点型") '槽数
            setParamValue(m6, "d22", "", "角度"， Hd(read("tag_angle"))) '标记与Y轴偏角
            setParamValue(m6, "d23", read("tag_diameter"), "浮点型") '标记直径
            setParamValue(m6, "d26", read("sl_wide"), "浮点型") '顶槽宽
            setParamValue(m6, "d27", read("sl_y_high"), "浮点型") '顶槽高
            setParamValue(m6, "d28", "", "角度", Hd(read("sl_angle"))) '顶槽角度
            setParamValue(m6, "p32", read("sl_number"), "浮点型") '顶槽数
        ElseIf read("type") = 3 Then
            setParamValue(m6, "d5", read("outer_diameter"), "浮点型") '外径
            setParamValue(m6, "d4", read("thickness"), "浮点型") '厚度
            setParamValue(m6, "d7", read("inside_diameter"), "浮点型") '内径
            setParamValue(m6, "d9", "", "角度"， Hd(read("slot_angle"))) '槽斜角度
            setParamValue(m6, "d10", read("slot_wide_mouth"), "浮点型") '槽口宽
            setParamValue(m6, "d12", read("slot_wide_neck"), "浮点型") '槽肩宽
            setParamValue(m6, "d13", read("slot_diameter"), "浮点型") '槽圆角直径
            setParamValue(m6, "d11", read("slot_high_mouth"), "浮点型") '槽口高
            setParamValue(m6, "d15", read("slot_high_shoulder"), "浮点型") '槽肩高
            setParamValue(m6, "d14", read("slot_high_neck"), "浮点型") '槽颈高
            setParamValue(m6, "p19", read("slot_number"), "浮点型") '槽数
            setParamValue(m6, "d25", "", "角度"， Hd(read("tag_angle"))) '标记与Y轴偏角
            setParamValue(m6, "d22", read("tag_long_high"), "浮点型") '标记长高
            setParamValue(m6, "d23", read("tag_high"), "浮点型") '标记高
            setParamValue(m6, "d24", read("tag_wide"), "浮点型") '标记宽
            setParamValue(m6, "d28", read("sl_diameter"), "浮点型") '顶槽直径
            setParamValue(m6, "d27", "", "角度"， Hd(read("sl_angle"))) '顶槽角度
            setParamValue(m6, "p32", read("sl_number"), "浮点型") '顶槽数
        End If
        m6.Regenerate(Nothing)
        m6.Rename(read("code"), True)
        smodel = m6.Descr
        smodel.Path = path
        m6.Backup(smodel)
        endApp()

    End Sub
End Class






