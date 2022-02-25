Imports System.Data
Imports System.Data.OleDb

Namespace DynamicComponents

    <ComClass(DataManager.ClassId, DataManager.InterfaceId, DataManager.EventsId)> _
    Public Class DataManager



#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "979fe8cf-41c9-4d64-be82-fd7e73c18c38"
        Public Const InterfaceId As String = "89c73d6d-017c-4eb5-8ae8-7a86a78f9592"
        Public Const EventsId As String = "0ee9d8ed-a7ff-4fea-88e9-c6660ff54bd5"
#End Region

        ' A creatable COM class must have a Public Sub New() 
        ' with no parameters, otherwise, the class will not be 
        ' registered in the COM registry and cannot be created 
        ' via CreateObject.

        Public Sub New()
            MyBase.New()
        End Sub
        Private m_KeyFields As String
        Private Col_ControlName As New Collection()
        Private Col_ControlIndex As New Collection()
        Private Col_KeyValue As New Collection()
        Private Col_KeyText As New Collection()
        Private Col_KeyFields As New Collection()
        Private Col_MasterFields As New Collection()
        Private Col_DetailFields As New Collection()
        Private Col_RequiredFields As New Collection()
        Private Col_GridKeyValue As New Collection()
        Private Col_FieldsType(30, 1) As String
        Private Col_FieldsTypePos As Byte = 0
        Private m_SpecialChars As String
        Private Key_ZeroPad As Byte = 0
        Private m_KeyLeaveField As String = ""
        Private KeyLeavePos As Byte = 0
        Private FilterString As String
        Private DummyFilterString As String = ""
        Private HasGrid As Boolean = False
        Private oMaster As New DataTable
        Private oDetails As New DataTable
        Private MyForm As New System.Windows.Forms.Form()
        Private MyGrid As New DataGridView
        Private m_DecimalPlaces As Byte
        Private ImagePath As String
        Private ImageButtons() As String
        Private ImageMotion As Boolean = False
        Private FlipState As Boolean = False
        Private Col_GridFields As New Collection()
        Private m_MasterFlagField As String = ""
        Private m_DetailFlagField As String = ""
        Private m_FlagValue As String = ""
        Private HelpIdSender As String
        Private HasImage As Boolean = False
        Private RequiredFields_Msg As String = "Uncomplete Entries"
        Private RequiredFields_ShowMsg As Boolean = True
        Private m_HoldSaving As Boolean = False
        Private m_ReleaseSaving As Boolean = False
        Dim ds As New DataSet
        Dim bm As BindingManagerBase
        Dim bm_dt As BindingManagerBase
        Dim bm_st As BindingManagerBase
        Dim bm_ng As BindingManagerBase
        Dim bm_hp As BindingManagerBase
        Dim da_master As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
        Dim da_detials As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
        Dim da_rec As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
        Dim bs_dt As BindingSource = New BindingSource()
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter
        Private Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
            On Error GoTo EndMe
            If str_String <> "" Then
                Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
            End If
EndMe:
            Return str_String
        End Function
        Public Sub InitForm(ByRef dm_adapter As OleDb.OleDbDataAdapter, ByRef dm_OleConnection As OleDb.OleDbConnection, ByRef ds As DataSet, ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)


            Dim TxtCtrl As New Control()
            Dim X As Byte

            If Not (dm_Grid Is Nothing) Then
                HasGrid = True
            End If

            MyGrid = dm_Grid
            MyForm = dm_Form
            oMaster = dm_MasterTable
            oDetails = dm_DetailsTable
            bm = dm_Form.BindingContext(ds, dm_MasterTable.TableName)
            bm_dt = dm_Form.BindingContext(ds, dm_DetailsTable.TableName)
            bs_dt.DataSource = dm_DetailsTable
            da = dm_adapter

            If HasGrid Then
                AddHandler dm_Grid.RowsAdded, AddressOf dm_Grid_OnAddNew    'maa
                'AddHandler dm_Grid.CellEndEdit, AddressOf dm_Grid_AfterColEdit  'maa 
                AddHandler dm_Grid.KeyDown, AddressOf dm_Grid_KeyDown
            End If
            AddHandler dm_Form.Paint, AddressOf MyForm_Paint

            PrepareHelp()
            ReadInitialValues()

            CN = dm_OleConnection
            X = 0

            For Each TxtCtrl In dm_Form.Controls
                If TypeName(TxtCtrl) = "TextBox" Or TypeName(TxtCtrl) = "ComboBox" Or TypeName(TxtCtrl) = "ListBox" Or TypeName(TxtCtrl) = "CheckBox" Or TypeName(TxtCtrl) = "RadioButton" Then
                    Col_ControlName.Add(TxtCtrl)
                    Col_ControlIndex.Add(X)
                End If
                X += 1
            Next TxtCtrl


            If HasGrid Then
                MyGrid.DataSource = oDetails
            End If

        End Sub
        Private Sub MyForm_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
            Static Flag As Byte = 0

            Flag += 1
            If Flag = 1 Then
                If Col_KeyValue.Count = 8 Then
                    Dim Num As Byte
                    For Num = 1 To 16
                        Col_KeyValue.Add(KeyLeavePos)
                    Next Num
                ElseIf Col_KeyValue.Count = 16 Then
                    Dim Num As Byte
                    For Num = 1 To 8
                        Col_KeyValue.Add(KeyLeavePos)
                    Next Num
                End If
                If Col_GridKeyValue.Count = 7 Then
                    Dim Num As Byte
                    For Num = 1 To 14
                        Col_GridKeyValue.Add(KeyLeavePos)
                    Next Num
                ElseIf Col_GridKeyValue.Count = 14 Then
                    Dim Num As Byte
                    For Num = 1 To 7
                        Col_GridKeyValue.Add(KeyLeavePos)
                    Next Num
                End If

            End If

            If HelpRtnID <> "" Then
                If UCase(sender.Controls(Col_KeyValue(8)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(8)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(1)).Text = HelpRtnName
                ElseIf UCase(sender.Controls(Col_KeyValue(16)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(16)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(9)).Text = HelpRtnName
                ElseIf UCase(sender.Controls(Col_KeyValue(24)).Name) = UCase(HelpIdSender) Then
                    sender.Controls(Col_KeyValue(24)).Text = HelpRtnID
                    sender.Controls(Col_KeyValue(17)).Text = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(6)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(6) + 2 <= MyGrid.Columns.Count Then
                        'MyGrid.Col = Col_GridKeyValue(6) + 2  //maa
                    Else
                        'MyGrid.Col = Col_GridKeyValue(6) - 2 // maa
                    End If
                    MyGrid.Columns(Col_GridKeyValue(6)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(6) + 1).Value = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(13)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(13) + 2 <= MyGrid.Columns.Count Then
                        'MyGrid.Col = Col_GridKeyValue(13) + 2 // maa
                    Else
                        'MyGrid.Col = Col_GridKeyValue(13) - 2 // maa
                    End If
                    MyGrid.Columns(Col_GridKeyValue(13)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(13) + 1).Value = HelpRtnName
                ElseIf UCase(MyGrid.Name) + "_" + UCase(MyGrid.Columns(Col_GridKeyValue(20)).DataField) = UCase(HelpIdSender) Then
                    If Col_GridKeyValue(20) + 2 <= MyGrid.Columns.Count Then
                        'MyGrid.Col = Col_GridKeyValue(20) + 2 // maa
                    Else
                        'MyGrid.Col = Col_GridKeyValue(20) - 2 // maa
                    End If
                    MyGrid.Columns(Col_GridKeyValue(20)).Value = HelpRtnID
                    MyGrid.Columns(Col_GridKeyValue(20) + 1).Value = HelpRtnName
                End If
            End If

        End Sub
        Public Sub PopulateForm(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Dim Num As Byte
            Dim Num2 As Integer
            Dim CtrlName As String
            Dim myCheckBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton
            On Error Resume Next 'keep it very important


            For Num = 1 To Col_ControlName.Count()
                CtrlName = Col_ControlName(Num).Name
                If UCase(Left(CtrlName, 1)) <> "X" Then
                    'If Not IsDBNull(dm_MasterTable(CtrlName).Value) Then
                    'If dm_MasterTable(CtrlName).Value <> "" Then
                    If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                        dm_Form.Controls(Col_ControlIndex(Num)).Text = dm_MasterTable.Rows(bm.Position)(CtrlName)
                    ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                        myCheckBox = dm_Form.Controls(Col_ControlIndex(Num))
                        myCheckBox.Checked = IIf(dm_MasterTable.Rows(bm.Position)(CtrlName) = True, True, False)
                    Else 'RadioButton
                        myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                        myRadioButton.Checked = IIf(dm_MasterTable.Rows(bm.Position)(CtrlName) = True, True, False)
                    End If
                    'End If
                    'End If
                End If
            Next Num

            For Num = 1 To Col_KeyText.Count
                Col_KeyText.Remove(Num)
            Next Num
            For Num = 1 To Col_KeyFields.Count / 2
                Col_KeyText.Add(dm_Form.Controls(Col_KeyFields(Num * 2)).Text)
            Next Num

            For Num = 1 To Col_KeyValue.Count() / 8
                Num2 = ((Num - 1) * 8) + 1
                dm_Form.Controls(Col_KeyValue(Num2)).Text = GetRelatedValue(Col_KeyValue(Num2 + 2), Col_KeyValue(Num2 + 3), dm_MasterTable.Rows(bm.Position)(Col_KeyValue(Num2 + 4)), Col_KeyValue(Num2 + 5))
            Next Num
            If HasGrid Then
                PopulateGrid(dm_Grid, dm_MasterTable, dm_DetailsTable)
            End If
            'dm_Form.Controls(KeyLeavePos).Focus()
        End Sub
        Public Sub GoFirst(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
            If bm.Position < bm.Count Then
                bm.Position = 0
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
            End If

        End Sub

        Public Sub GoLast(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
            If bm.Position < bm.Count Then
                bm.Position = bm.Count - 1
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
            End If

        End Sub
        Public Sub GoNext(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
            If bm.Position < bm.Count Then
                bm.Position += 1
                If bm.Position >= bm.Count Then
                    bm.Position = bm.Count - 1
                End If
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
            End If
        End Sub
        Public Sub GoPrevious(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
            If bm.Position < bm.Count Then
                bm.Position -= 1
                If bm.Position >= bm.Count Then
                    bm.Position = 0
                End If
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
            End If
        End Sub
        Public Sub ClearData(ByRef dm_Form As System.Windows.Forms.Form, Optional ByVal dm_DetailsTable As DataTable = Nothing, Optional ByRef dm_Grid As DataGridView = Nothing)
            Dim Num As Byte
            Dim myChechBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton

            For Num = 1 To Col_ControlIndex.Count()
                If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                    dm_Form.Controls(Col_ControlIndex(Num)).Text = ""
                ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                    myChechBox = dm_Form.Controls(Col_ControlIndex(Num))
                    myChechBox.Checked = False
                Else 'RadioButton
                    myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                    myRadioButton.Checked = False
                End If
            Next Num
            If HasGrid = True Then
                dm_DetailsTable.Select("")

                'dm_Grid.DataSource = Nothing // maa new and 3 more
                'dm_Grid.Rows.Clear()
                'dm_Grid.DataSource = dm_DetailsTable
                dm_Grid.Refresh()

            End If
        End Sub
        Public Sub NewRecord(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)

            If HasGrid = True Then
                Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
            Else
                Me.ClearData(dm_Form)
            End If
            If bm.Position < bm.Count Then
                bm.Position = bm.Count - 1
            End If
            If m_KeyLeaveField <> "" Then
                If bm.Position < bm.Count Then
                    If Not IsDate(dm_MasterTable.Rows(bm.Position)(m_KeyLeaveField) And Val(dm_MasterTable.Rows(bm.Position)(m_KeyLeaveField)) <> 0) Then
                        dm_Form.Controls(KeyLeavePos).Text = ZeroPad(dm_MasterTable.Rows(bm.Position)(m_KeyLeaveField) + 1, Key_ZeroPad)
                    End If
                End If
                dm_Form.Controls(KeyLeavePos).Focus()
            End If

        End Sub
        Public Sub DeleteRecord(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Dim cFilter As String = ""
            Dim Num As Byte
            Dim sql_delete As String = ""

            Dim command As New OleDb.OleDbCommand
            da.DeleteCommand = CN.CreateCommand()


            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num
            If m_MasterFlagField <> "" Then
                cFilter += " and " + m_MasterFlagField + " ='" + m_FlagValue + "'"
            End If

            If m_KeyFields = "" Then
                MsgBox("You must assign KeyField Property first ,so you can delete records", , "DC DataManger error msg")
                Return
            End If

            If cFilter <> "" Then
                If bm.Position < bm.Count Then
                    dm_MasterTable.Select(cFilter)
                End If

                If bm.Position < bm.Count Then
                    If MsgBox(aInitValues(20), MsgBoxStyle.YesNo) = MsgBoxResult.Yes And bm.Position < bm.Count Then
                        sql_delete = "DELETE from " + dm_MasterTable.TableName + " Where " + cFilter
                        da.DeleteCommand.CommandText = sql_delete
                        da.DeleteCommand.ExecuteNonQuery()
                        'dm_MasterTable.AcceptChanges() // maa new

                        dm_MasterTable.Select("")
                        If HasGrid Then
                            sql_delete = "DELETE from " + dm_DetailsTable.TableName + " Where " + cFilter
                            da.DeleteCommand.CommandText = sql_delete
                            da.DeleteCommand.ExecuteNonQuery()
                            'dm_DetailsTable.AcceptChanges() // maa new
                            Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
                        End If

                        GoPrevious(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
                    End If
                End If
            End If
        End Sub
        Private Sub SubSaveData(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Dim Num As Byte
            Dim CtrlName As String
            Dim CtrlValue As String
            Dim myCheckBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton
            Dim sql_update As String = " UPDATE " + dm_MasterTable.TableName + " set "
            Dim cFilter As String = ""

            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num


            da.SelectCommand = New OleDb.OleDbCommand("Select * from OrderDetails", CN)
            da.Fill(ds, "OrderDetails")

            da.SelectCommand = New OleDb.OleDbCommand("Select * from Orders", CN)
            da.Fill(ds, "Orders")


            For Num = 1 To Col_ControlName.Count()
                'dm_MasterTable.Rows(bm.Position).BeginEdit()
                CtrlName = Col_ControlName(Num).Name
                If UCase(Left(CtrlName, 1)) <> "X" Then
                    If Num > 1 Then
                        sql_update += " , "
                    End If
                    If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                        CtrlValue = dm_Form.Controls(Col_ControlIndex(Num)).Text
                        If CtrlValue <> "" And Not IsDBNull(CtrlValue) Then
                            dm_MasterTable.Rows(bm.Position)(CtrlName) = CtrlValue
                            sql_update += CtrlName + "=" + CtrlValue
                        ElseIf CtrlValue = "" Then
                            dm_MasterTable.Rows(bm.Position)(CtrlName) = Nothing
                        End If
                    ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                        myCheckBox = dm_Form.Controls(Col_ControlIndex(Num))
                        dm_MasterTable.Rows(bm.Position)(CtrlName) = IIf(myCheckBox.Checked, True, False)
                    Else 'RadioButton
                        myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                        dm_MasterTable.Rows(bm.Position)(CtrlName) = IIf(myRadioButton.Checked, True, False)
                    End If

                End If
                'dm_MasterTable.Rows(bm.Position).AcceptChanges()
            Next Num

            If m_MasterFlagField <> "" Then
                dm_MasterTable.Rows(bm.Position)(m_MasterFlagField) = m_FlagValue
                'dm_MasterTable.Rows(bm.Position).AcceptChanges()
            End If

            sql_update += " WHERE  " + cFilter

            Dim command As New OleDb.OleDbCommand(sql_update, CN)
            'command.ExecuteNonQuery()




            Dim cmBuilder As New OleDbCommandBuilder(da)
            Dim tmpchanges As DataSet = ds.GetChanges(DataRowState.Unchanged)
            Dim tmpdelete As DataSet = ds.GetChanges(DataRowState.Deleted)
            Dim tmpdaded As DataSet = ds.GetChanges(DataRowState.Added)
            da.Update(ds, "Orders")
            da.Update(ds, "OrderDetails")
            'da.Update(ds.Tables(3))
            ds.AcceptChanges()


        End Sub
        Public Sub HoldSaving(ByVal Mode As Boolean)
            If Mode = True Then
                m_HoldSaving = True
            Else : m_HoldSaving = False
            End If
        End Sub
        Public Sub ReleaseSaving(ByVal Mode As Boolean)
            If Mode = True Then
                m_ReleaseSaving = True
                m_HoldSaving = False
            Else : m_ReleaseSaving = False
            End If

            If m_ReleaseSaving Then
                SaveData(MyForm, oMaster, MyGrid, oDetails)
            End If
            m_ReleaseSaving = False
        End Sub
        Public Sub SaveData(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Dim cFilter As String = ""
            Dim Num As Byte

            If Not ValidateForm(dm_Form) Then Exit Sub
            If m_ReleaseSaving Then GoTo ReleaseSaving

            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num

            If m_MasterFlagField <> "" Then
                cFilter += " and " + m_MasterFlagField + " ='" + m_FlagValue + "'"
            End If

            If cFilter <> "" Then
                If bm.Position < bm.Count Then
                    dm_MasterTable.Select(cFilter)
                End If

                If bm.Position >= bm.Count Then
                    'dm_MasterTable.AddNew() // maa
                End If

                Me.SubSaveData(dm_Form, dm_MasterTable)
ReleaseSaving:
                If Not m_HoldSaving Then
                    'dm_MasterTable.Update() // maa
                    dm_MasterTable.Select("") ' order of line is important
                    Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
                End If
            End If

            If m_KeyLeaveField <> "" And Not m_HoldSaving Then
                If bm.Position < bm.Count Then
                    bm.Position = bm.Count - 1
                    If Not IsDate(dm_MasterTable.Rows(bm.Position)(m_KeyLeaveField)) And Val(dm_MasterTable.Rows(bm.Position)(m_KeyLeaveField)) <> 0 Then
                        dm_Form.Controls(KeyLeavePos).Text = ZeroPad(dm_MasterTable.Rows(bm.Position)(m_KeyLeaveField) + 1, Key_ZeroPad)
                    End If
                End If
            End If

        End Sub
        Public Sub KeyLeave(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Dim Num As Byte
            Dim cFilter As String = ""
            Dim ColFields As New Collection()

            For Num = 1 To Col_KeyFields.Count / 2
                cFilter += dm_Form.Controls(Col_KeyFields(Num * 2)).Name + " = '" + dm_Form.Controls(Col_KeyFields(Num * 2)).Text + "'"
                ColFields.Add(dm_Form.Controls(Col_KeyFields(Num * 2)).Text)
                If Num <> Col_KeyFields.Count / 2 Then
                    cFilter += " and "
                End If
            Next Num
            If m_MasterFlagField <> "" Then
                cFilter += " and " + m_MasterFlagField + " ='" + m_FlagValue + "'"
            End If

            If cFilter <> "" Then
                If bm.Position < bm.Count Then
                    dm_MasterTable.Select(cFilter)
                End If

                If HasGrid Then
                    dm_DetailsTable.Select("")
                    Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
                Else
                    Me.ClearData(dm_Form)
                End If

                If bm.Position < bm.Count - 1 Then
                    Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
                Else
                    For Num = 1 To Col_KeyFields.Count / 2
                        dm_Form.Controls(Col_KeyFields(Num * 2)).Text = ColFields(Num)
                    Next Num
                End If
                dm_MasterTable.Select("")
            End If

        End Sub
        Public Function GetValue(ByVal str_Table As String, ByVal str_Key As String, ByVal str_value As String, ByVal str_RetField As String) As String
            Dim oRecSet As New DataTable()
            Dim str_where As String = str_Key + " = '" + str_value + "'"
            Dim ret As String = ""

            da_rec.SelectCommand = New OleDb.OleDbCommand("Select * from " + str_Table + " WHERE " + str_where, CN)
            da_rec.Fill(ds, str_Table)
            oRecSet = ds.Tables(str_Table)
            If (oRecSet.Rows.Count = 1) Then
                ret = oRecSet.Rows(0)(str_RetField)
            End If
            oRecSet.Clear()
            oRecSet.Dispose()

            Return ret

        End Function
        Public Sub AddRelatedValue(ByRef str_Table As String, ByVal str_Key As String, ByVal str_Control As String, ByVal str_RetValue As String, ByVal str_RetControl As String, Optional ByVal n_ZeroPad As Byte = 0)
            Dim Num As Byte
            Dim pos As Byte
            Static Flag As Byte = 0

            Flag += 1
            pos = ((Flag - 1) * 8) + 1
            For Num = 1 To Col_ControlName.Count
                If UCase(Col_ControlName(Num).Name) = UCase(str_RetControl) Then
                    Col_KeyValue.Add(Col_ControlIndex(Num))
                    Exit For
                End If
            Next Num

            For Num = 1 To Col_ControlName.Count
                If UCase(Col_ControlName(Num).Name) = UCase(str_Control) Then
                    If Flag = 1 Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf MyTextBox1_Leave
                        AddHandler txt_control.KeyDown, AddressOf MyTextBox1_KeyDown
                    ElseIf Flag = 2 Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf MyTextBox2_Leave
                        AddHandler txt_control.KeyDown, AddressOf MyTextBox2_KeyDown
                    ElseIf Flag = 3 Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf MyTextBox3_Leave
                        AddHandler txt_control.KeyDown, AddressOf MyTextBox3_KeyDown
                    End If
                    Exit For
                End If
            Next Num

            Col_KeyValue.Add(str_RetControl)
            Col_KeyValue.Add(str_Table)
            Col_KeyValue.Add(str_Key)
            Col_KeyValue.Add(str_Control)
            Col_KeyValue.Add(str_RetValue)
            Col_KeyValue.Add(n_ZeroPad)
            Col_KeyValue.Add(Col_ControlIndex(Num))
        End Sub
        Public Sub AddGridRelatedValue(ByVal str_Table As String, ByVal str_TableKey As String, ByVal str_Column As String, ByVal str_TableRetField As String, ByVal str_GridRetColumn As String, Optional ByVal n_ZeroPad As Byte = 0)
            Dim Num As Byte

            Col_GridKeyValue.Add(str_Table)
            Col_GridKeyValue.Add(str_TableKey)
            Col_GridKeyValue.Add(str_Column)
            Col_GridKeyValue.Add(str_TableRetField)
            Col_GridKeyValue.Add(str_GridRetColumn)

            For Num = 0 To MyGrid.Columns.Count - 1
                If UCase(MyGrid.Columns(Num).Name) = UCase(str_Column) Then
                    Col_GridKeyValue.Add(Num)
                    Col_GridKeyValue.Add(n_ZeroPad)
                    Exit Sub
                End If
            Next
            Col_GridKeyValue.Add(-1)
            Col_GridKeyValue.Add(n_ZeroPad)
        End Sub
        Private Function GetRelatedValue(ByVal str_Table As String, ByVal str_Key As String, ByVal str_value As String, ByVal str_RetField As String) As String
            Dim oRecSet As New DataTable()
            Dim str_where As String = str_Key + " = '" + str_value + "'"
            Dim ret As String = ""

            da_rec.SelectCommand = New OleDb.OleDbCommand("Select * from " + str_Table + " WHERE " + str_where, CN)
            da_rec.Fill(ds, str_Table)
            oRecSet = ds.Tables(str_Table)

            If (oRecSet.Rows.Count = 1) Then
                ret = oRecSet.Rows(0)(str_RetField)
            End If

            oRecSet.Clear()
            oRecSet.Dispose()

            Return ret

        End Function
        Public Sub FlagField(ByVal str_MasterFlagField As String, ByVal str_FlagValue As String, Optional ByVal str_DetailFlagField As String = "")
            Dim Num As Byte

            m_MasterFlagField = str_MasterFlagField
            m_FlagValue = str_FlagValue

            If HasGrid Then
                m_DetailFlagField = str_DetailFlagField
                For Num = 0 To MyGrid.Columns.Count - 1
                    'If UCase(MyGrid.Columns(Num).DataField) = UCase(m_DetailFlagField) Then // maa 3
                    'MyGrid.Columns(Num).Visible = False
                    'End If
                Next Num
            End If

        End Sub
        Public Sub KeyFields(ByVal str_KeyFields As String)
            Dim Num, Num2 As Integer
            Dim StartPos As Integer
            Dim StrPart As String
            Dim Index As Integer

            m_KeyFields = str_KeyFields
            str_KeyFields += "+"
            StartPos = 1
            Index = 1

            For Num = 1 To Len(str_KeyFields)
                If Mid(str_KeyFields, Num, 1) = "+" Then
                    StrPart = Mid(str_KeyFields, StartPos, Num - StartPos)
                    Col_KeyFields.Add(StrPart)
                    For Num2 = 1 To Col_ControlName.Count
                        If UCase(Col_ControlName(Num2).Name) = UCase(Col_KeyFields(Index)) Then
                            Col_KeyFields.Add(Col_ControlIndex(Num2))
                            Exit For
                        End If
                    Next Num2
                    StartPos = Num + 1
                    Index = Index + 2
                End If
            Next Num

        End Sub
        Public Sub SetLink(ByVal str_MasterFields As String, ByVal str_DetailFields As String)
            Dim Num, Num2 As Integer
            Dim StartPos As Integer
            Dim StrPart As String
            Dim Index As Integer

            str_MasterFields += "+"
            str_DetailFields += "+"
            StartPos = 1
            Index = 1

            For Num = 1 To Len(str_MasterFields)
                If Mid(str_MasterFields, Num, 1) = "+" Then
                    StrPart = Mid(str_MasterFields, StartPos, Num - StartPos)
                    Col_MasterFields.Add(StrPart)
                    StartPos = Num + 1
                End If
            Next Num


            StartPos = 1
            Index = 1
            For Num = 1 To Len(str_DetailFields)
                If Mid(str_DetailFields, Num, 1) = "+" Then
                    StrPart = Mid(str_DetailFields, StartPos, Num - StartPos)
                    Col_DetailFields.Add(StrPart)
                    StartPos = Num + 1
                    Index = Index + 1
                End If
            Next Num

            For Num = 0 To MyGrid.Columns.Count - 1
                For Num2 = 1 To Col_DetailFields.Count
                    'If UCase(MyGrid.Columns(Num).DataField) = UCase(Col_DetailFields(Num2)) Then  // maa 3
                    'MyGrid.Columns(Num).Visible = False
                    'End If
                Next Num2
            Next Num

        End Sub
        Private Sub PopulateGrid(ByRef dm_Grid As DataGridView, ByRef dm_MasterTable As DataTable, ByRef dm_DetailsTable As DataTable)
            Dim cFilter As String = ""
            Dim Num As Byte
            Dim X As Byte
            Dim X_ As Byte
            Dim oRecSet As New DataTable()
            Dim Pos As Byte


            For Num = 1 To Col_MasterFields.Count
                cFilter += Col_DetailFields(Num) + " = '" + dm_MasterTable.Rows(bm.Position)(Col_MasterFields(Num)) + "'"
                If Num <> Col_MasterFields.Count Then
                    cFilter += " and "
                End If
            Next Num

            'dm_DetailsTable.Select(cFilter) '// maa new
            dm_Grid.BeginEdit(True)
            bs_dt.Filter = cFilter

            For X = 1 To Col_GridKeyValue.Count / 6
                X_ = (X - 1) * 6 + 1
                da_rec.SelectCommand = New OleDb.OleDbCommand("Select * from " + Col_GridKeyValue(X_), CN)
                da_rec.Fill(ds, Col_GridKeyValue(X_))
                oRecSet = ds.Tables(Col_GridKeyValue(X_))

                bm_dt.Position = 0
                For I2 As Integer = 0 To dm_DetailsTable.Rows.Count - 1

                    bm_st.Position = 0
                    'oRecSet.Find(Col_GridKeyValue(X_ + 1) + " = '" + dm_DetailsTable.Fields(Col_GridKeyValue(X_ + 2)).Value + "'") // maa
                    If bm_st.Position < bm_st.Count Then
                        'dm_DetailsTable.Fields(Col_GridKeyValue(X_ + 4)).Value = oRecSet(Col_GridKeyValue(X_ + 3)).Value() // maavip
                    End If
                    bm_dt.Position += 1

                Next ' loop
                oRecSet.Dispose()
            Next X

            If FlipState = True Then
                For Pos = 0 To dm_Grid.Columns.Count - 1
                    dm_Grid.Columns(Pos).HeaderText = Col_GridFields(Pos + 1)
                Next
            End If

            'dm_Grid.DataSource = dm_DetailsTable
            dm_Grid.AllowUserToAddRows = True
            dm_Grid.AllowUserToDeleteRows = True
            dm_Grid.AllowUserToOrderColumns = True
            dm_Grid.AllowUserToResizeColumns = True
            dm_Grid.AllowUserToResizeRows = True
        End Sub
        Public Sub KeyLeaveField(ByRef dm_MasterTable As DataTable, ByVal str_KeyLeaveField As String, Optional ByVal n_ZeroPad As Byte = 0)
            Dim Num As Byte

            m_KeyLeaveField = str_KeyLeaveField
            For Num = 1 To Col_KeyFields.Count / 2
                If UCase(Col_KeyFields(Num * 2 - 1)) = UCase(m_KeyLeaveField) Then
                    KeyLeavePos = Col_KeyFields(Num * 2)
                    Exit For
                End If
            Next Num

            Key_ZeroPad = n_ZeroPad
            AddHandler MyForm.Controls(KeyLeavePos).Leave, AddressOf MyTextBox0_Leave

        End Sub
        Public Sub RequiredFields(ByVal str_RequiredFields As String, Optional ByVal b_ShowMsgBox As Boolean = True, Optional ByVal str_Msg As String = "Uncomplete Entries")
            Dim Num, Num2 As Integer
            Dim StartPos As Integer
            Dim StrPart As String


            RequiredFields_Msg = str_Msg
            RequiredFields_ShowMsg = b_ShowMsgBox

            str_RequiredFields += "+"
            StartPos = 1

            For Num = 1 To Len(str_RequiredFields)
                If Mid(str_RequiredFields, Num, 1) = "+" Then
                    StrPart = Mid(str_RequiredFields, StartPos, Num - StartPos)
                    For Num2 = 1 To Col_ControlName.Count
                        If UCase(Col_ControlName(Num2).Name) = UCase(StrPart) Then
                            Col_RequiredFields.Add(Col_ControlIndex(Num2))
                            Exit For
                        End If
                    Next Num2
                    StartPos = Num + 1
                End If
            Next Num

        End Sub
        Public Function ValidateForm(ByRef dm_Form As System.Windows.Forms.Form) As Boolean
            Dim Num As Byte

            For Num = 1 To Col_RequiredFields.Count
                If dm_Form.Controls(Col_RequiredFields(Num)).Text = "" Then
                    dm_Form.Controls(Col_RequiredFields(Num)).Focus()
                    If RequiredFields_ShowMsg Then
                        MsgBox(RequiredFields_Msg, MsgBoxStyle.OkOnly)
                    End If
                    Return False
                    Exit For
                End If
            Next Num
            Return True

        End Function
        Public Sub Search(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_Grid As DataGridView = Nothing, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            Static SearchFlag As Integer = 1
            Dim cFilter As String = ""
            Dim Num As Byte
            Dim ControlText As String
            Dim myChechBox As System.Windows.Forms.CheckBox
            Dim myRadioButton As System.Windows.Forms.RadioButton

            SearchFlag += 1

            If SearchFlag Mod 2 = 0 Then
                'MyForm.Text += " Filter Mode"
                Me.ClearData(dm_Form, dm_DetailsTable, dm_Grid)
            Else
                'MyForm.Text = Mid(MyForm.Text, 1, Len(MyForm.Text) - 12)
                For Num = 1 To Col_ControlIndex.Count()
                    If Col_ControlIndex(Num) <> KeyLeavePos And UCase(Left(Col_ControlName(Num).Name, 1)) <> "X" Then
                        If TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "TextBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ComboBox" Or TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "ListBox" Then
                            ControlText = dm_Form.Controls(Col_ControlIndex(Num)).Text
                            If ControlText <> "" Then
                                cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " like '*" + ControlText + "*' and "
                            End If
                        ElseIf TypeName(dm_Form.Controls(Col_ControlIndex(Num))) = "CheckBox" Then
                            myChechBox = dm_Form.Controls(Col_ControlIndex(Num))
                            ControlText = IIf(myChechBox.Checked, "True", "False")
                            If ControlText <> "" Then
                                cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " = " + ControlText + " and "
                            End If
                        Else 'RadioButton
                            myRadioButton = dm_Form.Controls(Col_ControlIndex(Num))
                            ControlText = IIf(myRadioButton.Checked, "True", "False")
                            If ControlText <> "" Then
                                cFilter += dm_Form.Controls(Col_ControlIndex(Num)).Name + " = " + ControlText + " and "
                            End If
                        End If

                    End If
                Next Num
                If cFilter <> "" Then
                    cFilter = Mid(cFilter, 1, cFilter.Length - 4) ' remove 'and' from filter tail
                End If
                dm_MasterTable.Select(cFilter)
                Me.PopulateForm(dm_Form, dm_MasterTable, dm_Grid, dm_DetailsTable)
            End If
        End Sub
        Private Sub dm_Grid_OnAddNew(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            'sender.DataSource.rows.add(New String() {"v1", "v2"}) ' maavip
            ' ds.DataTable.AddRow(1, "John Doe", true)
            If m_DetailFlagField <> "" Then
                sender.DataSource.Fields(m_DetailFlagField).value = m_FlagValue
            End If
            For Num = 1 To Col_KeyText.Count
                'sender.DataSource.Fields(Col_KeyFields((Num * 2) - 1)).Value() = Col_KeyText(Num) // maavip
            Next Num
            'sender.DataSource.Refresh() 'maa
        End Sub
        Private Sub dm_Grid_AfterColEdit(ByVal sender As Object, ByVal e As DataGridView) 'AxMSDataGridLib.DDataGridEvents_AfterColEditEvent // maa
            Dim Pos1 As Integer = -1
            Dim Pos2 As Integer = -1
            Dim Pos3 As Integer = -1

            On Error Resume Next ' Keep it
            Pos1 = Col_GridKeyValue(6)
            Pos2 = Col_GridKeyValue(13)
            Pos3 = Col_GridKeyValue(20)

            'If e.colIndex = Pos1 Then
            'MyGrid.Columns(Pos1).Value = ZeroPad(MyGrid.Columns(Pos1).Value, Col_GridKeyValue(7))
            'MyGrid.Columns(Pos1 + 1).Value = GetValue(Col_GridKeyValue(1), Col_GridKeyValue(2), MyGrid.Columns(Pos1).Value, Col_GridKeyValue(4))
            'ElseIf e.colIndex = Pos2 Then
            'MyGrid.Columns(Pos2).Value = ZeroPad(MyGrid.Columns(Pos2).Value, Col_GridKeyValue(14))
            'MyGrid.Columns(Pos2 + 1).Value = GetValue(Col_GridKeyValue(8), Col_GridKeyValue(9), MyGrid.Columns(Pos2).Value, Col_GridKeyValue(11))
            'ElseIf e.colIndex = Pos3 Then
            'MyGrid.Columns(Pos3).Value = ZeroPad(MyGrid.Columns(Pos3).Value, Col_GridKeyValue(21))
            'MyGrid.Columns(Pos3 + 1).Value = GetValue(Col_GridKeyValue(14), Col_GridKeyValue(15), MyGrid.Columns(Pos3).Value, Col_GridKeyValue(17))
            'End If

        End Sub
        Private Sub dm_Grid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) '// Handles datagridview1.keydown  // maa  as datagridview.keydown
            Dim Pos1 As Integer = -1
            Dim Pos2 As Integer = -1
            Dim Pos3 As Integer = -1
            Dim oHelpForm As New DataHelpForm()
            Dim Num As Byte

            On Error Resume Next ' Keep it
            If e.KeyCode = Keys.F1 Then
                Pos1 = Col_GridKeyValue(6)
                Pos2 = Col_GridKeyValue(13)
                Pos3 = Col_GridKeyValue(20)

                If sender.Col = Pos1 Then
                    Num = 1
                    HelpFile = Col_GridKeyValue(Num)
                    HelpID = Col_GridKeyValue(Num + 1)
                    HelpName = Col_GridKeyValue(Num + 3)
                    HelpIdSender = sender.Name + "_" + sender.Columns(Pos1).DataField
                    oHelpForm.Show()
                ElseIf sender.Col = Pos2 Then
                    Num = 8
                    HelpFile = Col_GridKeyValue(Num)
                    HelpID = Col_GridKeyValue(Num + 1)
                    HelpName = Col_GridKeyValue(Num + 3)
                    HelpIdSender = sender.Name + "_" + sender.Columns(Pos2).DataField
                    oHelpForm.Show()
                ElseIf sender.Col = Pos3 Then
                    Num = 15
                    HelpFile = Col_GridKeyValue(Num)
                    HelpID = Col_GridKeyValue(Num + 1)
                    HelpName = Col_GridKeyValue(Num + 3)
                    HelpIdSender = sender.Name + "_" + sender.Columns(Pos3).DataField
                    oHelpForm.Show()
                End If
            End If
        End Sub
        Private Sub MyTextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            Num = 1
            If sender.text <> "" Then
                sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
                MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
            End If
        End Sub
        Private Sub MyTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F1 Then
                Dim Num As Byte = 1
                Dim oHelpForm As New DataHelpForm()
                HelpFile = Col_KeyValue(Num + 2)
                HelpID = Col_KeyValue(Num + 3)
                HelpName = Col_KeyValue(Num + 5)
                HelpIdSender = sender.Name
                oHelpForm.Show()
            End If
        End Sub

        Private Sub MyTextBox2_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            Num = 9
            If sender.text <> "" Then
                sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
                MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
            End If
        End Sub
        Private Sub MyTextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F1 Then
                Dim Num As Byte = 9
                Dim oHelpForm As New DataHelpForm()
                HelpFile = Col_KeyValue(Num + 2)
                HelpID = Col_KeyValue(Num + 3)
                HelpName = Col_KeyValue(Num + 5)
                HelpIdSender = sender.Name
                oHelpForm.Show()
            End If
        End Sub

        Private Sub MyTextBox3_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim Num As Byte

            Num = 17
            If sender.text <> "" Then
                sender.Text = ZeroPad(sender.Text, Col_KeyValue(Num + 6))
                MyForm.Controls(Col_KeyValue(Num)).Text = GetValue(Col_KeyValue(Num + 2), Col_KeyValue(Num + 3), sender.Text, Col_KeyValue(Num + 5))
            End If
        End Sub
        Private Sub MyTextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F1 Then
                Dim Num As Byte = 17
                Dim oHelpForm As New DataHelpForm()
                HelpFile = Col_KeyValue(Num + 2)
                HelpID = Col_KeyValue(Num + 3)
                HelpName = Col_KeyValue(Num + 5)
                HelpIdSender = sender.Name
                oHelpForm.Show()
            End If
        End Sub
        Public Sub NavigationButtons(ByVal dm_First As String, ByVal dm_Previous As String, ByVal dm_Next As String, ByVal dm_Last As String)
            Dim MyButton As New System.Windows.Forms.Control()
            Dim Num As Byte = 0
            Dim cButton As String
            On Error Resume Next

            For Each MyButton In MyForm.Controls
                If UCase(MyButton.Name) = UCase(dm_First) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(0))
                    AddHandler MyForm.Controls(Num).Click, AddressOf FirstButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf FirstButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf FirstButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf FirstButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf FirstButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf FirstButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf FirstButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_Previous) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(3))
                    AddHandler MyForm.Controls(Num).Click, AddressOf PreviousButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf PreviousButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf PreviousButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf PreviousButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf PreviousButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf PreviousButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf PreviousButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_Next) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(6))
                    AddHandler MyForm.Controls(Num).Click, AddressOf NextButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf NextButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf NextButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf NextButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf NextButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf NextButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf NextButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_Last) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(9))
                    AddHandler MyForm.Controls(Num).Click, AddressOf LastButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf LastButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf LastButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf LastButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf LastButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf LastButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf LastButton_Enter
                End If
                Num += 1
            Next MyButton

        End Sub
        Public Sub ManipulationButtons(ByVal dm_Save As String, ByVal dm_New As String, ByVal dm_Delete As String, ByVal dm_Close As String, Optional ByVal dm_Search As String = Nothing)
            Dim MyButton As New System.Windows.Forms.Control()
            Dim Num As Byte = 0
            On Error Resume Next

            For Each MyButton In MyForm.Controls
                If UCase(MyButton.Name) = UCase(dm_Save) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(12))
                    AddHandler MyForm.Controls(Num).Click, AddressOf SaveButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf SaveButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf SaveButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf SaveButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf SaveButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf SaveButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf SaveButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_New) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(15))
                    AddHandler MyForm.Controls(Num).Click, AddressOf NewButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf NewButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf NewButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf NewButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf NewButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf NewButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf NewButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_Delete) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(18))
                    AddHandler MyForm.Controls(Num).Click, AddressOf DeleteButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf DeleteButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf DeleteButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf DeleteButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf DeleteButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf DeleteButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf DeleteButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_Close) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(21))
                    AddHandler MyForm.Controls(Num).Click, AddressOf CloseButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf CloseButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf CloseButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf CloseButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf CloseButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf CloseButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf CloseButton_Enter
                ElseIf UCase(MyButton.Name) = UCase(dm_Search) Then
                    Dim MyImageButton As Button
                    MyImageButton = CType(MyButton, Button)
                    MyImageButton.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(24))
                    AddHandler MyForm.Controls(Num).Click, AddressOf SearchButton_Click
                    AddHandler MyForm.Controls(Num).MouseEnter, AddressOf SearchButton_MouseEnter
                    AddHandler MyForm.Controls(Num).MouseLeave, AddressOf SearchButton_MouseLeave
                    AddHandler MyForm.Controls(Num).MouseDown, AddressOf SearchButton_MouseDown
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf SearchButton_KeyDown
                    AddHandler MyForm.Controls(Num).Leave, AddressOf SearchButton_Leave
                    AddHandler MyForm.Controls(Num).Enter, AddressOf SearchButton_Enter
                End If
                Num += 1
            Next MyButton
        End Sub
        Public Sub PrepareImageButtons(ByVal ImagesArray() As String, ByVal ImageFullPath As String, ByVal Modtion As Boolean)
            HasImage = True
            ImageMotion = Modtion
            ImagePath = ImageFullPath
            ImageButtons = ImagesArray
        End Sub
        Private Sub FirstButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            ''''''''''First Button Handles
            GoFirst(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub FirstButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(0))
            End If
        End Sub
        Private Sub FirstButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(1))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub FirstButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(2))
            End If
        End Sub
        Private Sub FirstButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(0))
            End If
        End Sub
        Private Sub FirstButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(1))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub FirstButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(2))
            End If
        End Sub
        Private Sub PreviousButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Previous Button Handles
            GoPrevious(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub PreviousButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(3))
            End If
        End Sub
        Private Sub PreviousButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(4))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub PreviousButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(5))
            End If
        End Sub
        Private Sub PreviousButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(3))
            End If
        End Sub
        Private Sub PreviousButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(4))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub PreviousButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(5))
            End If
        End Sub

        ''''''''''Next Button Handles
        Private Sub NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            GoNext(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub NextButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(6))
            End If
        End Sub
        Private Sub NextButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(7))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub NextButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(8))
            End If
        End Sub
        Private Sub NextButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(6))
            End If
        End Sub
        Private Sub NextButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(7))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub NextButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(8))
            End If
        End Sub
        Private Sub LastButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Last Button Hadles
            GoLast(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub LastButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(9))
            End If
        End Sub
        Private Sub LastButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(10))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub LastButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(11))
            End If
        End Sub
        Private Sub LastButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(9))
            End If
        End Sub
        Private Sub LastButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(10))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub LastButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(11))
            End If
        End Sub
        Private Sub SaveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Save Button Handles
            SaveData(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub SaveButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(12))
            End If
        End Sub
        Private Sub SaveButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(13))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub SaveButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(14))
            End If
        End Sub
        Private Sub SaveButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(12))
            End If
        End Sub
        Private Sub SaveButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(13))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub SaveButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(14))
            End If
        End Sub
        Private Sub NewButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'New Button Handles
            NewRecord(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub NewButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(15))
            End If
        End Sub
        Private Sub NewButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(16))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub NewButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(17))
            End If
        End Sub
        Private Sub NewButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(15))
            End If
        End Sub
        Private Sub NewButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(16))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub NewButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(17))
            End If
        End Sub
        Private Sub DeleteButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Delete Button Handles
            DeleteRecord(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub DeleteButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(18))
            End If
        End Sub
        Private Sub DeleteButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(19))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub DeleteButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(20))
            End If
        End Sub
        Private Sub DeleteButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(18))
            End If
        End Sub
        Private Sub DeleteButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(19))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub DeleteButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(20))
            End If
        End Sub
        Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Close Button Handles
            CloseForm(MyForm, oMaster, oDetails)
        End Sub
        Private Sub CloseButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(21))
            End If
        End Sub
        Private Sub CloseButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(22))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub CloseButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(23))
            End If
        End Sub
        Private Sub CloseButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(21))
            End If
        End Sub
        Private Sub CloseButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(22))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub CloseButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(23))
            End If
        End Sub
        Private Sub SearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Searrch Button Handles
            Search(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Private Sub SearchButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(24))
            End If
        End Sub
        Private Sub SearchButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(25))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub SearchButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(26))
            End If
        End Sub
        Private Sub SearchButton_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(24))
            End If
        End Sub
        Private Sub SearchButton_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(25))
            End If
            If ImageMotion = True Then
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
                sender.Image.RotateFlip(RotateFlipType.Rotate180FlipNone)
                sender.Refresh()
                System.Threading.Thread.Sleep(150)
            End If
        End Sub
        Private Sub SearchButton_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Space And HasImage Then
                sender.Image = System.Drawing.Image.FromFile(ImagePath + ImageButtons(26))
            End If
        End Sub
        Private Sub CloseForm(ByRef dm_Form As System.Windows.Forms.Form, ByRef dm_MasterTable As DataTable, Optional ByRef dm_DetailsTable As DataTable = Nothing)
            On Error Resume Next

            dm_DetailsTable.Dispose()
            If HasGrid Then
                dm_DetailsTable.Dispose()
            End If
            dm_Form.Close()
        End Sub

        Private Sub MyTextBox0_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            If Key_ZeroPad <> 0 Then
                sender.Text = ZeroPad(sender.Text, Key_ZeroPad)
            End If
            KeyLeave(MyForm, oMaster, MyGrid, oDetails)
        End Sub
        Public Sub SpecialChars(ByVal str_Chars As String)
            m_SpecialChars = str_Chars
        End Sub
        Public Sub SpecialCharsFields(ByVal ParamArray str_SpecialFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_SpecialFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.KeyPress, AddressOf SpecialCharsFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Special Characters (" + m_SpecialChars + ")"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub SpecialCharsFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)

            If (m_SpecialChars <> "" And (InStr(1, m_SpecialChars, Chr(KeyAscii)) > 0)) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub
        Public Sub AlphaNumericFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.KeyPress, AddressOf AlphaNumericFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Alphabetic & Numeric"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub AlphaNumericFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub
        Public Sub AlphabeticFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.KeyPress, AddressOf AlphabeticFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Alphabetic"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub AlphabeticFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Public Sub DecimalFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf DecimalFields_Leave
                        AddHandler txt_control.KeyPress, AddressOf DecimalFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Integer & Decimal" + IIf(m_DecimalPlaces <> 0, " with " + m_DecimalPlaces.ToString + " Decimal Digits", "")
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub DecimalFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub
        Private Sub DecimalFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Static Flag As Byte = 0
            Dim TextBoxColor As System.Drawing.Color
            On Error Resume Next
            Flag += 1
            If Flag = 1 Then
                TextBoxColor = sender.ForeColor
            End If
            Dim j As TextBox

            If sender.Text <> "" And Not IsNumeric(sender.Text) Then
                sender.ForeColor = System.Drawing.Color.Red
            Else
                sender.ForeColor = TextBoxColor
                Dim Num As Decimal = CDec(sender.Text)
                If m_DecimalPlaces <> 0 Then
                    Num = Format(Num, "############." + New String("0", m_DecimalPlaces))
                End If
                sender.Text = Num
            End If
        End Sub
        Public Sub NumericFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf NumericFields_Leave
                        AddHandler txt_control.KeyPress, AddressOf NumericFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Numeric"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub NumericFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub
        Private Sub NumericFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Static Flag As Byte = 0
            Dim TextBoxColor As System.Drawing.Color
            On Error Resume Next
            Flag += 1
            If Flag = 1 Then
                TextBoxColor = sender.ForeColor
            End If

            If sender.Text <> "" And Not IsNumeric(sender.Text) Then
                sender.ForeColor = System.Drawing.Color.Red
            Else
                sender.ForeColor = TextBoxColor
            End If
        End Sub
        Public Sub DateFields(ByVal ParamArray str_DateFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_DateFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf DateFields_Leave
                        AddHandler txt_control.KeyPress, AddressOf DateFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Date"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub DateFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 47) Or (KeyAscii = 92) Or (KeyAscii = 8) Or (KeyAscii = 13) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub
        Private Sub DateFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Static Flag As Byte = 0
            Dim TextBoxColor As System.Drawing.Color
            On Error Resume Next
            Flag += 1
            If Flag = 1 Then
                TextBoxColor = sender.ForeColor
            End If

            If sender.Text <> "" And Not IsDate(sender.Text) Then
                sender.ForeColor = System.Drawing.Color.Red
            Else
                sender.ForeColor = TextBoxColor
            End If
        End Sub
        Public Sub UpperCaseFields(ByVal ParamArray str_UpperCaseFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_UpperCaseFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf UpperCaseFields_Leave
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub UpperCaseFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            sender.text = UCase(sender.text)
        End Sub
        Public Sub LowerCaseFields(ByVal ParamArray str_LowerCaseFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_LowerCaseFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf LowerCaseFields_Leave
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub LowerCaseFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            sender.text = LCase(sender.text)
        End Sub
        Public Sub FirstCharOnlyFields(ByVal ParamArray str_FirstCharOnlyFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_FirstCharOnlyFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf FirstCharOnlyFields_Leave
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub FirstCharOnlyFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            sender.Text = UCase(Left(sender.Text, 1)) + Right(sender.Text, Len(sender.Text) - 1)
        End Sub
        Public Sub FirstCharOfWordsFields(ByVal ParamArray str_FirstCharOfWordsFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_FirstCharOfWordsFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf FirstCharOfWordsFields_Leave
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub FirstCharOfWordsFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim MyChar As String
            Dim PrevChar As String = ""
            Dim newValue As String = ""
            Dim X As Integer

            sender.Text = UCase(Left(sender.Text, 1)) + Right(sender.Text, Len(sender.Text) - 1)

            For X = 1 To Len(sender.Text)
                MyChar = Mid(sender.Text, X, 1)
                If PrevChar = " " Then
                    newValue = newValue + UCase(MyChar)
                Else
                    newValue = newValue + MyChar
                End If
                PrevChar = MyChar
            Next X
            sender.Text = newValue
        End Sub
        Public Sub EnableReturnKey(ByVal Mode As Boolean)
            Dim MyControl As System.Windows.Forms.Control

            If Mode = True Then
                For Each MyControl In MyForm.Controls
                    AddHandler MyControl.KeyPress, AddressOf EnableTab_KeyPress
                Next MyControl
            End If

        End Sub
        Public Sub DecimalPlaces(ByVal n_DecimalPlaces As Byte)
            m_DecimalPlaces = n_DecimalPlaces
        End Sub
        Private Sub EnableTab_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            If e.KeyChar = Chr(13) Then
                System.Windows.Forms.SendKeys.Send("{Tab}")
            End If
        End Sub
        Private Sub PrepareHelp()
            Dim Num As Integer = 0
            Dim MyControl As System.Windows.Forms.Control

            For Each MyControl In MyForm.Controls
                If TypeName(MyControl) <> "Label" Then
                    AddHandler MyForm.Controls(Num).KeyDown, AddressOf ControlHelp_KeyDown
                End If
                Num += 1
            Next MyControl

            If HasGrid Then
                AddHandler MyGrid.KeyDown, AddressOf ColumnHelp_KeyDown
            End If

        End Sub
        Private Sub ControlHelp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.F12 Then
                Dim oHelpFileForm As New HelpForm()
                Dim Num As Byte
                HelpFieldProperty(0) = sender.FindForm.Tag
                HelpFieldProperty(1) = sender.Name
                HelpFieldProperty(2) = ""
                For Num = 0 To Col_FieldsType.GetLength(0) - 1
                    If UCase(Col_FieldsType(Num, 0)) = UCase(sender.Name) Then
                        HelpFieldProperty(2) = Col_FieldsType(Num, 1)
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(3) = ""
                If TypeName(sender) = "TextBox" Then
                    Dim MyTextBox As TextBox = CType(sender, TextBox)
                    HelpFieldProperty(3) = IIf(MyTextBox.MaxLength <> 0, MyTextBox.MaxLength, "")
                End If
                HelpFieldProperty(4) = ""
                For Num = 1 To Col_RequiredFields.Count
                    If UCase(MyForm.Controls(Col_RequiredFields(Num)).Name) = UCase(sender.Name) Then
                        HelpFieldProperty(4) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(5) = ""
                For Num = 5 To Col_KeyValue.Count Step 8
                    If UCase(Col_KeyValue(Num)) = UCase(sender.Name) Then
                        HelpFieldProperty(5) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(6) = TypeName(sender)
                oHelpFileForm.Show()
            End If
        End Sub
        Private Sub ColumnHelp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) '// maa dataview keydown 
            If e.KeyCode = Keys.F12 Then
                Dim oHelpFileForm As New HelpForm()
                Dim Num As Byte
                HelpFieldProperty(0) = sender.FindForm.Tag
                HelpFieldProperty(1) = sender.Name + "_" + sender.Columns(sender.Col).DataField
                HelpFieldProperty(2) = ""
                For Num = 0 To Col_FieldsType.GetLength(0) - 1
                    If UCase(Col_FieldsType(Num, 0)) = UCase(sender.Name) Then
                        HelpFieldProperty(2) = Col_FieldsType(Num, 1)
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(3) = ""
                If TypeName(sender) = "TextBox" Then
                    Dim MyTextBox As TextBox = CType(sender, TextBox)
                    HelpFieldProperty(3) = IIf(MyTextBox.MaxLength <> 0, MyTextBox.MaxLength, "")
                End If
                HelpFieldProperty(4) = ""
                For Num = 1 To Col_RequiredFields.Count
                    If UCase(MyForm.Controls(Col_RequiredFields(Num)).Name) = UCase(sender.Name) Then
                        HelpFieldProperty(4) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(5) = ""
                For Num = 5 To Col_KeyValue.Count Step 8
                    If UCase(Col_KeyValue(Num)) = UCase(sender.Name) Then
                        HelpFieldProperty(5) = "Yes"
                        Exit For
                    End If
                Next Num
                HelpFieldProperty(6) = TypeName(sender)

                oHelpFileForm.Show()
            End If
        End Sub
        Public Sub Right2Left(ByVal Mode As Boolean)
            Dim MyControl As New System.Windows.Forms.Control()
            Right2LeftState = IIf(Mode = True, 1, 0)
            For Each MyControl In MyForm.Controls
                If TypeName(MyControl) = "TextBox" Then
                    If Mode = True Then
                        MyControl.RightToLeft = RightToLeft.Yes
                    Else
                        MyControl.RightToLeft = RightToLeft.No
                    End If
                End If
            Next MyControl

            MyForm.RightToLeft = Right2LeftState
            If HasGrid Then
                MyGrid.RightToLeft = Right2LeftState
            End If
        End Sub
        Private Sub ReadInitialValues()
            Dim MyFile As String
            Dim MyFullPathFile As String
            Dim oInitValues As New InitValues()
            Dim FileNum As Byte

            Dim WinSys As String
            Dim WshShell As New Object()
            On Error Resume Next

            WshShell = CreateObject("WScript.Shell")
            WinSys = WshShell.SpecialFolders("Fonts")
            WinSys = Mid(WinSys, 1, Len(WinSys) - 5)
            WinSys += "System32\"
            FileNum = FreeFile()
            MyFullPathFile = WinSys + "DCDM30_Lang.dll"
            MyFile = Dir(WinSys + "DCDM30_Lang.dll")

            FileOpen(1, MyFullPathFile, OpenMode.Random, OpenAccess.ReadWrite, OpenShare.Shared, 1000)
            If UCase(MyFile) = "DCDM30_LANG.DLL" Then
                FileGet(FileNum, oInitValues, 1)
                aInitValues(0) = oInitValues.Help_Caption.Trim + " "
                aInitValues(1) = oInitValues.Help_DataEntryType.Trim + " "
                aInitValues(2) = oInitValues.Help_MaxLenght.Trim + " "
                aInitValues(3) = oInitValues.Help_Required.Trim + " "
                aInitValues(4) = oInitValues.Help_HasDataHelp.Trim + " "
                aInitValues(5) = oInitValues.Help_Const_NotDefined.Trim + " "
                aInitValues(6) = oInitValues.Help_Const_Characters.Trim + " "
                aInitValues(7) = oInitValues.Help_Const_Description.Trim + " "
                aInitValues(8) = oInitValues.Help_Const_Yes.Trim + " "
                aInitValues(9) = oInitValues.Help_Const_No.Trim + " "
                ' Keep 5 room for Future add
                aInitValues(15) = oInitValues.DataHelp_Caption.Trim + " "
                aInitValues(16) = oInitValues.DataHelp_Id.Trim + " "
                aInitValues(17) = oInitValues.DataHelp_Name.Trim + " "
                ' Keep 2 room for Future add
                aInitValues(20) = oInitValues.Delete_Message.Trim + " "
                aInitValues(20) = "Are you sure you want to delete this record"
            End If
        End Sub
        Public Sub TranslateForm(ByRef dm_Form As System.Windows.Forms.Form, ByVal dm_Language As Byte)
            Dim MyControl As System.Windows.Forms.Control
            Dim oLang As New DataTable()

            FlipState = True
            'oLang.Open("MultiLanguage")
            da_rec.SelectCommand = New OleDb.OleDbCommand("Select * from MultiLanguage", CN)
            da_rec.Fill(ds, "MultiLanguage")
            oLang = ds.Tables("MultiLanguage")

            oLang.Select("Tag = '" + dm_Form.Name + "' and Id = '" + dm_Form.Name + "'")
            If bm_ng.Position < bm_ng.Count Then
                ' dm_Form.Text = oLang.Fields("Language" + dm_Language.ToString).Value // maa
            End If

            For Each MyControl In dm_Form.Controls
                If TypeName(MyControl) = "Label" Then
                    oLang.Select("Tag = '" + dm_Form.Name + "' and Id = '" + MyControl.Name + "'")
                    If bm_ng.Position < bm_ng.Count Then
                        'MyControl.Text = oLang.Fields("Language" + dm_Language.ToString).Value // maa
                    End If
                End If
            Next MyControl

            If HasGrid Then
                Dim Num As Byte
                For Num = 0 To MyGrid.Columns.Count - 1
                    'oLang.Filter = "Tag = '" + dm_Form.Name + "' and Id = '" + MyGrid.Name + "_" + MyGrid.Columns(Num).DataField + "'" // maa
                    ' oLang.Select("Tag = '" + dm_Form.Name + "' and Id = '" + MyGrid.Name + "_" + MyGrid.Columns(Num).DataField + "'"  // maa
                    If bm_ng.Position < bm_ng.Count Then
                        'Col_GridFields.Add(oLang.Fields("Language" + dm_Language.ToString).Value) // maa
                    Else
                        Col_GridFields.Add(MyGrid.Columns(Num).HeaderText)
                    End If
                Next Num
            End If

        End Sub
        Public Sub FlipForm(ByRef dm_Form As System.Windows.Forms.Form)
            Dim MyControl As System.Windows.Forms.Control

            For Each MyControl In dm_Form.Controls
                MyControl.Left = dm_Form.Width - (MyControl.Left + MyControl.Width)
            Next MyControl
        End Sub


    End Class
End Namespace