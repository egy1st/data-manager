Imports System.Data

Module Module1
    Public Structure InitValues
        <VBFixedString(50)> Public Help_Caption As String
        <VBFixedString(50)> Public Help_DataEntryType As String
        <VBFixedString(50)> Public Help_MaxLenght As String
        <VBFixedString(50)> Public Help_Required As String
        <VBFixedString(50)> Public Help_HasDataHelp As String
        <VBFixedString(30)> Public Help_Const_NotDefined As String
        <VBFixedString(30)> Public Help_Const_Characters As String
        <VBFixedString(30)> Public Help_Const_Description As String
        <VBFixedString(15)> Public Help_Const_Yes As String
        <VBFixedString(15)> Public Help_Const_No As String
        <VBFixedString(50)> Public DataHelp_Caption As String
        <VBFixedString(20)> Public DataHelp_Id As String
        <VBFixedString(20)> Public DataHelp_Name As String
        <VBFixedString(70)> Public Delete_Message As String
    End Structure

    Public CN As New OleDb.OleDbConnection
    Public HelpFile As String
    Public HelpID As String
    Public HelpName As String
    Public HelpRtnID As String
    Public HelpRtnName As String
    Public HelpFieldProperty(6) As String
    Public aInitValues(20) As String
    Public Right2LeftState As Byte = 0

End Module
