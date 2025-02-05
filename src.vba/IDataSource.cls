'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module    : IDataSource
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 2025/01/12
' Purpose   : Data Source Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit



'***************************************************************************************************************************************************************
'    Public Functions and Subs
'***************************************************************************************************************************************************************

Public Function GetValue(ByVal dataAccess As CRecordableDA, ByVal strFieldName As String, ByRef oResult As CResult) As CField
End Function

Public Function GetRecord(ByVal dataAccess As CRecordableDA, ByRef oResult As CResult) As CRecord
End Function

Public Function GetRecordList(ByVal dataAccess As CRecordableDA, ByRef oResult As CResult) As CRecordList
End Function

Public Function AddValues(ByVal dataAccess As CRecordableDA) As Variant
End Function

Public Function UpdateValues(ByVal dataAccess As CRecordableDA) As Variant
End Function

Public Function DeleteValues(ByVal dataAccess As CRecordableDA) As Variant
End Function
