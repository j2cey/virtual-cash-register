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

Public Function GetValue(ByVal dataAccess As CRecordableDA, ByRef oFieldValue As CField) As CResult
End Function

Public Function GetRecord(ByVal dataAccess As CRecordableDA, ByRef oRecordValue As CRecord) As CResult
End Function

Public Function GetRecordList(ByVal dataAccess As CRecordableDA, ByVal oRecordTemplate As CRecord, ByRef oRecordListValue As CRecordList) As CResult
End Function

Public Function AddValues(ByVal dataAccess As CRecordableDA) As Variant
End Function

Public Function UpdateValues(ByVal dataAccess As CRecordableDA) As Variant
End Function

Public Function DeleteValues(ByVal dataAccess As CRecordableDA) As Variant
End Function
