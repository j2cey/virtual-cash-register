'Build 000
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Module    : mEnums
'   Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
'   Created   : 2025/01/08
'   Purpose   : All the User defined Enums
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum enuFormStatus
    Default = 0
    Active = 1
End Enum

Public Enum enuResultCode
    Default = 0
    Active = 1
End Enum

Public Enum enuDataSourceClass
  fileSource = 1
  sheetSource = 2
  databaseSource = 3
End Enum

Public Enum enuDatabaseClass
  None = 1
  access2007 = 2
  sqlserver2014 = 3
End Enum

Public Enum enuHoverAction
    changeSize = 1
    changeBackground = 2
End Enum