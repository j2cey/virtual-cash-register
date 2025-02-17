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
  noDataSource = 0
  fileSource = 1
  sheetSource = 2
  databaseSource = 3
End Enum

Public Enum enuDatabaseClass
  noDatabase = 0
  access2007 = 1
  sqlserver2014 = 2
End Enum

Public Enum enuHoverAction
    changeSize = 1
    changeBackground = 2
End Enum

Public Enum enuStoreFieldName
    doNotStore = 0
    toTheLeft = 1
    toTheRight = 2
End Enum