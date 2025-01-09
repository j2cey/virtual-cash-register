Option Explicit

Public Enum ctlRole
    simplectl = 0
    mainId = 1
    mainctl = 2
    mainlist = 3
    searchcriteria = 4
    SaveBtn = 5
    delBtn = 6
    CancelBtn = 7
    SearchBtn = 8
    closeBtn = 9
End Enum

Public Enum valueType
    varval = 0
    textval = 1
    intval = 2
    longval = 3
    boolval = 4
    dateval = 5
    encryptedval = 6
End Enum

Public Enum ctlActionType
    SetValue = 0
    setEnabilityState = 1
    setVisibilityState = 2
    setLockState = 3
End Enum

Public Enum saveFormStatusType
    None = 0
    Add = 1
    Update = 2
    Delete = 3
    Show = 4
End Enum

Public Enum formStatusType
    Default = 0
    Active = 1
End Enum