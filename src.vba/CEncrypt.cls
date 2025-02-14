'Build 000
'*****************************************************************************************************************************
'
' Authors:  Jude Parfait NGOM NZE, jud10parfait@gmail.com
'
' (c)2016-2017, all rights reserved to the authors
'
'*****************************************************************************************************************************

'-------------------------------------------------------------------------
' Module    : clsENCRYPT
' Author    : Jude Parfait NGOM NZE (jud10parfait@gmail.com)
' Created   : 02-08-2017
' Purpose   : Ecryption class
'-------------------------------------------------------------------------
Option Explicit

Option Compare Binary
Private Const MIN_ASC As Integer = 32
Private Const MAX_ASC As Integer = 126
Private Const NO_OF_CHARS As Integer = MAX_ASC - MIN_ASC + 1

Private Const C_ENCRYPT_KEY As String = "thismyturn"

Private m_sKey As String


Public Property Let Key(sKey As String)
    m_sKey = sKey
End Property

Public Property Get Key() As String
    Key = m_sKey
End Property


Private Function MoveAsc(ByVal a, ByVal mLvl)
    'Move the Asc value so it stays inside interval MIN_ASC and MAX_ASC
    mLvl = mLvl Mod NO_OF_CHARS
    a = a + mLvl
    If a < MIN_ASC Then
        a = a + NO_OF_CHARS
    ElseIf a > MAX_ASC Then
        a = a - NO_OF_CHARS
    End If
    MoveAsc = a
End Function

Public Function encrypt(ByVal s As String)
    Dim p, keyPos, c, E, K, chkSum
    If Me.Key = "" Then
        encrypt = s
        Exit Function
    End If
    For p = 1 To Len(s)
        If Asc(Mid(s, p, 1)) < MIN_ASC Or Asc(Mid(s, p, 1)) > MAX_ASC Then
            MsgBox "Char at position " & p & " is invalid!"
            Exit Function
        End If
    Next p
    For keyPos = 1 To Len(Me.Key)
        chkSum = chkSum + Asc(Mid(Me.Key, keyPos, 1)) * keyPos
    Next keyPos
    keyPos = 0
    For p = 1 To Len(s)
        c = Asc(Mid(s, p, 1))
        keyPos = keyPos + 1
        If keyPos > Len(Me.Key) Then keyPos = 1
        K = Asc(Mid(Me.Key, keyPos, 1))
        c = MoveAsc(c, K)
        c = MoveAsc(c, K * Len(Me.Key))
        c = MoveAsc(c, chkSum * K)
        c = MoveAsc(c, p * K)
        c = MoveAsc(c, Len(s) * p) 'This is only for getting new chars for different word lengths
        E = E & Chr(c)
    Next p
    
    E = Replace(E, Chr(34), "Dbl_Qte")
    E = Replace(E, Chr(39), "Sgl_Qte")
    encrypt = E
End Function

Public Function decrypt(ByVal s As String)
    Dim p, keyPos, c, d, K, chkSum
    
    s = Replace(s, "Dbl_Qte", Chr(34))
    s = Replace(s, "Sgl_Qte", Chr(39))
    
    If Me.Key = "" Then
        decrypt = s
        Exit Function
    End If
    For keyPos = 1 To Len(Me.Key)
        chkSum = chkSum + Asc(Mid(Me.Key, keyPos, 1)) * keyPos
    Next keyPos
    keyPos = 0
    For p = 1 To Len(s)
        c = Asc(Mid(s, p, 1))
        keyPos = keyPos + 1
        If keyPos > Len(Me.Key) Then keyPos = 1
        K = Asc(Mid(Me.Key, keyPos, 1))
        'Do MoveAsc in reverse order from encrypt, and with a minus sign this time(to unmove)
        c = MoveAsc(c, -(Len(s) * p))
        c = MoveAsc(c, -(p * K))
        c = MoveAsc(c, -(chkSum * K))
        c = MoveAsc(c, -(K * Len(Me.Key)))
        c = MoveAsc(c, -K)
        d = d & Chr(c)
    Next p
    decrypt = d
End Function

Private Sub Class_Initialize()
    Me.Key = C_ENCRYPT_KEY
End Sub