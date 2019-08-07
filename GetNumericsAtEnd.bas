Function GetNumericsAtEnd(ByVal str As String, Optional iSkip As Integer = 0, Optional AtStart As Integer = 0, Optional EnsureComma As Integer = 0) As Integer

    Dim k As Integer
    Dim keepon As Boolean
    Dim cToCheck As Boolean
    awords = Split(str, " ")
    k = 0
    keepon = True
    iSkipped = 0
    While (keepon)
        If ((UBound(awords) - k) < 0) Then
            keepon = False
        Else
            If (AtStart = 0) Then
                cToCheck = IsNumeric(awords(UBound(awords) - k)) Or IsDate(awords(UBound(awords) - k))
                If (EnsureComma = 1) Then cToCheck = (cToCheck And ((InStr(1, awords(UBound(awords) - k), ",") > 0) Or (InStr(1, awords(UBound(awords) - k), ".") > 0)))
            Else
                cToCheck = (IsNumeric(awords(k)) Or IsDate(awords(k)))
                If (EnsureComma = 1) Then cToCheck = (cToCheck And ((InStr(1, awords(k), ",") > 0) Or (InStr(1, awords(k), ".") > 0)))
            End If
            If (cToCheck) Then
                k = k + 1
            Else
                If (iSkipped < iSkip) Then
                    k = k + 1
                    iSkipped = iSkipped + 1
                Else
                    keepon = False
                End If
            End If
        End If
    Wend
    GetNumericsAtEnd = k
    
End Function