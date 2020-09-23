Attribute VB_Name = "mdlStuff"
Public Function GetHHMMSS(ByVal ms As Long) As String
    sg = Int(ms / 1000)
    mn = Int(sg / 60)
    hh = Fix(mn / 60)
    If mn > 59 Then
        mn = mn Mod 60
    End If
    zz = sg Mod 60
    GetHHMMSS = Format(Str$(hh), "0#") + ":" + Format(Str$(mn), "0#") + ":" + Format(Str$(zz), "0#")
End Function

