Attribute VB_Name = "mdlWMEnc"
Public Encoder As WMEncoder
Public ProfileMgr As WMEncProfileManager
Public ProColl As IWMEncProfileCollection
Public Pro As IWMEncProfile
Public Stats As IWMEncStatistics
Public iEncoderState As WMENC_ENCODER_STATE
Public SrcGrpColl As IWMEncSourceGroupCollection
Public SrcGrp As IWMEncSourceGroup
Public SrcAud As IWMEncSource
Public SrcVid As IWMEncVideoSource2
Public File As IWMEncFile
Public Descr As IWMEncDisplayInfo

Public SourceDuration As Date
Public CurrentDuration As Date

Public SourceFileName As String
Public DestFileName As String
Public EncodingProfileName As String
Public EncodingProfileIndex As Integer

Public AuthorName As String
Public CopyrightInformation As String
Public ContentDescription As String
Public RatingInformation As String
Public ContentTitle As String

Public TimeElapsed As Double
Public Percentage As String

Public Sub Main()
    On Error GoTo HELL
    Set Encoder = New WMEncoder
    Set ProfileMgr = New WMEncProfileManager
    Set ProColl = Encoder.ProfileCollection
    Set Stats = Encoder.Statistics
    Set SrcGrpColl = Encoder.SourceGroupCollection
    Set SrcGrp = SrcGrpColl.Add("SG_1")
    Set SrcAud = SrcGrp.AddSource(WMENC_AUDIO)
    Set SrcVid = SrcGrp.AddSource(WMENC_VIDEO)
    Set File = Encoder.File
    Set Descr = Encoder.DisplayInfo
    Load frmMain
    frmMain.Show
    Exit Sub
HELL:
    MsgBox "Encoder-Engine could not be initialized ! Please re-install Windows Media Encoder !"
    CloseProgram
End Sub

Public Sub CloseProgram()
    On Error Resume Next
    Set Encoder = Nothing
    Set ProfileMgr = Nothing
    Set ProColl = Nothing
    Set Stats = Nothing
    Set SrcGrpColl = Nothing
    Set SrcGrp = Nothing
    Set SrcAud = Nothing
    Set SrcVid = Nothing
    Set File = Nothing
    Set Descr = Nothing
    End
End Sub

Public Sub EnumProfiles(List As ListBox)
    Dim sProfileName As String
    Dim i As Integer
    List.Clear
    For i = 0 To ProColl.Count - 1
        List.AddItem ProColl.Item(i).Name, i
    Next i
    If List.ListCount > 0 Then List.ListIndex = 0
End Sub

Public Sub DisplayProfileManager()
    ProfileMgr.WMEncProfileList WMENC_FILTER_AV, 0
End Sub

Public Sub DisplayProfileDetails(Box As TextBox, ProfileList As ListBox, ProfileIndex As Long)
    Box.Text = ProfileMgr.GetDetailsString(ProfileList.List(ProfileIndex), 0)
End Sub

Public Sub SetInputFile(ByVal FileName As String)
    Dim TempVal As Long
    DoEvents
    SrcAud.SetInput FileName
    SrcVid.SetInput FileName
    SrcGrp.Profile = ProColl.Item(0)
    Encoder.PrepareToEncode True
    TempVal = SrcVid.Duration
    SourceDuration = GetHHMMSS(TempVal)
End Sub

Public Sub SetOutputFile(ByVal FileName As String)
    File.LocalFileName = FileName
End Sub

Public Sub SetDetails()
    Descr.Author = AuthorName
    Descr.Copyright = CopyrightInformation
    Descr.Description = ContentDescription
    Descr.Rating = RatingInformation
    Descr.Title = ContentTitle
End Sub
