Attribute VB_Name = "modPAN"
'polygon drawing stuff
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'drawing stuff
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'PAN files
'Polygon data Structure
Public Type PolyShape
       PolyType As Byte '0 = polygon, 1 = rect, 2=line, 3=Ellips
       PolyPnt() As POINTAPI
       PntCount As Long 'if its a polygon, this is the count of points
       PolyColor As Long
End Type

'Frame Data Structure
Public Type PolyFrame
       PolyShp() As PolyShape 'multiple shapes/polygons
       PolyCount As Byte      'number of shapes/polygons
End Type

'File Data Structure
Public Type polyPAN
     Polys() As PolyFrame
     OutLineColor As Long     'what color is everything outlined in?
     FrameCount As Long
End Type

'=====================
' editor only
'=====================
Public TempPANI As polyPAN

Public CurFrame As PolyFrame
Public CurShape As PolyShape
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'save structure 'should be dimed in save procedure so it
'can clean up properly...
'Public PANI As polyPAN
'===========================

Public Sub SavePan(Filename As String)
On Error GoTo ErrOut:
Dim IntBinaryFile As Long 'file handle
Dim I As Long 'counter
Dim C As Long 'counter 2
Dim Z As Long 'counter 3

Dim PANI As polyPAN 'save structure

'dont try to save if they hit cancel
If Filename = "" Then Exit Sub
    
    frmMain.lblStat.Caption = "Saving File..."
    frmMain.lblStat.Refresh

'======================
'match file data with temp structure

'set frame count
PANI.FrameCount = TempPANI.FrameCount
PANI.OutLineColor = TempPANI.OutLineColor

'resize frame structure
ReDim PANI.Polys(1 To PANI.FrameCount)

      
'for each frame...
For I = 1 To PANI.FrameCount
'how many polygons in this frame?
PANI.Polys(I).PolyCount = TempPANI.Polys(I).PolyCount

'size array to file each polygon...
ReDim Preserve PANI.Polys(I).PolyShp(1 To PANI.Polys(I).PolyCount)

'for each polygon in the frame we're in
    For C = 1 To PANI.Polys(I).PolyCount
        PANI.Polys(I).PolyShp(C).PntCount = TempPANI.Polys(I).PolyShp(C).PntCount
        PANI.Polys(I).PolyShp(C).PolyColor = TempPANI.Polys(I).PolyShp(C).PolyColor
        PANI.Polys(I).PolyShp(C).PolyType = TempPANI.Polys(I).PolyShp(C).PolyType
        
        'make room for each point...
        ReDim Preserve PANI.Polys(I).PolyShp(C).PolyPnt(1 To PANI.Polys(I).PolyShp(C).PntCount)
        
        'loop through each point and collect data
            For Z = 1 To PANI.Polys(I).PolyShp(C).PntCount
                PANI.Polys(I).PolyShp(C).PolyPnt(Z).x = TempPANI.Polys(I).PolyShp(C).PolyPnt(Z).x
                PANI.Polys(I).PolyShp(C).PolyPnt(Z).y = TempPANI.Polys(I).PolyShp(C).PolyPnt(Z).y
            Next
     Next
Next
'=================================================
'============ DONE MATCHING DATA =================
'=================================================

'open binary file to write to
    IntBinaryFile = FreeFile
        
        Open Filename For Binary Access Write Lock Write As IntBinaryFile
            
            'write data
            Put IntBinaryFile, 1, PANI

        'close file cuz we're done with it
        Close IntBinaryFile
    
    frmMain.lblStat.Caption = ""
Exit Sub
ErrOut:
    If PANI.FrameCount = 0 Then
      MsgBox "Unable to process file: Blank/Empty movie.", vbExclamation, "Save Error"
      Else
      MsgBox Err.Description, vbExclamation
     End If
End Sub

Public Sub LoadPan(Filename As String)
On Error GoTo ErrOut:
If Filename = "" Then Exit Sub
Dim IntBinaryFile
    IntBinaryFile = FreeFile
frmMain.lblStat.Caption = "Loading file.."
Open Filename For Binary Access Read Lock Write As IntBinaryFile
    
    'Extract the data
    Get IntBinaryFile, 1, TempPANI
Close IntBinaryFile

If TempPANI.FrameCount > 0 Then
frmMain.cmdTB(9).Enabled = True
frmMain.cmdTB(2).Enabled = False
frmMain.lstFrames.Clear
For I = 1 To TempPANI.FrameCount
    frmMain.lstFrames.AddItem "Frame: " & TempPANI.Polys(I).PolyCount & " Objects"
Next
End If
frmMain.lblStat.Caption = "Animation Loaded: " & TempPANI.FrameCount & " frames"
Exit Sub
ErrOut:
MsgBox "An error occured when trying to load this file.  Please make sure that this is a valid PAN file and that it is currently accessable.", vbExclamation
End Sub

Public Sub DrawFrame(FrameZ As Long, Animation As polyPAN, hdc As Long)
Dim I As Long
Dim P As POINTAPI
'On Error Resume Next
'draw color...

''DeleteObject SelectObject(hdc, CreatePen(0, 1, Animation.OutLineColor))

frmMain.lblStat.Caption = "Frame: " & FrameZ

'for each shape in this frame
    For I = 1 To Animation.Polys(FrameZ).PolyCount
        'select color for this polygon...
        DeleteObject SelectObject(hdc, CreateSolidBrush(Animation.Polys(FrameZ).PolyShp(I).PolyColor))
        
        'what type are we drawing...
        Select Case Animation.Polys(FrameZ).PolyShp(I).PolyType
         Case Is = 0 'polygon
           Polygon hdc, Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1), Animation.Polys(FrameZ).PolyShp(I).PntCount
         Case Is = 1 'rect
           Rectangle hdc, Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1).x, _
                          Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1).y, _
                          Animation.Polys(FrameZ).PolyShp(I).PolyPnt(2).x, _
                          Animation.Polys(FrameZ).PolyShp(I).PolyPnt(2).y
         Case Is = 2 'line
           MoveToEx hdc, Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1).x, _
                         Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1).y, P
                         
           LineTo hdc, Animation.Polys(FrameZ).PolyShp(I).PolyPnt(2).x, _
                       Animation.Polys(FrameZ).PolyShp(I).PolyPnt(2).y
         Case Is = 3 'ellipse
            Ellipse hdc, Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1).x, _
                        Animation.Polys(FrameZ).PolyShp(I).PolyPnt(1).y, _
                        Animation.Polys(FrameZ).PolyShp(I).PolyPnt(2).x, _
                        Animation.Polys(FrameZ).PolyShp(I).PolyPnt(2).y
        End Select
    Next
End Sub
