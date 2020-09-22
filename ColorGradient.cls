VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ColorGradient is an improvement over the ColorFade class.  With this
' class, you can draw multi-colored gradients on forms or picture boxes.
'
' ColorFade allowed two-color gradients.  This was made a bit more special
' because you could actually read a specific color value from anywhere on
' the fade.  Unfortunately, this isn't implemented in ColorGradient.  It
' very well could be, but that wasn't in my list of goals.
'
' Using ColorGradient: ColorGradient accepts up to 11 colors.  At first it
' was 10, but I wanted 50% right in the middle, with start color at 0%.
'
' Now, to set a color, use the SetColor method.  You specify an index 0 to
' 10.  This puts the color in the spectrum.  You could also use GetColor
' if you want to know what you have been setting.
'
' I hope the comments are enough to explain what is going on...
'


Option Explicit 'I don't normally use this, but here it's ok.


' API Declarations for Radial Gradient.  Thanks to Roger Gilchrist.
Private Const WM_SETREDRAW      As Long = &HB
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long



Private Type FadeColorType
    'Here is the color structure.  fColor is obvious, but fExplicit needs
    'more explaining.  The fader automatically fills in colors between
    'explicit values.  It's really difficult to explain, but read the
    'discussion at the UnSetColor sub.
    fColor      As Long
    fExplicit   As Boolean
End Type
Private Fade(0 To 10) As FadeColorType  'The 10 colors.



Private Type ColorGradientRGBType
    'This is used to move around RGB values without having to use
    'three different variables.  (I know, I know).
    cRed        As Integer
    cGreen      As Integer
    cBlue       As Integer
End Type



Public Enum ColorGradientSourceConstants
    ' The gradient source allows you to specify what direction to start
    ' the fade.  The path leads to the opposite corner/side.  Radial
    ' Gradient has long been a desire of mine, but it is not ready, so
    ' I have included it for forward compatibility.
    gTopLeft = 0
    gTopCenter = 1
    gTopRight = 2
    gCenterRight = 3
    gLowerRight = 4
    gLowerCenter = 5
    gLowerLeft = 6
    gCenterLeft = 7
    gCenterCenter = 8
End Enum



Private Sub Class_Initialize()
    ' We must set the first and last indexes to explicit.  Please read
    ' the discussion in the UnSetColor subroutine.
    
    Fade(0).fColor = RGB(0, 0, 0)           'Default to Black
    Fade(0).fExplicit = True                'Make Explicit
    
    Fade(10).fColor = RGB(255, 255, 255)    'Default to Black
    Fade(10).fExplicit = True               'Make Explicit
    
    UpdateColors    'This re-works the fade.  You can use this class
                    'right out of the box, eg all colors are set.
        
End Sub





Private Sub UpdateColors()
    ' This procedure does all the work of the class.  It iterates through
    ' the color array and fills in the blanks.  Blanks are the particular
    ' indexes that have not been set as explicit.
        
    Dim Explicits() As Integer 'An array to hold the explicits.  Each
                                'element holds an explicit, so when the
                                'class initializes, the entries are
                                '0 and 10.
                                
    Dim iLoop As Integer        'Loop Counter
    Dim jLoop As Integer        'Loop Counter
    
    Dim ExplicitCount As Integer    'The number of explicits that have been
                                    'found in this search.
    
    Dim Distance As Integer         'The distance between two explicits.
    Dim CurPercent As Integer       'The percentage value of the explicit
    
    
    Dim Colors(0 To 2) As ColorGradientRGBType
                                    'Set aside some space to store some
                                    'RGBs.  These will be faded one at a
                                    'time to produce the final product.
       
    ' Make sure the first and last values are explicit.  This is very
    ' important as the fade needs a firm start and end color.
    If Not Fade(0).fExplicit Then Fade(0).fExplicit = True
    If Not Fade(10).fExplicit Then Fade(10).fExplicit = True
        
        
    For iLoop = 0 To 10
        ' We now will loop through all 11 colors and find out who is
        ' explicit.  These colors will not change.  Those in between will.
        If Fade(iLoop).fExplicit Then
                ExplicitCount = ExplicitCount + 1
                
                ReDim Preserve Explicits(ExplicitCount)
                Explicits(ExplicitCount) = iLoop
                'This stuff above should be self-explanatory...
        End If
    Next iLoop
    
    
    
    ' Now the fun starts.  This is a very complex loop that will iterate
    ' through each explicit.  The distance is calculated, and all slots
    ' in between the explicits are filled...
    For iLoop = 1 To ExplicitCount - 1
            ' DetermineRGB is an artifact function from the ColorFade class
            ' that pulls individual color channels from a color long number.
            Call DetermineRGB(Fade(Explicits(iLoop)).fColor, Colors(0))
            Call DetermineRGB(Fade(Explicits(iLoop + 1)).fColor, Colors(2))
            
            ' How far from the first explicit to the next explicit?  We use
            ' this for the next loop.
            Distance = Explicits(iLoop + 1) - Explicits(iLoop)
            
            ' Now we loop through the colors between the explicits.  Each
            ' slot gets a midpoint between the previous slot and the next
            ' slot.  Makes the code way more robust...
            For jLoop = Explicits(iLoop) + 1 To Explicits(iLoop + 1) - 1 Step 1
                    
                ' We need to know what the percentage is between the two
                ' explicits.  If distance is 3, then the first slot will
                ' be 33%.  This formula figures it out for us...
                CurPercent = 100 - (((Explicits(iLoop + 1) - jLoop) * (100 / (Distance * 10))) * 10)
                
                ' Now we use the artifact DetermineFade to find out the
                ' midpoint color.  Another remnant of ColorFade.
                Fade(jLoop).fColor = DetermineFade(Colors(0), _
                                                   Colors(2), _
                                                   CurPercent)
                                                   
            Next jLoop
    Next iLoop

    
End Sub




Public Sub SetColor(ByVal ColorIndex As Integer, ByVal ColorValue As Long)
    ' This procedure sets a color slot to a specified value.  If you
    ' specify a color, then you must want it to be explicit, so we will
    ' take care of that here.
    
    If ColorIndex < 0 Or ColorIndex > 10 Then Exit Sub  'Valid?
    
    If ColorValue < 0 Then
        'If they set an invalid color, just un-explicit it.  But use
        'the UnSetValue method.
        Me.UnSetColor ColorIndex
    Else
        'Set the value
        Fade(ColorIndex).fColor = ColorValue
        Fade(ColorIndex).fExplicit = True
        
        ' Color Table has changed.  Update the slots.
        UpdateColors
    End If
    
End Sub

Public Sub SetColorRGB(ByVal ColorIndex As Integer, ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer)
    ' Wrapper function for SetColor that allows you to specify
    ' individual color channels.
    SetColor ColorIndex, RGB(Red, Green, Blue)
End Sub





Public Sub UnSetColor(ByVal ColorIndex As Integer)
    ' This turns off the explicit indicator for a color slot.  The
    ' indicator assures that a color will not change in the UpdateColors
    ' routine.  Slots that do not have fExplicit set to true are subject
    ' to change based on it's surroundings.
    
    ' The first and last colors have to be explicit
    If ColorIndex < 1 Or ColorIndex > 9 Then Exit Sub  'Valid?
    
    'No need to change the color itself.  UpdateColors will take
    'care of that for us.
    Fade(ColorIndex).fExplicit = False
    UpdateColors
    
End Sub





Public Function GetColor(ByVal ColorIndex As Integer) As Long
    ' Returns the color from the specified slot, even if the
    ' color has not been explicitly set.
    
    If ColorIndex < 0 Or ColorIndex > 10 Then Exit Function 'Valid?
    
    GetColor = Fade(ColorIndex).fColor  'Return the color
    
End Function

















Public Sub PaintObject(obj As Object, Gradient As ColorGradientSourceConstants)
    '
    ' Wrapper function for the private PaintObject function.  You can
    ' specify other types of paintable objects here.  The only actual
    ' requirement is that the object MUST support the Line method using
    ' VBs syntax.
    '
    If (TypeOf obj Is Form) Or (TypeOf obj Is PictureBox) Then
        ' We only support forms and picture boxes for now.
        DoPaintObject obj, Gradient
    End If

End Sub


Private Sub DoPaintObject(obj As Object, Gradient As ColorGradientSourceConstants)
    ' This is the gradient method for objects.  Use this to automate the
    ' drawing of gradients using VBs Line method.  This function does almost
    ' no error checking so make sure that you keep it private and wrap it
    ' with a public function instead.
    
    Dim ScreenX As Long     'Screen dimension
    Dim ScreenY As Long     'Screen dimension
    ScreenX = Screen.TwipsPerPixelX 'Determine the horizontal screen dimension
    ScreenY = Screen.TwipsPerPixelY 'Determine the vertical screen dimension
       
    Dim ObjectWidthPixels As Long   '
    Dim ObjectHeightPixels As Long  '  We need both twips and pixels for
    Dim ObjectWidthTwips As Long    '  the calculations.
    Dim ObjectHeightTwips As Long   '
    
    ObjectWidthTwips = obj.Width
    ObjectHeightTwips = obj.Height
    ObjectWidthPixels = obj.Width / ScreenX
    ObjectHeightPixels = obj.Height / ScreenY
        
    
    Dim HorizontalOffset As Long '45 degrees equals height/2
    HorizontalOffset = ObjectHeightPixels / 2
    
    Dim iLoop As Long   'Loop Counter
    Dim jLoop As Long   'Loop Counter
    
    Dim CurPercent As Single    'The current percentage we are working on
                                'within the fade.  This is a detailed
                                'number.
        
    Dim Radius As Long      'These three are reserved for the radial
    Dim hCenter As Long     'gradients.  Whether or not they work I
    Dim vCenter As Long     'have included the framework.
    
    
    
    Select Case Gradient
        Case gCenterLeft 'Vertical Lines, left to right
            For iLoop = 0 To ObjectWidthPixels
                
                CurPercent = CSng((iLoop / ObjectWidthPixels) * 100)
                obj.Line (iLoop * ScreenX, 0)-(iLoop * ScreenX, ObjectHeightTwips), GetFadeColor(CurPercent)
            Next iLoop
            
        Case gCenterRight 'Vertical lines, right to left
            For iLoop = ObjectWidthPixels To 0 Step -1
                
                CurPercent = 100 - (CSng((iLoop / ObjectWidthPixels) * 100))
                obj.Line (iLoop * ScreenX, 0)-(iLoop * ScreenX, ObjectHeightTwips), GetFadeColor(CurPercent)
            Next iLoop
        
        Case gTopCenter 'Horizontal Lines, top to bottom
            For iLoop = 0 To ObjectHeightPixels
            
                CurPercent = CSng((iLoop / ObjectHeightPixels) * 100)
                obj.Line (0, iLoop * ScreenY)-(ObjectWidthTwips, iLoop * ScreenY), GetFadeColor(CurPercent)
            Next iLoop
        
        
        Case gLowerCenter
            For iLoop = ObjectHeightPixels To 0 Step -1
            
                CurPercent = 100 - (CSng((iLoop / ObjectHeightPixels) * 100))
                obj.Line (0, iLoop * ScreenY)-(ObjectWidthTwips, iLoop * ScreenY), GetFadeColor(CurPercent)
            Next iLoop
                                    
        
        Case gTopLeft ' Diagonal Lines, Upper Left to Lower Right
            For iLoop = 0 - HorizontalOffset To ObjectWidthPixels + HorizontalOffset
                
                CurPercent = CSng((iLoop / ObjectWidthPixels) * 100)
                obj.Line ((iLoop + HorizontalOffset) * ScreenX, 0)-((iLoop - HorizontalOffset) * ScreenX, ObjectHeightTwips), GetFadeColor(CurPercent)
            Next iLoop
           
           
        Case gLowerRight ' Diagonal Lines, Lower Right to Upper Left
            For iLoop = ObjectWidthPixels + HorizontalOffset To 0 - HorizontalOffset Step -1
                
                CurPercent = 100 - (CSng((iLoop / ObjectWidthPixels) * 100))
                obj.Line ((iLoop + HorizontalOffset) * ScreenX, 0)-((iLoop - HorizontalOffset) * ScreenX, ObjectHeightTwips), GetFadeColor(CurPercent)
            Next iLoop
           
        
        Case gTopRight ' Diagonal Lines, Upper Right to Lower Left
            For iLoop = ObjectWidthPixels + HorizontalOffset To 0 - HorizontalOffset Step -1
                
                CurPercent = 100 - (CSng((iLoop / ObjectWidthPixels) * 100))
                obj.Line ((iLoop - HorizontalOffset) * ScreenX, 0)-((iLoop + HorizontalOffset) * ScreenX, ObjectHeightTwips), GetFadeColor(CurPercent)
            Next iLoop
            
            
        Case gLowerLeft ' Diagonal Lines, Lower Left to Upper Right
            For iLoop = 0 - HorizontalOffset To ObjectWidthPixels + HorizontalOffset
                
                CurPercent = CSng((iLoop / ObjectWidthPixels) * 100)
                obj.Line ((iLoop - HorizontalOffset) * ScreenX, 0)-((iLoop + HorizontalOffset) * ScreenX, ObjectHeightTwips), GetFadeColor(CurPercent)
            Next iLoop
            
        Case gCenterCenter 'Radial.  It is broken.
            ' The problem is that when you draw concentric circles you get
            ' dots in the seems.  These are where two circles cannot properly
            ' meet.  You have to understand that circles are emulated in
            ' the computing world.
            
            With obj
                Dim sngLoop As Single
                vCenter = .Height / 2
                hCenter = .Width / 2
                
                Radius = Sqr(vCenter ^ 2 + hCenter ^ 2)
                'the old hypotenus of a triangle rule
                'to get the diagonal of the object(from centre to any corner)
                
                .FillStyle = 0
                
                SendMessage .hWnd, WM_SETREDRAW, False, 0                       ' stop redraw until its done (interesting effect if you comment it out)
                
                For sngLoop = Radius To 0 Step -Radius / 200                     'Step div value 200 is a good value, smaller=circles are obvious, larger= slower
                        .FillColor = GetFadeColor(CSng((sngLoop / Radius) * 100))
                        obj.Circle (hCenter, vCenter), sngLoop, obj.FillColor
                        .Refresh
                Next sngLoop
                SendMessage .hWnd, WM_SETREDRAW, True, 0                      ' reactivate redraw
                .Refresh
            End With
            
        Case Else
            MsgBox "Nope"
            ' The gradient is not found!  For simplicity's sake we
            ' will ignore this error and just do nothing.
    End Select
    


End Sub
























Private Function GetFadeColor(ByVal Percentage As Single) As Long

    ' This returns a color based on a percentage.  If you make this
    ' Public, you will effectively turn this into an advanced ColorFade
    ' class.  If you take the time to read these comments, then you
    ' get good action, eh?
    
    
    ' Correct the percentage.  No need to raise an error.
    If Percentage > 100 Then Percentage = 100
    If Percentage < 0 Then Percentage = 0
        
    Dim iLoop As Integer    'Loop Counter
    
    Dim UpperColor As Long  'The end color for this fade
        
    Dim Colors(0 To 1) As ColorGradientRGBType  'Placeholder for some RGB values.
        
    ' We use this to determine where the percentage lies based on
    ' the indeces.  Yes, you could say "percent / 10" but that wasn't
    ' working.  This method works great and fast enough.
    For iLoop = 0 To 9
        If Percentage < CSng(iLoop * 10) Then Exit For
    Next iLoop
        
    ' Now we know the index of the EndColor.  StartColor is the
    ' same as the endcolor - 1.  Note the these are indexes, not
    ' actual color values.
    UpperColor = iLoop
        
    ' We need to determine
    Percentage = (Percentage - ((iLoop - 1) * 10)) * 10
    
    'Find out the separate color channels for the start and end color.
    Call DetermineRGB(Fade(UpperColor - 1).fColor, Colors(0))
    Call DetermineRGB(Fade(UpperColor).fColor, Colors(1))
            
    GetFadeColor = DetermineFade(Colors(0), _
                                 Colors(1), _
                                 CLng(Percentage))
         
End Function












Private Sub DetermineRGB(ByVal varRGB As Long, ColorRGB As ColorGradientRGBType)
    ' Determines the individual Red, Green, and Blue values from a long
    ' number.  Please note that there is probably a faster mathematical
    ' way of doing this, but I don't know it, and since this procedure is
    ' only called when you set the start or end color to something new, it
    ' really isn't important.
    '
    ' NOTICE:  You must "CALL" this procedure, not just branch.
    '
    ' Update thanks to Roger Gilchrist
            
    With ColorRGB
        .cRed = varRGB Mod 256
        .cGreen = (varRGB And &HFF00FF00) / 256
        .cBlue = (varRGB And &HFF0000) / 65536
    End With
    
End Sub




Private Function DetermineFade(StartColor As ColorGradientRGBType, EndColor As ColorGradientRGBType, ByVal Percent As Long) As Long
        ' Returns a color that sits somewhere between a start and end color.
        ' Basically its a three-part wrapper for the simple MidPoint function.
        '
        ' If you want to improve speed this is the place to do it.
        '
        ' Update thanks to Roger Gilchrest

    
        If Percent > 100 Then Percent = 100
        If Percent < 0 Then Percent = 0
        
        DetermineFade = RGB( _
                MidPoint(StartColor.cRed, EndColor.cRed, Percent), _
                MidPoint(StartColor.cGreen, EndColor.cGreen, Percent), _
                MidPoint(StartColor.cBlue, EndColor.cBlue, Percent))
        
End Function




Private Function MidPoint(x1 As Integer, x2 As Integer, p As Long) As Long
        ' This is a simple interpolation function.  It will find a midpoint
        ' number based on the percentage.  This function is the single most
        ' called function in the entire class, so keep it short and sweet.

        If x1 > x2 Then
            MidPoint = x1 - ((x1 - x2) * (p / 100))
        Else
            MidPoint = ((x2 - x1) * (p / 100)) + x1
        End If

End Function

