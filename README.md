# VB.Net-Smooth-Windows
Smoothly Animate WinForms - Position, Resize, Fade, and Chain effects together. 
Single VB.Module  - works with any project using .NET Framework 4.5 or above.  


Introducing Smooth Windows for VB.Net<br>

# Smooth Form Movement and Resizing Module

A module designed to simplify the process of moving and resizing a form with smooth animations, including fading transparency. <br>
It allows resizing from a specific corner of the window for a more visually appealing effect, based on the screen quadrant.  
it has many easing functions too. 
both duration and Easing functions can be overridden per function call to blend the animations into a larger effect. 


## Usage

### Exposed Functions

#### `Smooth.CreateDynamicForm()`
Displays a dynamic form for testing and exploring the module out of the box. This function can be removed when importing the module into your own projects.
<Br>
<br>

The primary function for moving and/or resizing a window, adjusting its opacity, and animating the process is:

#### `Smooth.MoveWindow(ByRef Form, newX, newY, optional newWidth, optional newHeight, optional opacityStart, optional opacityEnd, optional duration, optional easingFunction, optional callbackFunction)`
<br>
<br>

### Other Available Functions

#### `Smooth.GetTaskbarPosition()`
Returns the position of the taskbar as `"Top"`, `"Bottom"`, `"Left"`, or `"Right"`.

#### `Smooth.GetTaskbarThickness()`
Returns an integer representing the thickness of the taskbar, either in height or width depending on orientation.

#### `Smooth.Wait(milliseconds)`
Non-Blocking Delay, can be used as a sleep command without blocking UI, but also for chaining multiple animations together to create effects

#### `Smooth.SetWindowTransparency(ByRef form, opacityInt1to100)`
Sets the transparency level of a form, where opacity is an integer value between 1 and 100.

#### `Smooth.MoveWindowToScreen(ByRef form, screenIndex)`
Moves a form to a specified monitor index with coordinates set to `x=0`, `y=0`.

#### `Smooth.GetScreenInfo(Optional ByRef form)`
Returns detailed information about the monitor as a `Dictionary(Of String, Dictionary(Of String, Object))`.  
Optional: Can accept a form to return screen data for that specific form.

#### `Smooth.FormatScreenInfo(dictionary)`
Converts the dictionary returned by `GetScreenInfo()` into a neatly formatted string.

#### `Smooth.GetMonitorForForm(dictionary, ByRef form)`
Determines the monitor index where a given form is located.

---

## Example Code Using the Dictionary to Retrieve Screen Information:

```vb
Dim tempDict As Dictionary(Of String, Dictionary(Of String, Object)) = GetScreenInfo(Me)
Dim tempString As String = FormatScreenInfo(tempDict)
MessageBox.Show(tempString)
```
<br>
<br>
<br>
**## Core Module**
~~~~~~~~~~~~~~~~~~~~~


Module smooth

    ' Global variables
    Public moduleAuthor As String = "CoffeeCodeConverter"
    Public moduleCreatedDate As String = "15/01/2025"
    Public moduleVersion As String = "1.0.0.0"


    '   Configurable Variables
    '   ========================================================
    Public animationEnabled As Boolean = True                   ' Disable for better performance. The MoveWindow() Function is still useful for resizing and positioning
    Public animationTimerInterval As Integer = 10               ' Lower means smoother animation but higher CPU usage, i.e 16 = 60fps
    Public animationDuration As Integer = 800                   ' Default Duration in milliseconds, but can be overridden in the Smooth.MoveWindow() function call
    Public animationEasingFunction As String = "EaseOutQuart"   ' Change the easing function - EaseInOutSin, EaseOutCubic, EaseOutQuad, EaseOutQuart
    Public globalWindowMargin As Integer = 10                   ' Margin (pixels) from screen edges and taskbar
    '   ========================================================
    Private isAnimating As Boolean = False                      ' Not used by default, but its there if you need
    Private animationTimer As Timer
    Private startTime As DateTime
    Private startX As Single
    Private startY As Single
    Private startWidth As Single
    Private startHeight As Single
    Private targetX As Single
    Private targetY As Single
    Private targetWidth As Single
    Private targetHeight As Single
    Private currentForm As Form
    Private formOpacityStart As Integer
    Private formOpacityEnd As Integer
    Private currentAnchor As String = "TopLeft" ' Default anchor point
    Private callback As Action
    Public Event AnimationStarted As EventHandler
    Public Event AnimationFinished As EventHandler





    ' Move and/or Resize any Form 
    Public Sub MoveWindow(ByRef form As Form, ByVal newX As Integer, ByVal newY As Integer, Optional ByVal newWidth As Integer = -1, Optional ByVal newHeight As Integer = -1, Optional ByVal opacityStart As Integer = Nothing, Optional ByVal opacityEnd As Integer = Nothing, Optional duration As Integer = 0, Optional easingFunc As String = "", Optional ByVal callbackFunc As Action = Nothing)
        ' Ensure both a start and end were provided, otherwise override to full opacity
        If opacityStart <= 0 Or opacityEnd <= 0 Then
            opacityStart = 100
            opacityEnd = 100
        End If

        If opacityStart > 100 Then
            opacityStart = 100
        End If

        If opacityEnd > 100 Then
            opacityEnd = 100
        End If

        targetX = newX
        targetY = newY

        If newWidth > -1 Then
            targetWidth = newWidth
        Else
            targetWidth = form.Width
        End If

        If newHeight > -1 Then
            targetHeight = newHeight
        Else
            targetHeight = form.Height
        End If

        If duration = 0 Then
            animationDuration = 800
        ElseIf duration > 0 Then
            animationDuration = duration
        End If

        If Not String.IsNullOrEmpty(easingFunc) Then
            Select Case easingFunc
                Case "EaseInOutSin"
                    animationEasingFunction = "EaseInOutSin"
                Case "EaseOutCubic"
                    animationEasingFunction = "EaseOutCubic"
                Case "EaseOutQuad"
                    animationEasingFunction = "EaseOutQuad"
                Case "EaseOutQuart"
                    animationEasingFunction = "EaseOutQuart"
                Case "EaseOutBounce"
                    animationEasingFunction = "EaseOutBounce"
                Case "EaseOutElastic"
                    animationEasingFunction = "EaseOutElastic"
                Case "EaseOutQuint"
                    animationEasingFunction = "EaseOutQuint"
                Case "EaseInOutQuad"
                    animationEasingFunction = "EaseInOutQuad"
                Case "EaseInOutCubic"
                    animationEasingFunction = "EaseInOutCubic"
                Case "EaseInOutQuart"
                    animationEasingFunction = "EaseInOutQuart"
                Case "EaseInOutQuint"
                    animationEasingFunction = "EaseInOutQuint"
                Case "EaseOutBack"
                    animationEasingFunction = "EaseOutBack"
                Case Else
                    animationEasingFunction = "EaseOutQuart"
            End Select
        End If

        ' Record the start values
        startX = form.Left
        startY = form.Top
        startWidth = form.Width
        startHeight = form.Height
        currentForm = form

        formOpacityStart = opacityStart
        formOpacityEnd = opacityEnd
        currentForm.Opacity = formOpacityStart

        ' Attach event handlers to the events
        AddHandler AnimationStarted, AddressOf OnAnimationStarted
        AddHandler AnimationFinished, AddressOf OnAnimationFinished

        callback = callbackFunc

        ' Get the center point of the form to determine which screen qudarant its in
        Dim formCenter As Point = GetFormCenterPoint(currentForm)
        Select Case GetScreenQuadrant(formCenter)
            Case "TopLeft"
                currentAnchor = "TopLeft"
            Case "TopRight"
                currentAnchor = "TopRight"
            Case "BottomLeft"
                currentAnchor = "BottomLeft"
            Case "BottomRight"
                currentAnchor = "BottomRight"
        End Select

        ' Get the taskbar position and size
        Dim taskbarPosition As String = GetTaskbarPosition()
        Dim taskbarMargin As Integer = 0

        If taskbarPosition = "Top" OrElse taskbarPosition = "Bottom" Then
            taskbarMargin = FuncTaskbarHeight()
        ElseIf taskbarPosition = "Left" OrElse taskbarPosition = "Right" Then
            taskbarMargin = FuncTaskbarWidth()
        End If

        ' Adjust the target position based on the taskbar position so it never moves, or resizes off-screen
        Dim taskbarHeight As Integer = FuncTaskbarHeight()
        Dim taskbarWidth As Integer = FuncTaskbarWidth()
        If taskbarPosition = "Top" Then
            If targetX < globalWindowMargin Then
                targetX = globalWindowMargin
            End If

            If targetX + targetWidth > (Screen.PrimaryScreen.Bounds.Width - globalWindowMargin - taskbarWidth) Then
                targetX = Screen.PrimaryScreen.Bounds.Width - targetWidth - globalWindowMargin
            End If

            If targetY < (taskbarHeight + globalWindowMargin) Then
                targetY = taskbarHeight + globalWindowMargin
            End If

            If targetY + targetHeight > (Screen.PrimaryScreen.Bounds.Height - globalWindowMargin) Then
                targetY = Screen.PrimaryScreen.Bounds.Height - targetHeight - globalWindowMargin
            End If

        ElseIf taskbarPosition = "Bottom" Then
            If targetX < globalWindowMargin Then
                targetX = globalWindowMargin
            End If

            If targetX + targetWidth > (Screen.PrimaryScreen.Bounds.Width - globalWindowMargin - taskbarWidth) Then
                targetX = Screen.PrimaryScreen.Bounds.Width - targetWidth - globalWindowMargin
            End If

            If targetY < globalWindowMargin Then
                targetY = globalWindowMargin
            End If

            If targetY + targetHeight > (Screen.PrimaryScreen.Bounds.Height - globalWindowMargin - taskbarHeight) Then
                targetY = Screen.PrimaryScreen.Bounds.Height - targetHeight - taskbarHeight - globalWindowMargin
            End If

        ElseIf taskbarPosition = "Left" Then
            If targetX < taskbarWidth + globalWindowMargin Then
                targetX = taskbarWidth + globalWindowMargin
            End If

            If targetX + targetWidth > (Screen.PrimaryScreen.Bounds.Width - globalWindowMargin - taskbarWidth) Then
                targetX = Screen.PrimaryScreen.Bounds.Width - targetWidth - globalWindowMargin
            End If

            If targetY < globalWindowMargin Then
                targetY = globalWindowMargin
            End If

            If targetY + targetHeight > (Screen.PrimaryScreen.Bounds.Height - globalWindowMargin) Then
                targetY = Screen.PrimaryScreen.Bounds.Height - targetHeight - globalWindowMargin
            End If

        ElseIf taskbarPosition = "Right" Then
            If targetX < globalWindowMargin Then
                targetX = globalWindowMargin
            End If

            If targetX + targetWidth > (Screen.PrimaryScreen.Bounds.Width - globalWindowMargin - taskbarWidth) Then
                targetX = (Screen.PrimaryScreen.Bounds.Width - globalWindowMargin - targetWidth - taskbarWidth)
            End If

            If targetY < globalWindowMargin Then
                targetY = globalWindowMargin
            End If

            If targetY + targetHeight > (Screen.PrimaryScreen.Bounds.Height - globalWindowMargin) Then
                targetY = Screen.PrimaryScreen.Bounds.Height - targetHeight - globalWindowMargin
            End If
        End If


        ' Raise Event always, even if animation is Disabled
        RaiseEvent AnimationStarted(currentForm, EventArgs.Empty)

        If animationEnabled = True Then
            ' Calculate the total duration of the animation
            startTime = DateTime.Now
            isAnimating = True

            ' Initialize the timer to run every 10ms
            If animationTimer Is Nothing Then
                animationTimer = New Timer()
                AddHandler animationTimer.Tick, AddressOf AnimationTick
                animationTimer.Interval = animationTimerInterval
            End If
            animationTimer.Start()

        Else
            ' Just Snap To Size and Position without Animating 
            currentForm.Width = targetWidth
            currentForm.Height = targetHeight
            currentForm.Location = New Point(targetX, targetY)
            currentForm.Opacity = formOpacityEnd

            RaiseEvent AnimationFinished(currentForm, EventArgs.Empty)

            ' RUN CALLBACK FUNCTION if it's not Nothing
            If callback IsNot Nothing Then
                Debug.WriteLine("Callback triggered immediately (no animation).")
                callback.Invoke()
            End If
        End If
    End Sub






    ' Functions called by Raised Events - add any code you like to expand on this module code
    Private Sub OnAnimationStarted(sender As Object, e As EventArgs)
        'Debug.WriteLine("Animation Started!")

    End Sub

    Private Sub OnAnimationFinished(sender As Object, e As EventArgs)
        'Debug.WriteLine("Animation Finished!")

    End Sub




    ' overrides and insists topLeft Quadrant if both moving AND resizing the form
    ' Becasue visually, its idential to growing from a specific corner due to the X and Y moving to cancel out the resize
    ' creating the illusion that its growing from the correct corner based on its quadrant, but its techncially always growing from topleft
    Private Sub AnimationTick(sender As Object, e As EventArgs)
        ' progress of animation
        Dim progress As Double = (DateTime.Now - startTime).TotalMilliseconds / animationDuration
        Dim isResizing As Boolean = targetWidth <> startWidth OrElse targetHeight <> startHeight
        Dim isMoving As Boolean = targetX <> startX OrElse targetY <> startY

        ' If both moving and resizing by pass the anchor logic and assume TopLeft
        If isMoving AndAlso isResizing Then

            Dim newXPosition As Integer = CInt(startX + ((targetX - startX) * easingFunctionToUse(progress)))
            Dim newYPosition As Integer = CInt(startY + ((targetY - startY) * easingFunctionToUse(progress)))
            Dim newWidth As Integer = CInt(startWidth + ((targetWidth - startWidth) * easingFunctionToUse(progress)))
            Dim newHeight As Integer = CInt(startHeight + ((targetHeight - startHeight) * easingFunctionToUse(progress)))

            currentForm.Location = New Point(newXPosition, newYPosition)
            currentForm.Size = New Size(newWidth, newHeight)

        ElseIf isResizing Then
            ' If only resizing, apply the existing anchor logic
            Dim newXPosition As Integer = startX
            Dim newYPosition As Integer = startY
            Dim newWidth As Integer = CInt(startWidth + ((targetWidth - startWidth) * easingFunctionToUse(progress)))
            Dim newHeight As Integer = CInt(startHeight + ((targetHeight - startHeight) * easingFunctionToUse(progress)))

            ' Apply anchor-based logic for resizing from a specific corner of the form
            ' Top Left is Default anyway, don't need to specify 
            If currentAnchor = "TopRight" Then
                newXPosition = startX + startWidth - newWidth
            ElseIf currentAnchor = "BottomLeft" Then
                newYPosition = startY + startHeight - newHeight
            ElseIf currentAnchor = "BottomRight" Then
                newXPosition = startX + startWidth - newWidth
                newYPosition = startY + startHeight - newHeight
            End If

            currentForm.Location = New Point(newXPosition, newYPosition)
            currentForm.Size = New Size(newWidth, newHeight)

        ElseIf isMoving Then
            ' If only moving, apply the eased position normally
            Dim newXPosition As Integer = CInt(startX + ((targetX - startX) * easingFunctionToUse(progress)))
            Dim newYPosition As Integer = CInt(startY + ((targetY - startY) * easingFunctionToUse(progress)))
            currentForm.Location = New Point(newXPosition, newYPosition)
        End If

        ' Calculate the new opacity based on the progress
        Dim newOpacity As Double = formOpacityStart + ((formOpacityEnd - formOpacityStart) * progress)
        currentForm.Opacity = Math.Min(Math.Max(newOpacity, 0), 100) / 100 ' Ensure opacity is between 0 and 100


        ' If animation complete, stop the timer
        If progress >= 1 Then
            animationTimer.Stop()
            isAnimating = False

            RaiseEvent AnimationFinished(currentForm, EventArgs.Empty)

            ' RUN CALLBACK FUNCTION if it's not Nothing
            If callback IsNot Nothing Then
                Debug.WriteLine("Callback triggered (after animation).")
                callback.Invoke()
            End If

        End If
    End Sub







    ' To Determine which screen Quadrant the centre point falls in 
    Private Function GetFormCenterPoint(ByRef frm As Form) As Point
        Dim centerX As Integer = frm.Left + (frm.Width \ 2)
        Dim centerY As Integer = frm.Top + (frm.Height \ 2)
        Return New Point(centerX, centerY)
    End Function


    Private Function GetScreenQuadrant(point As Point) As String
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Dim centerX As Integer = screenWidth \ 2
        Dim centerY As Integer = screenHeight \ 2

        ' Determine which quadrant the point falls into
        If point.X < centerX And point.Y < centerY Then
            Return "TopLeft"
        ElseIf point.X >= centerX And point.Y < centerY Then
            Return "TopRight"
        ElseIf point.X < centerX And point.Y >= centerY Then
            Return "BottomLeft"
        Else
            Return "BottomRight"
        End If
    End Function







    ' Get the taskbar position (Top, Bottom, Left, Right)
    Public Function GetTaskbarPosition() As String
        Dim taskbarPosition As String = ""
        Dim screen As Screen = System.Windows.Forms.Screen.PrimaryScreen    ' Task bar is always on primary screen 

        ' Check if the taskbar is vertical (Left or Right)
        If screen.Bounds.Width - screen.WorkingArea.Width > 0 Then
            If screen.WorkingArea.Top = 0 And screen.WorkingArea.Left > 0 Then
                taskbarPosition = "Left"
            ElseIf screen.WorkingArea.Top = 0 And screen.WorkingArea.Left = 0 Then
                taskbarPosition = "Right"
            End If
            ' Check if the taskbar is horizontal (Top or Bottom)
        ElseIf screen.Bounds.Height - screen.WorkingArea.Height > 0 Then
            If screen.WorkingArea.Left = 0 And screen.WorkingArea.Top > 0 Then
                taskbarPosition = "Top"
            ElseIf screen.WorkingArea.Left = 0 And screen.WorkingArea.Top = 0 Then
                taskbarPosition = "Bottom"
            End If
        End If

        Return taskbarPosition
    End Function

    Public Function GetTaskbarThickness() As Integer
        Dim taskbarPosition As String = GetTaskbarPosition()
        If GetTaskbarPosition() = "Top" OrElse taskbarPosition = "Bottom" Then
            Return FuncTaskbarHeight()
        Else
            Return FuncTaskbarWidth()
        End If
    End Function


    Private Function FuncTaskbarWidth() As Integer
        Dim screen As Screen = System.Windows.Forms.Screen.PrimaryScreen
        Return screen.Bounds.Width - screen.WorkingArea.Width
    End Function

    Private Function FuncTaskbarHeight() As Integer
        Dim screen As Screen = System.Windows.Forms.Screen.PrimaryScreen
        Return screen.Bounds.Height - screen.WorkingArea.Height
    End Function








    Public Sub SetWindowTransparency(ByRef form As Form, opacityPercentage As Integer)
        ' Ensure the value is between 0 and 100
        If opacityPercentage < 0 Then
            opacityPercentage = 0
        ElseIf opacityPercentage > 100 Then
            opacityPercentage = 100
        End If

        ' Convert the integer percentage to a double between 0.0 and 1.0
        form.Opacity = opacityPercentage / 100.0
    End Sub






    Public Async Function Wait(milliseconds As Integer) As Task
        ' Asynchronously wait for the specified time without blocking UI
        Await Task.Delay(milliseconds)

        ' Do something after the delay here
    End Function









    ' Change easing function across the board
    Private Function easingFunctionToUse(passedT As Single) As Single
        Select Case animationEasingFunction
            Case "EaseInOutSin"
                Return EaseInOutSin(passedT)
            Case "EaseOutCubic"
                Return EaseOutCubic(passedT)
            Case "EaseOutQuad"
                Return EaseOutQuad(passedT)
            Case "EaseOutQuart"
                Return EaseOutQuart(passedT)
            Case "EaseOutBounce"
                Return EaseOutBounce(passedT)
            Case "EaseOutElastic"
                Return EaseOutElastic(passedT)
            Case "EaseOutQuint"
                Return EaseOutQuint(passedT)
            Case "EaseInOutQuad"
                Return EaseInOutQuad(passedT)
            Case "EaseInOutCubic"
                Return EaseInOutCubic(passedT)
            Case "EaseInOutQuart"
                Return EaseInOutQuart(passedT)
            Case "EaseInOutQuint"
                Return EaseInOutQuint(passedT)
            Case "EaseOutBack"
                Return EaseOutBack(passedT)
            Case Else
                Return EaseOutQuart(passedT)
        End Select
    End Function



    ' Easing functions

    Private Function EaseInOutSin(t As Single) As Single
        Return 0.5 * (1 - Math.Cos(Math.PI * t))
    End Function

    ' --- EaseOut (Mild to Dramatic) ---

    Private Function EaseOutQuad(t As Single) As Single
        Return 1 - (1 - t) * (1 - t)
    End Function

    Private Function EaseOutCubic(t As Single) As Single
        Return 1 - Math.Pow(1 - t, 3)
    End Function

    Private Function EaseOutQuart(t As Single) As Single
        Return 1 - Math.Pow(1 - t, 4)
    End Function

    Private Function EaseOutQuint(t As Single) As Single
        Return 1 - Math.Pow(1 - t, 5)
    End Function

    ' --- EaseInOut (Mild to Dramatic) ---

    Private Function EaseInOutQuad(t As Single) As Single
        If t < 0.5 Then
            Return 2 * t * t
        Else
            Return 1 - Math.Pow(-2 * t + 2, 2) / 2
        End If
    End Function

    Private Function EaseInOutCubic(t As Single) As Single
        If t < 0.5 Then
            Return 4 * t * t * t
        Else
            Return 1 - Math.Pow(-2 * t + 2, 3) / 2
        End If
    End Function

    Private Function EaseInOutQuart(t As Single) As Single
        If t < 0.5 Then
            Return 8 * t * t * t * t
        Else
            Return 1 - Math.Pow(-2 * t + 2, 4) / 2
        End If
    End Function

    Private Function EaseInOutQuint(t As Single) As Single
        If t < 0.5 Then
            Return 16 * t * t * t * t * t
        Else
            Return 1 - Math.Pow(-2 * t + 2, 5) / 2
        End If
    End Function

    ' --- Bounce (Mild to Dramatic) ---

    Private Function EaseOutBounce(t As Single) As Single
        If t < (1 / 2.75) Then
            Return 7.5625 * t * t
        ElseIf t < (2 / 2.75) Then
            t -= (1.5 / 2.75)
            Return 7.5625 * t * t + 0.75
        ElseIf t < (2.5 / 2.75) Then
            t -= (2.25 / 2.75)
            Return 7.5625 * t * t + 0.9375
        Else
            t -= (2.625 / 2.75)
            Return 7.5625 * t * t + 0.984375
        End If
    End Function

    ' --- Elastic (Mild to Dramatic) ---

    Private Function EaseOutElastic(t As Single) As Single
        If t = 0 Then Return 0
        If t = 1 Then Return 1
        Dim p As Single = 0.3
        Dim s As Single = p / 4
        Return Math.Pow(2, -10 * t) * Math.Sin((t - s) * (2 * Math.PI) / p) + 1
    End Function


    Private Function EaseOutBack(t As Single) As Single
        Dim s As Single = 1.70158
        Return 1 + s * Math.Pow(t - 1, 3) + s * Math.Pow(t - 1, 2)
    End Function
End Module
~~~~~~~~~~~~~~~~~~~~~
