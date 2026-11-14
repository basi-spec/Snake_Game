Attribute VB_Name = "Module1"

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Dim sx(1 To 400) As Integer
Dim sy(1 To 400) As Integer
Dim snakeLen As Integer
Dim dirX As Integer
Dim dirY As Integer
Dim lastX As Integer
Dim lastY As Integer
Dim gameRunning As Boolean
Dim paused As Boolean
Dim skipOver As Boolean
Dim foodX As Integer
Dim foodY As Integer
Dim score As Integer

Sub ClearBoard()
    Range("Board").Interior.Color = RGB(15, 27, 45)
End Sub

Sub ReadyScreen()
    ClearBoard
    Range("ScoreCell").Value = 0
    Range("StatusCell").Value = "PRESS START"
    skipOver = False
End Sub

Sub NewGame()
    ClearBoard
    Randomize
    snakeLen = 3
    sx(1) = 10
    sy(1) = 10
    sx(2) = 9
    sy(2) = 10
    sx(3) = 8
    sy(3) = 10
    dirX = 1
    dirY = 0
    lastX = 1
    lastY = 0
    score = 0
    Range("ScoreCell").Value = 0
    DrawSnake
    PlaceFood
End Sub

Sub DrawSnake()
    Dim i As Integer
    For i = 1 To snakeLen
        If i = 1 Then
            Range("Board").Cells(sy(i), sx(i)).Interior.Color = RGB(94, 234, 212)
        Else
            Range("Board").Cells(sy(i), sx(i)).Interior.Color = RGB(20, 145, 127)
        End If
    Next i
End Sub

Sub PlaceFood()
    Dim ok As Boolean
    Dim i As Integer
    Do
        foodX = Int(Rnd * 20) + 1
        foodY = Int(Rnd * 20) + 1
        ok = True
        For i = 1 To snakeLen
            If sx(i) = foodX And sy(i) = foodY Then ok = False
        Next i
    Loop Until ok
    Range("Board").Cells(foodY, foodX).Interior.Color = RGB(251, 191, 36)
End Sub

Sub MoveSnake()
    Dim i As Integer
    Dim tailX As Integer
    Dim tailY As Integer
    tailX = sx(snakeLen)
    tailY = sy(snakeLen)
    For i = snakeLen To 2 Step -1
        sx(i) = sx(i - 1)
        sy(i) = sy(i - 1)
    Next i
    sx(1) = sx(1) + dirX
    sy(1) = sy(1) + dirY
    lastX = dirX
    lastY = dirY
    If sx(1) = foodX And sy(1) = foodY Then
        snakeLen = snakeLen + 1
        sx(snakeLen) = tailX
        sy(snakeLen) = tailY
        score = score + 10
        Range("ScoreCell").Value = score
        PlaceFood
    Else
        Range("Board").Cells(tailY, tailX).Interior.Color = RGB(15, 27, 45)
    End If
    DrawSnake
End Sub

Function HitSomething() As Boolean
    Dim nx As Integer
    Dim ny As Integer
    Dim i As Integer
    nx = sx(1) + dirX
    ny = sy(1) + dirY
    HitSomething = False
    If nx < 1 Or nx > 20 Or ny < 1 Or ny > 20 Then
        HitSomething = True
    Else
        For i = 1 To snakeLen - 1
            If sx(i) = nx And sy(i) = ny Then HitSomething = True
        Next i
    End If
End Function

Sub GameOver()
    Dim i As Integer
    For i = 1 To snakeLen
        Range("Board").Cells(sy(i), sx(i)).Interior.Color = RGB(239, 68, 68)
    Next i
    If score > Range("HighScoreCell").Value Then
        Range("HighScoreCell").Value = score
        Range("StatusCell").Value = "NEW HIGH SCORE!"
    Else
        Range("StatusCell").Value = "GAME OVER"
    End If
End Sub

Sub StartGame()
    If gameRunning Then Exit Sub
    NewGame
    BlockKeys
    paused = False
    skipOver = False
    Range("StatusCell").Value = "PLAYING"
    gameRunning = True
    Do While gameRunning
        If HitSomething() Then
            gameRunning = False
        Else
            MoveSnake
            Pause 0.2
        End If
    Loop
    ResetKeys
    If skipOver Then
        ReadyScreen
    Else
        GameOver
    End If
End Sub

Sub TogglePause()
    If Not gameRunning Then Exit Sub
    paused = Not paused
    If paused Then
        Range("StatusCell").Value = "PAUSED"
    Else
        Range("StatusCell").Value = "PLAYING"
    End If
End Sub

Sub ResetGame()
    If gameRunning Then
        skipOver = True
        paused = False
        gameRunning = False
    Else
        ReadyScreen
    End If
End Sub

Sub Pause(seconds As Single)
    Dim t As Single
    t = Timer
    Do While Timer >= t And Timer < t + seconds
        CheckKeys
        DoEvents
    Loop
    Do While paused And gameRunning
        DoEvents
    Loop
End Sub

Sub CheckKeys()
    If GetAsyncKeyState(38) < 0 Or GetAsyncKeyState(87) < 0 Then GoUp
    If GetAsyncKeyState(40) < 0 Or GetAsyncKeyState(83) < 0 Then GoDown
    If GetAsyncKeyState(37) < 0 Or GetAsyncKeyState(65) < 0 Then GoLeft
    If GetAsyncKeyState(39) < 0 Or GetAsyncKeyState(68) < 0 Then GoRight
End Sub

Sub GoUp()
    If lastY <> 1 Then
        dirX = 0
        dirY = -1
    End If
End Sub

Sub GoDown()
    If lastY <> -1 Then
        dirX = 0
        dirY = 1
    End If
End Sub

Sub GoLeft()
    If lastX <> 1 Then
        dirX = -1
        dirY = 0
    End If
End Sub

Sub GoRight()
    If lastX <> -1 Then
        dirX = 1
        dirY = 0
    End If
End Sub

Sub BlockKeys()
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "w", ""
    Application.OnKey "a", ""
    Application.OnKey "s", ""
    Application.OnKey "d", ""
    Application.OnKey "W", ""
    Application.OnKey "A", ""
    Application.OnKey "S", ""
    Application.OnKey "D", ""
End Sub

Sub ResetKeys()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "w"
    Application.OnKey "a"
    Application.OnKey "s"
    Application.OnKey "d"
    Application.OnKey "W"
    Application.OnKey "A"
    Application.OnKey "S"
    Application.OnKey "D"
End Sub
