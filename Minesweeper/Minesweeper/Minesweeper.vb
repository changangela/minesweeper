Public Class MineSweeper
    Private Const GRID As Integer = 30
    Private Const BOMB_NUM As Integer = 200

    Private Const BOMB As Integer = 100
    Private Const SPACE As Integer = 99

    Private Const BUTTON_SIZE As Integer = 25

    Private gridBox(GRID - 1, GRID - 1) As Integer
    Private markBox(GRID - 1, GRID - 1) As Boolean
    Private Function initGrid()
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                gridBox(x, y) = SPACE
                markBox(x, y) = False
            Next
        Next
    End Function
    Private Function isInitGame() As Boolean
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                If Not isSpace(x, y) Then Return False
            Next
        Next
        Return True
    End Function
    Function bombGrid(ByVal x0 As Integer, ByVal y0 As Integer)
        Randomize()
        Dim num As Integer = 0
        While (num < BOMB_NUM)
            Dim x As Integer = CInt(Int(((GRID - 1) * Rnd()) + 1))
            Dim y As Integer = CInt(Int(((GRID - 1) * Rnd()) + 1))
            If (x = x0 And y = y0) Or isBomb(x, y) Then Continue While
            gridBox(x, y) = BOMB
            num = num + 1
        End While
    End Function
    Function numerGrid()
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                If Not isSpace(x, y) Then Continue For
                gridBox(x, y) = countBomb(x, y, False)
            Next
        Next
    End Function
    Function countBomb(ByVal x0 As Integer, ByVal y0 As Integer, ByVal markedBombFlag As Boolean) As Integer
        Dim num As Integer = 0
        For y = y0 - 1 To y0 + 1
            For x = x0 - 1 To x0 + 1
                If x = x0 And y = y0 Then Continue For
                If x < 0 Or y < 0 Or x >= GRID Or y >= GRID Then Continue For

                If (markedBombFlag = False And isBomb(x, y)) Or (markedBombFlag = True And isMarkedGrid(x, y)) Then
                    num = num + 1
                End If
            Next
        Next
        If num = 0 Then num = SPACE
        Return num
    End Function
    Function fillSpace(ByVal x0 As Integer, ByVal y0 As Integer)
        openGrid(x0, y0)
        unmarkGrid(x0, y0)
        Dim positions = New Point() {New Point(-1, 0), New Point(+1, 0), New Point(0, -1), New Point(0, +1), New Point(-1, -1), New Point(+1, +1), New Point(+1, -1), New Point(-1, +1)}
        For Each position As Point In positions
            Dim x = x0 + position.X
            Dim y = y0 + position.Y
            If x >= 0 And x < GRID And y >= 0 And y < GRID Then
                If isSpace(x, y) And Not isOpenedGrid(x, y) Then
                    fillSpace(x, y)
                Else
                    openGrid(x, y)
                    unmarkGrid(x, y)
                End If
            End If
        Next
    End Function
    Function getNumber(ByVal x As Integer, ByVal y As Integer) As Integer
        Return Math.Abs(gridBox(x, y))
    End Function
    Function isNumber(ByVal x As Integer, ByVal y As Integer) As Boolean
        Return Not (isSpace(x, y) Or isBomb(x, y))
    End Function
    Function isBomb(ByVal x As Integer, ByVal y As Integer) As Boolean
        Return gridBox(x, y) = BOMB Or gridBox(x, y) = -BOMB
    End Function
    Function isSpace(ByVal x As Integer, ByVal y As Integer) As Boolean
        Return gridBox(x, y) = SPACE Or gridBox(x, y) = -SPACE
    End Function
    Function openGrid(ByVal x As Integer, ByVal y As Integer)
        gridBox(x, y) = -Math.Abs(gridBox(x, y))
    End Function
    Function isOpenedGrid(ByVal x As Integer, ByVal y As Integer) As Boolean
        Return gridBox(x, y) < 0
    End Function
    Function isMarkedGrid(ByVal x As Integer, ByVal y As Integer) As Boolean
        Return markBox(x, y)
    End Function
    Function unmarkGrid(ByVal x As Integer, ByVal y As Integer)
        markBox(x, y) = False
    End Function
    Function markGrid(ByVal x As Integer, ByVal y As Integer)
        markBox(x, y) = True
    End Function
    Function initGame()
        initGrid()
        displayGame(False)
    End Function
    Function startGame(ByVal x0 As Integer, ByVal y0 As Integer)
        'initGrid()
        bombGrid(x0, y0)
        numerGrid()
        displayGame(False)
    End Function
    Function displayAround(ByVal x0 As Integer, ByVal y0 As Integer, ByVal normalFlag As Boolean)
        Dim positions = New Point() {New Point(-1, 0), New Point(+1, 0), New Point(0, -1), New Point(0, +1), New Point(-1, -1), New Point(+1, +1), New Point(+1, -1), New Point(-1, +1)}
        For Each position As Point In positions
            Dim x = x0 + position.X
            Dim y = y0 + position.Y
            If x >= 0 And x < GRID And y >= 0 And y < GRID Then
                If (Not isOpenedGrid(x, y)) And (Not isMarkedGrid(x, y)) Then
                    If normalFlag Then
                        btnArray(x, y).FlatStyle = FlatStyle.Standard
                        btnArray(x, y).FlatAppearance.BorderColor = Color.Black
                    Else
                        btnArray(x, y).FlatStyle = FlatStyle.Flat
                        btnArray(x, y).FlatAppearance.BorderColor = Color.LightGray
                    End If
                End If
            End If
        Next
    End Function
    Function displayGame(ByVal showAll As Boolean)
        Dim colors = New Color() {Color.Blue, Color.Green, Color.Red, Color.Purple, Color.Brown, Color.Turquoise, Color.Chocolate, Color.Cyan}
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                btnArray(x, y).Enabled = True
                btnArray(x, y).ForeColor = Color.Red
                If Not showAll And Not isOpenedGrid(x, y) And isMarkedGrid(x, y) Then
                    btnArray(x, y).FlatStyle = FlatStyle.Standard
                    btnArray(x, y).FlatAppearance.BorderColor = Color.Black
                    btnArray(x, y).Text = "?"
                Else
                    If isBomb(x, y) Then
                        btnArray(x, y).FlatStyle = FlatStyle.Standard
                        btnArray(x, y).FlatAppearance.BorderColor = Color.Black
                        If showAll Then
                            btnArray(x, y).Text = "@"
                        Else
                            btnArray(x, y).Text = ""
                        End If
                    ElseIf isSpace(x, y) Then
                        If showAll Or isOpenedGrid(x, y) Then
                            btnArray(x, y).Enabled = False
                            btnArray(x, y).FlatStyle = FlatStyle.Flat
                            btnArray(x, y).FlatAppearance.BorderColor = Color.LightGray
                        Else
                            btnArray(x, y).FlatStyle = FlatStyle.Standard
                            btnArray(x, y).FlatAppearance.BorderColor = Color.Black
                        End If
                        btnArray(x, y).Text = ""
                    ElseIf isNumber(x, y) Then
                        If isOpenedGrid(x, y) Then
                            btnArray(x, y).FlatStyle = FlatStyle.Flat
                            btnArray(x, y).FlatAppearance.BorderColor = Color.LightGray
                        End If
                        If showAll Or isOpenedGrid(x, y) Then
                            btnArray(x, y).ForeColor = colors(getNumber(x, y) - 1)
                            btnArray(x, y).Text = CStr(getNumber(x, y))
                        Else
                            btnArray(x, y).FlatStyle = FlatStyle.Standard
                            btnArray(x, y).FlatAppearance.BorderColor = Color.Black
                            btnArray(x, y).Text = ""
                        End If
                    End If
                End If
            Next
        Next
    End Function
    Private Function isWinGame() As Boolean
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                If (Not isOpenedGrid(x, y)) And (Not isBomb(x, y)) Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function
    Private Sub buttonMouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) ' Handles Me.MouseClick
        If failFlag Or winFlag Then
            If failFlag Then
                MsgBox("You lost!")
            ElseIf winFlag Then
                MsgBox("You won!")
            End If
            failFlag = False
            winFlag = False
            initGame()
            Return
        End If
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                If sender Is btnArray(x, y) Then
                    If e.Button = Windows.Forms.MouseButtons.Right Then
                        If isOpenedGrid(x, y) And isNumber(x, y) Then
                            displayAround(x, y, True)
                            Return
                        End If
                    End If
                End If
            Next
        Next
    End Sub
    Dim failFlag, winFlag As Boolean
    Private Sub buttonMouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) ' Handles Me.MouseClick
        Dim x, y As Integer
        Dim flagFind As Boolean = False
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                If sender Is btnArray(x, y) Then
                    flagFind = True
                    Exit For
                End If
            Next
            If flagFind Then Exit For
        Next
        If Not flagFind Then Return

        Dim gameStartedFlag As Boolean = Not isInitGame()
        If Not gameStartedFlag Then
            If e.Button = Windows.Forms.MouseButtons.Left Then
                startGame(x, y)
            End If
            gameStartedFlag = True
        End If
        If Not gameStartedFlag Then Return

        Dim continueFlag As Boolean = True
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If e.Clicks = 1 Then
                continueFlag = buttonRightClick(x, y)
                If isOpenedGrid(x, y) And isNumber(x, y) Then
                    displayAround(x, y, False)
                    Return
                End If
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
            If e.Clicks = 2 Then
                continueFlag = buttonDoubleClick(x, y)
            ElseIf e.Clicks = 1 Then
                continueFlag = buttonLeftClick(x, y)
            End If
        End If
        displayGame(Not continueFlag)
        If Not continueFlag Then
            failFlag = True
        ElseIf isWinGame() Then
            winFlag = True
        End If
    End Sub
    Private Function buttonRightClick(ByVal x As Integer, ByVal y As Integer) As Boolean
        If Not isOpenedGrid(x, y) Then
            If isMarkedGrid(x, y) Then unmarkGrid(x, y) Else markGrid(x, y)
        End If
        Return True
    End Function
    Private Function buttonLeftClick(ByVal x0 As Integer, ByVal y0 As Integer) As Boolean
        If isBomb(x0, y0) Then Return False
        If isOpenedGrid(x0, y0) Then Return True
        unmarkGrid(x0, y0)
        If isSpace(x0, y0) Then
            fillSpace(x0, y0)
        End If
        If Not isOpenedGrid(x0, y0) Then
            openGrid(x0, y0)
        End If
        Return True
    End Function
    Private Function buttonDoubleClick(ByVal x0 As Integer, ByVal y0 As Integer) As Boolean
        If isNumber(x0, y0) And isOpenedGrid(x0, y0) Then
            If getNumber(x0, y0) = countBomb(x0, y0, True) Then
                Dim locations = New Point() {New Point(-1, 0), New Point(+1, 0), New Point(0, -1), New Point(0, +1), New Point(-1, -1), New Point(+1, +1), New Point(+1, -1), New Point(-1, +1)}
                For Each location As Point In locations
                    Dim x = x0 + location.X
                    Dim y = y0 + location.Y
                    If x >= 0 And x < GRID And y >= 0 And y < GRID Then
                        If isBomb(x, y) And Not isMarkedGrid(x, y) Then
                            Return False
                        End If
                        If isSpace(x, y) And Not isOpenedGrid(x, y) Then
                            fillSpace(x, y)
                        ElseIf Not isBomb(x, y) Then
                            openGrid(x, y)
                        End If
                    End If
                Next
            End If
        End If
        Return True
    End Function
    Private btnArray(GRID, GRID) As Button
    Function initButtons()
        For y = 0 To GRID - 1
            For x = 0 To GRID - 1
                btnArray(x, y) = New Button()
                btnArray(x, y).Width = BUTTON_SIZE
                btnArray(x, y).Height = BUTTON_SIZE
                btnArray(x, y).Location = New Point(x * BUTTON_SIZE, y * BUTTON_SIZE)
                btnArray(x, y).Font = New Font(New FontFamily("Arial"), 11)
                btnArray(x, y).ForeColor = Color.Red
                btnArray(x, y).TabStop = 0
                Me.Controls.Add(btnArray(x, y))
                AddHandler btnArray(x, y).MouseDown, AddressOf buttonMouseDown
                AddHandler btnArray(x, y).MouseUp, AddressOf buttonMouseUp
            Next
        Next
    End Function
    Private Sub MineSweeper_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ClientSize = New Size(BUTTON_SIZE * GRID, BUTTON_SIZE * GRID)
        initButtons()
        initGame()
    End Sub
End Class
