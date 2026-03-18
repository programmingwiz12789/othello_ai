Public Class Othello
    Dim whitePiece As Image = Image.FromFile("white_piece.png")
    Dim blackPiece As Image = Image.FromFile("black_piece.png")
    Dim n As Integer = 8
    Dim cells(n, n) As Button
    Dim board(n, n) As Integer
    Dim over As Boolean

    Private Sub Othello_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i = 0 To n - 1
            For j = 0 To n - 1
                cells(i, j) = DirectCast(Controls.Find("cell" & i & j, False).First, Button)
                board(i, j) = 0
            Next
        Next
        For i = n \ 2 - 1 To n \ 2
            For j = n \ 2 - 1 To n \ 2
                If i = j Then
                    board(i, j) = 1
                Else
                    board(i, j) = -1
                End If
            Next
        Next
        UpdateBoardState(n, cells, board)
        over = False
    End Sub

    Private Function IsLegalMove(n As Integer, board(,) As Integer, row As Integer, col As Integer, piece As Integer)
        If board(row, col) = 0 Then
            Dim dy() As Integer = {-1, -1, 0, 1, 1, 1, 0, -1}
            Dim dx() As Integer = {0, 1, 1, 1, 0, -1, -1, -1}
            For i = 0 To 7
                If row + dy(i) >= 0 And row + dy(i) <= n - 1 And col + dx(i) >= 0 And col + dx(i) <= n - 1 Then
                    If board(row + dy(i), col + dx(i)) = -piece Then
                        Dim r As Integer = row + dy(i), c As Integer = col + dx(i)
                        Do While r >= 0 And r <= n - 1 And c >= 0 And c <= n - 1
                            If board(r, c) = -piece Then
                                r += dy(i)
                                c += dx(i)
                                Continue Do
                            End If
                            Exit Do
                        Loop
                        If r >= 0 And r <= n - 1 And c >= 0 And c <= n - 1 Then
                            If board(r, c) = piece Then
                                Return True
                            End If
                        End If
                    End If
                End If
            Next
        End If
        Return False
    End Function

    Private Sub MovePiece(n As Integer, board(,) As Integer, row As Integer, col As Integer, piece As Integer)
        Dim dy() As Integer = {-1, -1, 0, 1, 1, 1, 0, -1}
        Dim dx() As Integer = {0, 1, 1, 1, 0, -1, -1, -1}
        board(row, col) = piece
        For i = 0 To 7
            If row + dy(i) >= 0 And row + dy(i) <= n - 1 And col + dx(i) >= 0 And col + dx(i) <= n - 1 Then
                If board(row + dy(i), col + dx(i)) = -piece Then
                    Dim r As Integer = row + dy(i), c As Integer = col + dx(i)
                    Do While r >= 0 And r <= n - 1 And c >= 0 And c <= n - 1
                        If board(r, c) = -piece Then
                            r += dy(i)
                            c += dx(i)
                            Continue Do
                        End If
                        Exit Do
                    Loop
                    If r >= 0 And r <= n - 1 And c >= 0 And c <= n - 1 Then
                        If board(r, c) = piece Then
                            Dim currR As Integer = row + dy(i), currC As Integer = col + dx(i)
                            Do While currR <> r Or currC <> c
                                board(currR, currC) = piece
                                currR += dy(i)
                                currC += dx(i)
                            Loop
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub UpdateBoardState(n As Integer, cells(,) As Button, board(,) As Integer)
        For i = 0 To n - 1
            For j = 0 To n - 1
                If board(i, j) = -1 Then
                    cells(i, j).Image = whitePiece
                ElseIf board(i, j) = 1 Then
                    cells(i, j).Image = blackPiece
                Else
                    cells(i, j).Image = Nothing
                End If
            Next
        Next
    End Sub

    Private Function CanMove(n As Integer, board(,) As Integer, piece As Integer)
        For i = 0 To n - 1
            For j = 0 To n - 1
                If IsLegalMove(n, board, i, j, piece) Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function

    Private Function CountPiece(n As Integer, board(,) As Integer, piece As Integer)
        Dim pieceCnt As Integer = 0
        For i = 0 To n - 1
            For j = 0 To n - 1
                If board(i, j) = piece Then
                    pieceCnt += 1
                End If
            Next
        Next
        Return pieceCnt
    End Function

    Private Function GameOverState(n As Integer, board(,) As Integer)
        If Not CanMove(n, board, 1) And Not CanMove(n, board, -1) Then
            Dim pieceCntDiff As Integer = CountPiece(n, board, 1) - CountPiece(n, board, -1)
            If pieceCntDiff = 0 Then
                Return 0
            End If
            Return pieceCntDiff \ Math.Abs(pieceCntDiff)
        End If
        Return 1000
    End Function

    Private Function SBE(n As Integer, board(,) As Integer)
        If GameOverState(n, board) = 1 Then
            Return 1000000
        ElseIf GameOverState(n, board) = -1 Then
            Return -1000000
        ElseIf GameOverState(n, board) = 0 Then
            Return 0
        Else
            'Evaluation Components Constants
            Dim STABILITY_WEIGHT As Integer = 100
            Dim MOBILITY_WEIGHT As Integer = 30
            Dim DISC_DIFFERENCE_WEIGHT As Integer = 5
            Dim STABILITY_SCORES(,) As Integer = {
                {100, -30, 6, 2, 2, 6, -30, 100},
                {-30, -50, 0, 0, 0, 0, -30, -50},
                {6, 0, 0, 0, 0, 0, 0, 6},
                {2, 0, 0, 0, 0, 0, 0, 2},
                {2, 0, 0, 0, 0, 0, 0, 2},
                {6, 0, 0, 0, 0, 0, 0, 6},
                {-30, -50, 0, 0, 0, 0, -30, -50},
                {100, -30, 6, 2, 2, 6, -30, 100}
            }

            'Evaluation Components Variables
            Dim stability As Integer = 0
            Dim mobility As Integer = 0
            Dim diskDiff As Integer = 0

            'Calculate evaluation score for each component
            For i = 0 To n - 1
                For j = 0 To n - 1
                    If board(i, j) = 1 Then
                        'Stability
                        stability += STABILITY_SCORES(i, j)

                        'Disk Difference
                        diskDiff += 1
                    ElseIf board(i, j) = -1 Then
                        'Disk Difference
                        diskDiff -= 1
                    End If

                    'Mobility
                    If IsLegalMove(n, board, i, j, 1) Then
                        mobility += 1
                    End If
                Next
            Next

            'Return final evaluation score by combining all components along with
            'their weights
            Return STABILITY_WEIGHT * stability + MOBILITY_WEIGHT * mobility + DISC_DIFFERENCE_WEIGHT * diskDiff
        End If
    End Function

    Private Function Minimax(depth As Integer, n As Integer, board(,) As Integer, isMax As Boolean, alpha As Integer, beta As Integer)
        If depth = 2 Or GameOverState(n, board) <> 1000 Then
            Dim score As Integer = SBE(n, board)
            Return score - depth
        Else
            If isMax Then
                Dim bestScore As Integer = -10000000
                If Not CanMove(n, board, 1) Then
                    Return Minimax(depth + 1, n, board, Not isMax, alpha, beta)
                End If
                For i = 0 To n - 1
                    For j = 0 To n - 1
                        If IsLegalMove(n, board, i, j, 1) Then
                            Dim temp(n, n) As Integer
                            For r = 0 To n - 1
                                For c = 0 To n - 1
                                    temp(r, c) = board(r, c)
                                Next
                            Next
                            MovePiece(n, board, i, j, 1)
                            bestScore = Math.Max(bestScore, Minimax(depth + 1, n, board, Not isMax, alpha, beta))
                            For r = 0 To n - 1
                                For c = 0 To n - 1
                                    board(r, c) = temp(r, c)
                                Next
                            Next
                            alpha = Math.Max(alpha, bestScore)
                            If alpha >= beta Then
                                GoTo isMaxExit
                            End If
                        End If
                    Next
                Next
isMaxExit:      Return bestScore
            Else
                Dim bestScore As Integer = 10000000
                If Not CanMove(n, board, -1) Then
                    Return Minimax(depth + 1, n, board, Not isMax, alpha, beta)
                End If
                For i = 0 To n - 1
                    For j = 0 To n - 1
                        If IsLegalMove(n, board, i, j, -1) Then
                            Dim temp(n, n) As Integer
                            For r = 0 To n - 1
                                For c = 0 To n - 1
                                    temp(r, c) = board(r, c)
                                Next
                            Next
                            MovePiece(n, board, i, j, -1)
                            bestScore = Math.Min(bestScore, Minimax(depth + 1, n, board, Not isMax, alpha, beta))
                            For r = 0 To n - 1
                                For c = 0 To n - 1
                                    board(r, c) = temp(r, c)
                                Next
                            Next
                            beta = Math.Min(beta, bestScore)
                            If alpha >= beta Then
                                GoTo isMinExit
                            End If
                        End If
                    Next
                Next
isMinExit:      Return bestScore
            End If
        End If
    End Function

    Private Function FindBestMove(n As Integer, board(,) As Integer)
        Dim bestMove As Integer = -1, bestScore As Integer = -100000000
        For i = 0 To n - 1
            For j = 0 To n - 1
                If IsLegalMove(n, board, i, j, 1) Then
                    Dim temp(n, n) As Integer
                    For r = 0 To n - 1
                        For c = 0 To n - 1
                            temp(r, c) = board(r, c)
                        Next
                    Next
                    MovePiece(n, board, i, j, 1)
                    Dim score As Integer = Minimax(0, n, board, False, -1000, 1000)
                    For r = 0 To n - 1
                        For c = 0 To n - 1
                            board(r, c) = temp(r, c)
                        Next
                    Next
                    If score > bestScore Then
                        bestScore = score
                        bestMove = i * n + j
                    End If
                End If
            Next
        Next
        Return bestMove
    End Function

    Private Sub OppMove(n As Integer, cells(,) As Button, board(,) As Integer)
        Dim bestMove As Integer = FindBestMove(n, board)
        If bestMove <> -1 Then
            Dim bestRow As Integer = bestMove \ n, bestCol As Integer = bestMove Mod n
            MovePiece(n, board, bestRow, bestCol, 1)
            UpdateBoardState(n, cells, board)
            If GameOverState(n, board) = 1 Then
                GameOver()
                MessageBox.Show("Opponent wins!")
            ElseIf GameOverState(n, board) = -1 Then
                GameOver()
                MessageBox.Show("You win!")
            ElseIf GameOverState(n, board) = 0 Then
                GameOver()
                MessageBox.Show("Draw!")
            End If
        End If
    End Sub

    Private Sub GameOver()
        over = True
    End Sub

    Private Sub button_Click(sender As Object, e As EventArgs) Handles cell77.Click, cell76.Click, cell75.Click, cell74.Click, cell73.Click, cell72.Click, cell71.Click, cell70.Click, cell67.Click, cell66.Click, cell65.Click, cell64.Click, cell63.Click, cell62.Click, cell61.Click, cell60.Click, cell57.Click, cell56.Click, cell55.Click, cell54.Click, cell53.Click, cell52.Click, cell51.Click, cell50.Click, cell47.Click, cell46.Click, cell45.Click, cell44.Click, cell43.Click, cell42.Click, cell41.Click, cell40.Click, cell37.Click, cell36.Click, cell35.Click, cell34.Click, cell33.Click, cell32.Click, cell31.Click, cell30.Click, cell27.Click, cell26.Click, cell25.Click, cell24.Click, cell23.Click, cell22.Click, cell21.Click, cell20.Click, cell17.Click, cell16.Click, cell15.Click, cell14.Click, cell13.Click, cell12.Click, cell11.Click, cell10.Click, cell07.Click, cell06.Click, cell05.Click, cell04.Click, cell03.Click, cell02.Click, cell01.Click, cell00.Click
        If Not over Then
            Dim cellName As String = CType(sender, Button).Name
            Dim cellNum As Integer = CType(cellName.Substring(4), Integer)
            Dim row = cellNum \ 10, col = cellNum Mod 10
            If IsLegalMove(n, board, row, col, -1) Then
                MovePiece(n, board, row, col, -1)
                UpdateBoardState(n, cells, board)
                If GameOverState(n, board) = -1 Then
                    GameOver()
                    MessageBox.Show("You win!")
                ElseIf GameOverState(n, board) = 1 Then
                    GameOver()
                    MessageBox.Show("Opponent wins!")
                ElseIf GameOverState(n, board) = 0 Then
                    GameOver()
                    MessageBox.Show("Draw!")
                Else
                    OppMove(n, cells, board)
                    Do While GameOverState(n, board) = 1000 And Not CanMove(n, board, -1)
                        OppMove(n, cells, board)
                    Loop
                End If
            End If
        End If
    End Sub

    Private Sub restartBtn_Click(sender As Object, e As EventArgs) Handles restartBtn.Click
        For i = 0 To n - 1
            For j = 0 To n - 1
                board(i, j) = 0
            Next
        Next
        For i = n \ 2 - 1 To n \ 2
            For j = n \ 2 - 1 To n \ 2
                If i = j Then
                    board(i, j) = 1
                Else
                    board(i, j) = -1
                End If
            Next
        Next
        UpdateBoardState(n, cells, board)
        over = False
    End Sub
End Class
