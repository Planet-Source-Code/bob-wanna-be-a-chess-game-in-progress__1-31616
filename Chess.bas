Attribute VB_Name = "Module1"
Public Const WhiteKing = 0, WhiteQueen = 1, _
WhitePawn1 = 2, WhiteKnight1 = 10, _
WhiteRook1 = 12, WhiteBish1 = 14
Public Const BlackKing = 16, BlackQueen = 17, _
BlackPawn1 = 18, BlackKnight1 = 26, _
BlackRook1 = 28, BlackBish1 = 30
Public Const King = 1, Queen = 2, Pawn = 3, _
Knight = 4, Rook = 5, Bishop = 6
Public Type ChessBoard
Squares(0 To 8, 0 To 8) As Integer
End Type
Public Type Pieces
Square(1 To 2) As Integer
T As Integer
Moves As Integer
NExist As Boolean
End Type
Public P(0 To 32) As Pieces
Public Board As ChessBoard
Public Taken(0 To 32) As Integer
Public KV, KH, Turn, O
Public DestX, DestY, LastBM, LastWM
Sub ResetP()
Form1.Piece(WhiteKing) = vbCrLf & "King"
Form1.Piece(WhiteQueen) = vbCrLf & "Queen"
For i = 0 To 7
   Form1.Piece(WhitePawn1 + i) = vbCrLf & "Pawn"
Next i
Form1.Piece(WhiteKnight1) = vbCrLf & "Knight"
Form1.Piece(WhiteKnight1 + 1) = vbCrLf & "Knight"
Form1.Piece(WhiteRook1) = vbCrLf & "Rook"
Form1.Piece(WhiteRook1 + 1) = vbCrLf & "Rook"
Form1.Piece(WhiteBish1) = vbCrLf & "Bishop"
Form1.Piece(WhiteBish1 + 1) = vbCrLf & "Bishop"

Form1.Piece(BlackKing) = vbCrLf & "King"
Form1.Piece(BlackQueen) = vbCrLf & "Queen"
For i = 0 To 7
   Form1.Piece(BlackPawn1 + i) = vbCrLf & "Pawn"
Next i
Form1.Piece(BlackKnight1) = vbCrLf & "Knight"
Form1.Piece(BlackKnight1 + 1) = vbCrLf & "Knight"
Form1.Piece(BlackRook1) = vbCrLf & "Rook"
Form1.Piece(BlackRook1 + 1) = vbCrLf & "Rook"
Form1.Piece(BlackBish1) = vbCrLf & "Bishop"
Form1.Piece(BlackBish1 + 1) = vbCrLf & "Bishop"
For i = 0 To 31
   If i < 16 Then Form1.Piece(i).BackColor = vbWhite
   If i >= 16 Then Form1.Piece(i).BackColor = vbWhite
Next
End Sub
Sub SetPosAll()
Dim S As Object
Dim S1 As Object

Form1.SetPos WhitePawn1, 0, 6
For i = 1 To 7
   P(WhitePawn1 + i).T = Pawn
   P(WhitePawn1 + i).Moves = 1
   Form1.SetPos WhitePawn1 + i, i, 6
Next
Form1.SetPos WhiteKnight1, 6, 7
Form1.SetPos WhiteKnight1 + 1, 1, 7
Form1.SetPos WhiteKing, 3, 7
Form1.SetPos WhiteQueen, 4, 7
Form1.SetPos WhiteRook1, 0, 7
Form1.SetPos WhiteRook1 + 1, 7, 7
Form1.SetPos WhiteBish1, 2, 7
Form1.SetPos WhiteBish1 + 1, 5, 7

Form1.SetPos BlackPawn1, 0, 1
For i = 1 To 7
   P(BlackPawn1 + i).T = Pawn
   P(BlackPawn1 + i).Moves = 1
   Form1.SetPos BlackPawn1 + i, i, 1
Next
Form1.SetPos BlackKnight1, 6, 0
Form1.SetPos BlackKnight1 + 1, 1, 0
Form1.SetPos BlackKing, 3, 0
Form1.SetPos BlackQueen, 4, 0
Form1.SetPos BlackRook1, 0, 0
Form1.SetPos BlackRook1 + 1, 7, 0
Form1.SetPos BlackBish1, 2, 0
Form1.SetPos BlackBish1 + 1, 5, 0
For i = 16 To 31
   Form1.Piece(i).BackColor = vbBlack
   Form1.Piece(i).ForeColor = vbWhite
Next
P(WhiteKnight1).T = Knight
P(WhiteKnight1 + 1).T = Knight
P(WhiteBish1).T = Bishop
P(WhiteBish1 + 1).T = Bishop
P(WhiteKing).T = King
P(WhitePawn1).T = Pawn
P(WhiteQueen).T = Queen
P(WhiteRook1).T = Rook
P(WhiteRook1 + 1).T = Rook
P(BlackKnight1).T = Knight
P(BlackKnight1 + 1).T = Knight
P(BlackBish1).T = Bishop
P(BlackBish1 + 1).T = Bishop
P(BlackKing).T = King
P(BlackPawn1).T = Pawn
P(BlackQueen).T = Queen
P(BlackRook1).T = Rook
P(BlackRook1 + 1).T = Rook
Set S = Form1.Shape2(0)
Form1.Shape2(0).Left = 45
Form1.Show
For i = 0 To 31
   Form1.Shape2(i).Visible = True
Next
For i = 1 To 11
   Form1.Shape2(i).Left = (S.Left + (S.Width * i)) + (20 * i)
   Form1.Shape2(i).Top = S.Top
Next
XV = 0
For i = 12 To 23
   Form1.Shape2(i).Left = (S.Left + (S.Width * XV)) + (20 * XV)
   Form1.Shape2(i).Top = S.Top + S.Height + 45
   XV = XV + 1
Next
XV = 0
For i = 24 To 31
   Form1.Shape2(i).Left = (S.Left + (S.Width * XV) + 1000) + (20 * XV)
   Form1.Shape2(i).Top = S.Top + ((S.Height + 45) * 2)
   XV = XV + 1
Next
End Sub

Sub DisableNotTurn()
On Error Resume Next
If (O < 16) Then Form1.Piece(O).ForeColor = vbBlack
If (O > 15) Then Form1.Piece(O).ForeColor = vbWhite
If (O < 16) Then Form1.Piece(O).BackColor = vbWhite
If (O > 15) Then Form1.Piece(O).BackColor = vbBlack

If Turn = 0 Then
      Min = 16
      Max = 31
      MN2 = 0
      MX2 = 15
      DestX = 0
      DestY = 0
      LastWM = O
      O = -1
      If P(LastBM).Square(1) <> -1 Then O = LastBM
   Else
      Min = 0
      Max = 15
      MN2 = 16
      MX2 = 31
      DestX = 0
      DestY = 0
      LastBM = O
      O = -1
      If P(LastBM).Square(1) <> -1 Then O = LastWM
End If
For i = Min To Max
   Form1.Piece(i).Enabled = False
Next
For i = MN2 To MX2
   Form1.Piece(i).Enabled = True
Next
If O <> -1 Then Form1.Piece_Click (O)
End Sub

Public Function IW(Num)
IW = False
If Num < 16 Then IW = True
End Function
Public Function IB(Num)
IB = False
If Num > 15 Then IB = True
End Function

