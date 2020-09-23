Attribute VB_Name = "modEngine"
Public GameError As Boolean '// gameerror variable
'// this is where the map settings will be kept
Dim KeyRunning As Boolean
Dim MapData() As String
Dim CharData() As String
Dim FurniData() As String
'// the width of the map
Dim LengthX As Integer
'// tiles x,y axis points variable
Public TilesX() As Long
Public TilesY() As Long
'// animation running value
Public AniRunning As Boolean
'// speed of the animation
Public FrameSpeed As Single
'// the record of the lines that will be used to construct the walls
Dim WallLinesX(1 To 6) As Integer
Dim WallLinesY(1 To 6) As Integer
Dim RecordWalls(1 To 3, 1 To 2) As Integer
'// users account name
Public AccountName As String
'// message variables
Dim Messages As Integer
Dim FreeMessage As Integer
'// users chars info x,y points and width and height of the images used
Dim CharX As Long
Dim CharY As Long
Dim CharHeight As Long
Dim CharWidth As Long
'// furniture x,y points
Dim FurniX() As Long
Dim FurniY() As Long
'// how many tiles used on the map
Dim Tiles As Integer
'// wallcolor(wallpaper color)
Public WallColor As Long
'// direction of the user's char
Dim PthJumprActv As Boolean
Dim Drink As Boolean
Private Enum Directions
    Right = 4
    Left = 3
    down = 2
    Up = 1
End Enum
Dim Direction As Directions
'// graphic API functions
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Sub InitData(MData() As String, CData() As String, FData() As String)
On Error GoTo MsgError
'// Loads the map data
Dim Count As Integer
    ReDim MapData(1 To UBound(MData()), 1 To Len(MData(1)))
    ReDim CharData(1 To UBound(CData()), 1 To Len(CData(1)))
    ReDim FurniData(1 To UBound(FData()), 1 To Len(FData(1)))
    LengthX = Len(MData(1)) '//width of the map
    '//this inputs all the data
    For Y = 1 To UBound(MData())
        Count = Count + 1
        For X = 1 To Len(MData(1))
            MapData(Y, X) = Mid$(MData(Count), X, 1)
            CharData(Y, X) = Mid$(CData(Count), X, 1)
            FurniData(Y, X) = Mid$(FData(Count), X, 1)
        Next X
    Next Y
    '//load the images
    Load Memory
    Memory.Hide
    '// init the speed of the animation
    FrameSpeed = 0.0005
    frmMain.frmOptions.Left = 7
    frmMain.frmOptions.Top = 25
    Drink = False
    GameError = False
    Exit Sub
MsgError:
    '// common error description
    MsgBox "One of the settings don't fit the other settings. Check if the length and width is the same!"
    GameError = True
End Sub
Sub CreateRoom()
    '// calls out all the subs that are needed to create the map
    Call CreateTiles
    Call CreateWalls
    '// keep the original just tile and wall image just incase we need to reference
    frmMain.picTileBck.Picture = frmMain.Picture1.Image
    Call CreateFurni
    '// keep the room image just incase we need to reference to it and which we do
    frmMain.PicRealBck.Picture = frmMain.Picture1.Image
    Call CreateChars
    Call CreateFurni
    '// all process done show the UI(User Interface)
    frmMain.Show
End Sub
Sub CreateChars()
'// creates the users interface character
Dim Tile As Integer
    '//loops through the map trying to find where the character should be placed
    For Y = 1 To UBound(CharData()) '1 to height of the map
        For X = 1 To LengthX '//1 to the width of the map
            If CharData(Y, X) = "O" Then 'if this is where the character is then
                If Y = 1 Then '//if y =1 then place it on the tile that it was founded on
                    Tile = X
                Else
                    '//else this could be used a function to get which tile the user is on
                    Tile = FindTile(Y, X)
                End If
                '//creates the x and the y where the users image will be placed
                CharX = TilesX(Tile) + 7
                CharY = TilesY(Tile) - (Memory.picBodyRight.Height - Memory.picTile1.Height) - 10
                '//places the character
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRightMask.Width, Memory.picBodyRightMask.Height, Memory.picBodyRightMask.hdc, 0, 0, vbSrcAnd
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRight.Width, Memory.picBodyRight.Height, Memory.picBodyRight.hdc, 0, 0, vbSrcPaint
                '//updates the direction
                CharWidth = Memory.picBodyRight.Width
                CharHeight = Memory.picBodyRight.Height
                Direction = Right
            End If
        Next X
    Next Y
End Sub
Sub CreateTiles()
'// creates tiles and places them
On Error Resume Next
Dim CenterX As Integer
Dim CenterY As Integer
Dim TileX As Integer
Dim TileY As Integer
Dim WallStop As Integer
Dim WallChange As Boolean
    '// load the form but hide it untill everything it loaded
    Load frmMain
    frmMain.Hide

    Tiles = UBound(MapData()) * LengthX '// gets how many tiles theres going to be in the room
    ReDim TilesX(1 To Tiles) As Long '//redims to the qualified info
    ReDim TilesY(1 To Tiles) As Long
    '//creates the middle starting point where the tiles will start getting drawn
    CenterX = ((frmMain.Picture1.Width - Memory.picTile1.Width) / 2) - (Memory.picTile1.Width / 2)
    CenterY = ((frmMain.Picture1.Height - (UBound(MapData()) * Memory.picTile1.Height)) / 2) - (Memory.picTile1.Height / 2) + 100
    '//initiliaze the first starting point
    TileY = CenterY
    TileX = CenterX
    Tiles = 0
    '//WallStop = UBound(MapData())
    '//start looping through the tile info
    For Y = 1 To UBound(MapData())
        '//this is the points where the next tile will be drawn
        TileY = CenterY + (Memory.picTile1.Height / 2 - 4) * Y
        TileX = CenterX - (Memory.picTile1.Width / 2 - 3) * Y
        
        For X = 1 To LengthX
            '//adds on to the tiles x position to make it look 3d
            TileX = TileX + (Memory.picTile1.Width / 2) - 3
            '//if mapdata is not a dead spot then paste the image
            If MapData(Y, X) = "O" Then
                BitBlt frmMain.Picture1.hdc, TileX, TileY, Memory.PicTile1Mask.Width, Memory.PicTile1Mask.Height, Memory.PicTile1Mask.hdc, 0, 0, vbSrcAnd
                BitBlt frmMain.Picture1.hdc, TileX, TileY, Memory.picTile1.Width, Memory.picTile1.Height, Memory.picTile1.hdc, 0, 0, vbSrcPaint
            End If
            '//if the next stop is a dead spot make sure to record the wall spot
            If MapData(Y + 1, 1) <> "O" Then
                WallStop = Y
            End If
            
            '//update tiles info
            Tiles = Tiles + 1
            TilesX(Tiles) = TileX
            TilesY(Tiles) = TileY
            TileY = TileY + (Memory.picTile1.Height / 2) - 4
            
            '//record the walls info
            If Y = WallStop And X = 1 And WallChange = False Then
                '// this records the first point of the wall on the left bottom
                '//this if statement executes only if the wall had ended and
                '// the height of the room didnt
                RecordWalls(1, 1) = TileX + 1
                RecordWalls(1, 2) = TileY
                '//since the height isnt done but we did record the wall spot
                '// i make a dummy variable true so this if statement doesnt
                '// execute anymore since there can be more dead spots as the y
                '// increases
                If WallStop <> UBound(MapData()) Then
                    WallChange = True
                End If
            ElseIf Y = 1 And X = 1 Then
                '//records the middle point of the wall
                RecordWalls(2, 1) = TileX + Memory.picTile1.Width / 2 - 2
                RecordWalls(2, 2) = TileY - Memory.picTile1.Height / 2 + 4
            ElseIf Y = 1 And X = LengthX Then
                '//recprds the right bottom point of the wall
                RecordWalls(3, 1) = TileX + Memory.picTile1.Width - 5
                RecordWalls(3, 2) = TileY
            End If
        Next X
    Next Y
End Sub
Sub CreateFurni()
'// creates and places furni
    '// redims the furni info to the qualified info
    ReDim FurniX(1 To Tiles) As Long
    ReDim FurniY(1 To Tiles) As Long
    '//loops through the furni info
    For Y = 1 To UBound(FurniData())
        For X = 1 To LengthX
            '//count is the variable that would be used for which furni tile is on
            '//so basicly when count+= thats the same thing as adding to a tile
            '//if that clears things up
            Count = Count + 1
            '// checks the char codes then displays the furniture
            If FurniData(Y, X) = "D" Then
                '//records the x and y point of the furni where it was placed
                FurniX(Count) = TilesX(Count) - 4
                FurniY(Count) = TilesY(Count) - Memory.picMountinDew.Height + Memory.picTile1.Height - 3
                '//draws the corresponding furniture
                Call BitBlt(frmMain.Picture1.hdc, FurniX(Count), FurniY(Count), Memory.picMountinDewMask.Width, Memory.picMountinDewMask.Height, Memory.picMountinDewMask.hdc, 0, 0, vbSrcAnd)
                Call BitBlt(frmMain.Picture1.hdc, FurniX(Count), FurniY(Count), Memory.picMountinDew.Width, Memory.picMountinDew.Height, Memory.picMountinDew.hdc, 0, 0, vbSrcPaint)
            ElseIf FurniData(Y, X) = "I" Then
                FurniX(Count) = TilesX(Count)
                FurniY(Count) = TilesY(Count) - Memory.picItemBox.Height + Memory.picTile1.Height - 3
                Call BitBlt(frmMain.Picture1.hdc, FurniX(Count), FurniY(Count), Memory.picItemBoxMask.Width, Memory.picItemBoxMask.Height, Memory.picItemBoxMask.hdc, 0, 0, vbSrcAnd)
                Call BitBlt(frmMain.Picture1.hdc, FurniX(Count), FurniY(Count), Memory.picItemBox.Width, Memory.picItemBox.Height, Memory.picItemBox.hdc, 0, 0, vbSrcPaint)
            End If
        Next X
    Next Y
End Sub
Sub MoveBody(MoveLeft As Boolean, MoveRight As Boolean, MoveUp As Boolean, MoveDown As Boolean, Optional MyBody As Boolean)
On Error Resume Next '// just incase an error happens continue with the code
Dim FrameCount As Integer
Dim FrameSpeedX As Single
Dim FrameSpeedY As Single
Dim Frame As Single
Dim Frames As Integer
Dim Tile As Integer
    Frames = 3 '//we need three frames to finish the whole movement
    FrameCount = 0 '//init to 0 frames
    FrameSpeedY = 18 / Frames '//we need to travel 25 pixels up/down on the y axis
    '//since we have 3 frames we divide 25/3 and get how many pixels we need to travel
    FrameSpeedX = 6 / Frames '//same here only for the x axis
    If MyBody = True Then '//we need to check if its the user or a diff user on the net
        For Y = 1 To UBound(CharData()) '// loop through the char data to find where
        '// our user is standing
            For X = 1 To LengthX
                If AniRunning = False Then '// if other animation isnt running then run
                    '// animation
                    '// check what movement should be executed but before that
                    '// you need to know where the user is so that checks that to
                    If MoveRight = True And CharData(Y, X) = "O" Then
                        AniRunning = True '// animation is running
                        If X + 1 <= LengthX Then '// checks if the user isnt out of boundary
                            If CanMove(X + 1, Y) = False Then '//collision detection
                                '//if can move = false then on the right theres either an object or furniture
                                '//so now since were moving to the right if the user is already in the
                                '// right direction let him stay that way and not change anything
                                '// but if its a diff direction
                                If Direction <> Right Then
                                    '// we then erase the current image of the user
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                    '// then we paste the right image of the user
                                    '// but first check which image to paste
                                    If Drink = False Then
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRightMask.Width, Memory.picBodyRightMask.Height, Memory.picBodyRightMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRight.Width, Memory.picBodyRight.Height, Memory.picBodyRight.hdc, 0, 0, vbSrcPaint
                                    Else
                                        '// paste the image with the drink
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRightMask.Width, Memory.picBodyDrinkRightMask.Height, Memory.picBodyDrinkRightMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRight.Width, Memory.picBodyDrinkRight.Height, Memory.picBodyDrinkRight.hdc, 0, 0, vbSrcPaint
                                    End If
                                    '// update the drection
                                    Direction = Right
                                    '// since we changed the image there can be a
                                    '// diff height and width and those things we need
                                    '// to know for other calculation so we update it
                                    Call UpdateChar
                                    '// this will check if the user is behind any
                                    '// furni and then processes commands to create the
                                    '// illusion that the user is behind the furniture
                                    Call BehindFurni(X, Y)
                                    '// now just incase sometimes the picture does now
                                    '// show what we cut out or pasted on so we refresh
                                    frmMain.Picture1.Refresh
                                End If
                                AniRunning = False '// now all the processing of the animation is done
                                Exit Sub '// were done we dont need to go any further so we exit the sub
                            End If
                            '//if theres nothing on the right we need to update the charsdata()
                            '// we replace where the user was before with an x meaning nothing there
                            '// and now we update the spot on the right with an O which is the char code
                            '// for the user
                            CharData(Y, X) = "X"
                            CharData(Y, X + 1) = "O"
                            '// get the number of the right tile
                            Tile = FindTile(Y, X + 1)
                            '// this is block of code that does the animation
                            Do While FrameCount <= Frames '//checks if the animation is done else continue
                                Frame = Frame + FrameSpeed '//add to the frame
                                If Frame >= 1 Then '// if frame is greater then 1 or =1 one then one frame is done and we
                                '// need to do the animation corresponding to that frame
                                    FrameCount = FrameCount + 1
                                    If FrameCount = 1 Then '// if frame =1 then
                                        '// cut out the image of the character and replace it with the image that was there before it
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// the users x axis will move up making the illusion of moving to the right
                                        CharX = TilesX(Tile) + FrameSpeedX
                                        '// the users y axis will move down making the illusion of moving to the right
                                        CharY = TilesY(Tile) - Memory.picBodyRight.Height + FrameSpeedY
                                        '// now we paste on the image in its new x axis and y axis points
                                        '// but check which image to paste
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRightMask.Width, Memory.picBodyRightMask.Height, Memory.picBodyRightMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRight.Width, Memory.picBodyRight.Height, Memory.picBodyRight.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRightMask.Width, Memory.picBodyDrinkRightMask.Height, Memory.picBodyDrinkRightMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRight.Width, Memory.picBodyDrinkRight.Height, Memory.picBodyDrinkRight.hdc, 0, 0, vbSrcPaint
                                        End If
                                        '// update the direction
                                        Direction = Right
                                        '// update the char's height and width
                                        Call UpdateChar
                                    Else
                                        '// framecount =1 is the most important that gives us the starting point
                                        '// the next frames we just move add on
                                        '// remove the image and replace it with the image that was there before it
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// make the new points that will create the illusion of the movement
                                        CharX = CharX + FrameSpeedX
                                        CharY = CharY + FrameSpeedY
                                        '// paste the image in its new x and y points
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRightMask.Width, Memory.picBodyRightMask.Height, Memory.picBodyRightMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRight.Width, Memory.picBodyRight.Height, Memory.picBodyRight.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRightMask.Width, Memory.picBodyDrinkRightMask.Height, Memory.picBodyDrinkRightMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRight.Width, Memory.picBodyDrinkRight.Height, Memory.picBodyDrinkRight.hdc, 0, 0, vbSrcPaint
                                        End If
                                    End If
                                    Frame = 0 '//init the frame
                                End If
                                DoEvents '//command used to make the computer finish
                                '//up its execution so its doesnt fall back and doesnt
                                '//make the program crash
                            Loop
                            '// just incase that the user is behind the furni and
                            '// it doesnt look like he is we do this
                            Call BehindFurni(X + 1, Y)
                        Else '//else the person if the person does go to the right
                            '// it will be out of the boundary so we just
                            '// update the direction it should be facing
                            '// if its not in the right one
                            If Direction <> Right Then
                                '// cut out the image
                                BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                '// then paste the correct one
                                If Drink = False Then
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRightMask.Width, Memory.picBodyRightMask.Height, Memory.picBodyRightMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRight.Width, Memory.picBodyRight.Height, Memory.picBodyRight.hdc, 0, 0, vbSrcPaint
                                Else
                                    '// paste the image with the drink
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRightMask.Width, Memory.picBodyDrinkRightMask.Height, Memory.picBodyDrinkRightMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRight.Width, Memory.picBodyDrinkRight.Height, Memory.picBodyDrinkRight.hdc, 0, 0, vbSrcPaint
                                End If
                            End If
                            '// update the direction
                            Direction = Right
                            '// update the char's height and width
                            Call UpdateChar
                            '// check if the user is behind furni since we didnt move
                            '// we keep x and y
                            Call BehindFurni(X, Y)
                        End If
                    ElseIf MoveLeft = True And CharData(Y, X) = "O" Then
                        If X - 1 >= 1 Then '//checks if the user will go out of the boundary
                        '// if he goes to the left
                            AniRunning = True '//now animation is running
                            '//collision detection
                            If CanMove(X - 1, Y) = False Then
                                If Direction <> Left Then
                                    '// remove the image and restore the background
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                    '// paste the left direction image
                                    '// paste the correct image
                                    If Drink = False Then
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftMask.Width, Memory.picBodyLeftMask.Height, Memory.picBodyLeftMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeft.Width, Memory.picBodyLeft.Height, Memory.picBodyLeft.hdc, 0, 0, vbSrcPaint
                                    Else
                                        '// paste the image with the direction
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrinkMask.Width, Memory.picBodyLeftDrinkMask.Height, Memory.picBodyLeftDrinkMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrink.Width, Memory.picBodyLeftDrink.Height, Memory.picBodyLeftDrink.hdc, 0, 0, vbSrcPaint
                                    End If
                                    '// update the direction
                                    Direction = Left
                                    '// update the char's details
                                    Call UpdateChar
                                    '// check if the char is behind the furni
                                    Call BehindFurni(X, Y)
                                    '// refresh the pic to update
                                    frmMain.Picture1.Refresh
                                End If
                                AniRunning = False '//animation is done running
                                Exit Sub
                            End If
                            '// update the chardata
                            CharData(Y, X) = "X"
                            CharData(Y, X - 1) = "O" '//x-1 = 1 to the left
                            '// find the tile number to the left of the user's char
                            Tile = FindTile(Y, X - 1)
                            '//animation
                            Do While FrameCount <= Frames
                                '// add on to the frame
                                Frame = Frame + FrameSpeed
                                If Frame >= 1 Then
                                    FrameCount = FrameCount + 1
                                    If FrameCount = 1 Then '//start up frame
                                        '// cut out the image restore the cut with background
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// create the x,y points for where the image will be place
                                        '// that creates the illusion of going to the left
                                        CharX = TilesX(Tile) + FrameSpeedX
                                        CharY = TilesY(Tile) - Memory.picBodyLeft.Height + FrameSpeedY
                                        '// paste the image in the new points
                                        '// check which image to paste on
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftMask.Width, Memory.picBodyLeftMask.Height, Memory.picBodyLeftMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeft.Width, Memory.picBodyLeft.Height, Memory.picBodyLeft.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            CharX = CharX - 5
                                            CharY = CharY + 1
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrinkMask.Width, Memory.picBodyLeftDrinkMask.Height, Memory.picBodyLeftDrinkMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrink.Width, Memory.picBodyLeftDrink.Height, Memory.picBodyLeftDrink.hdc, 0, 0, vbSrcPaint
                                        End If
                                        '// update the direction
                                        Direction = Left
                                        '// update the char
                                        Call UpdateChar
                                    Else
                                        '// remove the image and restore the background
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// add on to the x,y points to make the illusion of going to the left
                                        CharX = CharX + FrameSpeedX
                                        CharY = CharY + FrameSpeedY
                                        '// copy the image on to the background
                                        '// check which imag to paste
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftMask.Width, Memory.picBodyLeftMask.Height, Memory.picBodyLeftMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeft.Width, Memory.picBodyLeft.Height, Memory.picBodyLeft.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrinkMask.Width, Memory.picBodyLeftDrinkMask.Height, Memory.picBodyLeftDrinkMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrink.Width, Memory.picBodyLeftDrink.Height, Memory.picBodyLeftDrink.hdc, 0, 0, vbSrcPaint
                                        End If
                                    End If
                                    Frame = 0 '// init frame
                                End If
                                DoEvents '// let comp do its execution
                            Loop
                            '// check if the char is behind the furni then
                            '// make the illusion that he/she is
                            Call BehindFurni(X - 1, Y)
                        Else '// out of boundary so just update the direction
                            If Direction <> Left Then
                                '// cut the image restore the background
                                BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                '// paste the image
                                '// paste the correct image
                                If Drink = False Then
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftMask.Width, Memory.picBodyLeftMask.Height, Memory.picBodyLeftMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeft.Width, Memory.picBodyLeft.Height, Memory.picBodyLeft.hdc, 0, 0, vbSrcPaint
                                Else
                                    '// paste the image with the drink
                                    CharX = CharX - 6
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrinkMask.Width, Memory.picBodyLeftDrinkMask.Height, Memory.picBodyLeftDrinkMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrink.Width, Memory.picBodyLeftDrink.Height, Memory.picBodyLeftDrink.hdc, 0, 0, vbSrcPaint
                                End If
                            End If
                            Direction = Left
                            '// update the char
                            Call UpdateChar
                            '// check if the user is behind furni
                            Call BehindFurni(X, Y)
                        End If
                    ElseIf MoveUp = True And CharData(Y, X) = "O" Then
                        If Y - 1 <> 0 Then '// check if the user can go out of the boundary
                            AniRunning = True '// animation is now in process
                            If CanMove(X, Y - 1) = False Then '// check collision
                                If Direction <> Up Then '// update the direction
                                    '// cut the image and restore with background
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                    '// paste the new image
                                    If Drink = False Then
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpMask.Width, Memory.picBodyUpMask.Height, Memory.picBodyUpMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUp.Width, Memory.picBodyUp.Height, Memory.picBodyUp.hdc, 0, 0, vbSrcPaint
                                    Else
                                        '// paste the image with the drink
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrinkMask.Width, Memory.picBodyUpDrinkMask.Height, Memory.picBodyUpDrinkMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrink.Width, Memory.picBodyUpDrink.Height, Memory.picBodyUpDrink.hdc, 0, 0, vbSrcPaint
                                    End If
                                    '// update the direction
                                    Direction = Up
                                    '// update the char's details
                                    Call UpdateChar
                                    '// check if the user is behind furni
                                    '// if so then it will make the illusion
                                    '// that it is
                                    Call BehindFurni(X, Y)
                                    '// refresh the picture to make it "uptodate"
                                    frmMain.Picture1.Refresh
                                End If
                                AniRunning = False '//animation has finished
                                Exit Sub
                            End If
                            '// update the chardata
                            CharData(Y, X) = "X"
                            CharData(Y - 1, X) = "O"
                            '// get the tile number
                            Tile = FindTile(Y, X - LengthX)
                            '// animation
                            Do While FrameCount <= Frames
                                Frame = Frame + FrameSpeed 'add on to the frame
                                If Frame >= 1 Then
                                    FrameCount = FrameCount + 1
                                    If FrameCount = 1 Then '//start up frame
                                        '// cut the image restore with background
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// add on or minus the char's x,y points so it looks like it moved up
                                        CharX = TilesX(Tile) + FrameSpeedX
                                        CharY = TilesY(Tile) - Memory.picBodyUp.Height + FrameSpeedY
                                        '// paste the new image with the new points
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpMask.Width, Memory.picBodyUpMask.Height, Memory.picBodyUpMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUp.Width, Memory.picBodyUp.Height, Memory.picBodyUp.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrinkMask.Width, Memory.picBodyUpDrinkMask.Height, Memory.picBodyUpDrinkMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrink.Width, Memory.picBodyUpDrink.Height, Memory.picBodyUpDrink.hdc, 0, 0, vbSrcPaint
                                        End If
                                        '// update the direction
                                        Direction = Up
                                        '// update the char details
                                        Call UpdateChar
                                    Else
                                        '// cut the image restore with background
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// add on to the points to make it look like its moving up
                                        CharX = CharX + FrameSpeedX
                                        CharY = CharY + FrameSpeedY
                                        '// paste the image on the new points
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpMask.Width, Memory.picBodyUpMask.Height, Memory.picBodyUpMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUp.Width, Memory.picBodyUp.Height, Memory.picBodyUp.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrinkMask.Width, Memory.picBodyUpDrinkMask.Height, Memory.picBodyUpDrinkMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrink.Width, Memory.picBodyUpDrink.Height, Memory.picBodyUpDrink.hdc, 0, 0, vbSrcPaint
                                        End If
                                    End If
                                    Frame = 0 '// init the frame
                                End If
                                DoEvents '// let the comp process its math
                            Loop
                            '// check if behind furni
                            Call BehindFurni(X, Y - 1)
                        Else '// if out of boundary then update the direction
                            '// checks if the the update is needed
                            If Direction <> Up Then
                                '// cut out the image restore the background
                                BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                '// paste the new image
                                If Drink = False Then
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpMask.Width, Memory.picBodyUpMask.Height, Memory.picBodyUpMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUp.Width, Memory.picBodyUp.Height, Memory.picBodyUp.hdc, 0, 0, vbSrcPaint
                                Else
                                    '// paste the image with the drink
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrinkMask.Width, Memory.picBodyUpDrinkMask.Height, Memory.picBodyUpDrinkMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrink.Width, Memory.picBodyUpDrink.Height, Memory.picBodyUpDrink.hdc, 0, 0, vbSrcPaint
                                End If
                            End If
                            '// update the direction
                            Direction = Up
                            '// update the char details
                            Call UpdateChar
                            '// check if its behind the furni then do some
                            '// math and process some commands to make it
                            '// look like it is behind it
                            Call BehindFurni(X, Y)
                        End If
                    ElseIf MoveDown = True And CharData(Y, X) = "O" Then
                        '// check if the user is ganna be out of the boundary
                        '// if moves down
                        If Y + 1 <= UBound(MapData()) Then
                            AniRunning = True '//animation is running
                            If CanMove(X, Y + 1) = False Then '// collision detection
                                '// since collision is detected the user is not moving
                                '// so the only thing we need to update is the direction
                                '// the user will be facing
                                If Direction <> down Then '// checks if the user is in the down direction
                                    '// cut the image and restore the background
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                    '// paste the image with the right direction
                                    If Drink = False Then
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDownMask.Width, Memory.picBodyDownMask.Height, Memory.picBodyDownMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDown.Width, Memory.picBodyDown.Height, Memory.picBodyDown.hdc, 0, 0, vbSrcPaint
                                    Else
                                        '// paste the image with the drink
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDownMask.Width, Memory.picBodyDrinkDownMask.Height, Memory.picBodyDrinkDownMask.hdc, 0, 0, vbSrcAnd
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDown.Width, Memory.picBodyDrinkDown.Height, Memory.picBodyDrinkDown.hdc, 0, 0, vbSrcPaint
                                    End If
                                    '// update the direction
                                    Direction = down
                                    '// update the char's detail
                                    Call UpdateChar
                                    '// check if the user is behind furni
                                    Call BehindFurni(X, Y)
                                    '// refresh the picture
                                    frmMain.Picture1.Refresh
                                End If
                                AniRunning = False '// animation is done running
                                Exit Sub
                            End If
                            '// update the charsmap data
                            CharData(Y, X) = "X"
                            CharData(Y + 1, X) = "O"
                            '// find the tile the one above the user's char
                            Tile = FindTile(Y + 1, X)
                            '// animation
                            Do While FrameCount <= Frames '// check if the animation is done
                                Frame = Frame + FrameSpeed '// add on to the frame
                                If Frame >= 1 Then '// see if the frame is full if so then check what to do
                                    FrameCount = FrameCount + 1
                                    If FrameCount = 1 Then '// if the framecount is one then we need to
                                        '// load the starting point
                                        '// remove the image and restore the background
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// create the starting points of the movement
                                        CharX = TilesX(Tile) + FrameSpeedX
                                        CharY = TilesY(Tile) - Memory.picBodyRight.Height + FrameSpeedY
                                        '// paste the image where it needs to be shown in its new x,y points
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDownMask.Width, Memory.picBodyDownMask.Height, Memory.picBodyDownMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDown.Width, Memory.picBodyDown.Height, Memory.picBodyDown.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the new image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDownMask.Width, Memory.picBodyDrinkDownMask.Height, Memory.picBodyDrinkDownMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDown.Width, Memory.picBodyDrinkDown.Height, Memory.picBodyDrinkDown.hdc, 0, 0, vbSrcPaint
                                        End If
                                        '// update the direction
                                        Direction = down
                                        '// update the char's detail
                                        Call UpdateChar
                                    Else
                                        '// remove the image and restore the background
                                        BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDown.Width, Memory.picBodyDown.Height, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                        '// add on to the x,y points making the illusion it's moving down
                                        CharX = CharX + FrameSpeedX
                                        CharY = CharY + FrameSpeedY
                                        '// paste the new image in the new x,y points
                                        If Drink = False Then
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDownMask.Width, Memory.picBodyDownMask.Height, Memory.picBodyDownMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDown.Width, Memory.picBodyDown.Height, Memory.picBodyDown.hdc, 0, 0, vbSrcPaint
                                        Else
                                            '// paste the image with the drink
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDownMask.Width, Memory.picBodyDrinkDownMask.Height, Memory.picBodyDrinkDownMask.hdc, 0, 0, vbSrcAnd
                                            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDown.Width, Memory.picBodyDrinkDown.Height, Memory.picBodyDrinkDown.hdc, 0, 0, vbSrcPaint
                                        End If
                                    End If
                                    Frame = 0 '// init the frame
                                End If
                                DoEvents '// let the comp process all of this math
                            Loop
                            '// check if the user is behind the furni
                            Call BehindFurni(X, Y + 1)
                        Else '// if out of boundary then update the direction
                            '// checks if the update is needed
                            If Direction <> down Then
                                '// cut the image restore the background
                                BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                                '// paste the new image on to the background
                                If Drink = False Then
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDownMask.Width, Memory.picBodyDownMask.Height, Memory.picBodyDownMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDown.Width, Memory.picBodyDown.Height, Memory.picBodyDown.hdc, 0, 0, vbSrcPaint
                                Else
                                    '// paste the image with the drink image
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDownMask.Width, Memory.picBodyDrinkDownMask.Height, Memory.picBodyDrinkDownMask.hdc, 0, 0, vbSrcAnd
                                    BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDown.Width, Memory.picBodyDrinkDown.Height, Memory.picBodyDrinkDown.hdc, 0, 0, vbSrcPaint
                                End If
                            End If
                            '// update the direction
                            Direction = down
                            '// update the char's details
                            Call UpdateChar
                            '// check if the user is behind the furni
                            Call BehindFurni(X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
    End If
    DoEvents
    AniRunning = False '// animation is done running
    frmMain.Picture1.Refresh '// refresh the picture and update the process that has been done
End Sub
Sub CreateWalls()
'// creates walls
    WallLinesX(1) = RecordWalls(1, 1)
    WallLinesY(1) = RecordWalls(1, 2)
    WallLinesX(2) = RecordWalls(2, 1)
    WallLinesY(2) = RecordWalls(2, 2)
    WallLinesX(3) = RecordWalls(3, 1)
    WallLinesY(3) = RecordWalls(3, 2)
    WallLinesX(4) = WallLinesX(1)
    WallLinesY(4) = WallLinesY(1) - 100
    WallLinesX(5) = WallLinesX(2)
    WallLinesY(5) = WallLinesY(2) - 100
    Call ShowWalls
    Call ColorWalls
End Sub
Sub ShowWalls()
'// show the wall lines
    frmMain.Picture1.Line (WallLinesX(1), WallLinesY(1))-(WallLinesX(1), WallLinesY(1) - 100), vbBlack
    frmMain.Picture1.Line (WallLinesX(2), WallLinesY(2))-(WallLinesX(2), WallLinesY(2) - 100), vbBlack
    frmMain.Picture1.Line (WallLinesX(3), WallLinesY(3))-(WallLinesX(3), WallLinesY(3) - 100), vbBlack
    frmMain.Picture1.Line (WallLinesX(4), WallLinesY(4))-(WallLinesX(2), WallLinesY(2) - 100), vbBlack
    frmMain.Picture1.Line (WallLinesX(5) - 1, WallLinesY(5) + 1)-(WallLinesX(3), WallLinesY(3) - 99), vbBlack
End Sub
Sub CheckKey(KeyStrokes As String)
Dim Keys(1 To 2) As Integer '1 for one key code like left,up,down,right
                            '2 for diagnole codes
Dim Diagnole As Boolean
If KeyRunning = False Then
'// get keys
    '// first we check if theres more then 2 key inputs if theres more we only take
    '// the first one
    KeyRunning = True
    Debug.Print KeyStrokes
    Debug.Print UBound(Split(KeyStrokes, ";"))
    If UBound(Split(KeyStrokes, ";")) = 2 Then
        Keys(1) = Val(Split(KeyStrokes, ";")(0))
        Keys(2) = Val(Split(KeyStrokes, ";")(1))
        Diagnole = True
    ElseIf UBound(Split(KeyStrokes, ";")) > 0 Then
        Keys(1) = Val(Split(KeyStrokes, ";")(0))
        Diagnole = False
    End If
'//checkkey input
    If AniRunning = False Then 'if anirunning let it finish before running more commands
        If Keys(1) = 39 And Keys(2) = 38 Or Keys(1) = 38 And Keys(2) = 39 Then '// right up
            Call modEngine.MoveBodyDiagnole(False, False, True, False, Keys(2))
        ElseIf Keys(2) = 38 And Keys(1) = 37 Or Keys(1) = 38 And Keys(2) = 37 Then '// up left
            Call modEngine.MoveBodyDiagnole(True, False, False, False, Keys(2))
        ElseIf Keys(1) = 40 And Keys(2) = 37 Or Keys(1) = 37 And Keys(2) = 40 Then '// down left
            Call modEngine.MoveBodyDiagnole(False, True, False, False, Keys(2))
        ElseIf Keys(1) = 40 And Keys(2) = 39 Or Keys(1) = 39 And Keys(2) = 40 Then '// right down
            Call modEngine.MoveBodyDiagnole(False, False, False, True, Keys(2))
        Else
            Diagnole = False
        End If
        If Diagnole = False Then
            If Keys(1) = 39 Then 'right
                Call modEngine.MoveBody(False, True, False, False, True)
            ElseIf Keys(1) = 37 Then 'left
                Call modEngine.MoveBody(True, False, False, False, True)
            ElseIf Keys(1) = 38 Then 'up
                Call modEngine.MoveBody(False, False, True, False, True)
            ElseIf Keys(1) = 40 Then 'down
                Call modEngine.MoveBody(False, False, False, True, True)
            ElseIf Keys(1) = 13 Then 'Enter
                Call modEngine.InteractCommand
            End If
        End If
    End If
    KeyRunning = False
End If
KeyStrokes = ""
End Sub
Sub ColorWalls()
'// color the walls
Dim ColoredLines(1 To 2) As Integer
Dim Lines As Integer
Dim Count As Integer
Dim LineHeight As Integer
Dim LastLine(1 To 4) As Integer
    '// hrmm i think i could of used the paint command
    '// instead of all this coding but w/e..
    LineHeight = WallLinesY(1) - (WallLinesY(1) - 100)
    '// this colors and creates the first wall
    ColoredLines(1) = WallLinesX(2) - WallLinesX(1) - 2
    ColoredLines(2) = WallLinesX(3) - WallLinesX(2) - 1
    Do While Count < ColoredLines(1)
        Lines = Lines + 1
        For walllines = 1 To 2
            Count = Count + 1
            frmMain.Picture1.Line (WallLinesX(1) + Count, WallLinesY(1) - Lines - 1)-(WallLinesX(1) + Count, WallLinesY(1) - Lines - LineHeight), WallColor
        Next walllines
        LastLine(1) = WallLinesX(1) + Count
        LastLine(2) = WallLinesY(1) - Lines - 1
        LastLine(3) = WallLinesX(1) + Count
        LastLine(4) = WallLinesY(1) - Lines - LineHeight
    Loop
    frmMain.Picture1.Line (LastLine(1) + 1, LastLine(2) - 1)-(LastLine(3) + 1, LastLine(4) - 1), WallColor
    frmMain.Picture1.Line (WallLinesX(1), WallLinesY(1) - 1)-(WallLinesX(1), WallLinesY(1) - 100), WallColor
    '// this colors and creates the second wall
    Count = 0
    Lines = 0
    Do While Count < ColoredLines(2) - 1
        Lines = Lines + 1
        For walllines = 1 To 2
            Count = Count + 1
            frmMain.Picture1.Line (WallLinesX(2) + Count, WallLinesY(2) + Lines - 1)-(WallLinesX(2) + Count, WallLinesY(2) + Lines - LineHeight + 1), WallColor
        Next walllines
    Loop
    frmMain.Picture1.Line (WallLinesX(3) - 1, WallLinesY(3) - 1)-(WallLinesX(3) - 1, WallLinesY(3) - 99), WallColor
    frmMain.Picture1.Line (WallLinesX(2) + 1, WallLinesY(2) - 1)-(WallLinesX(2) + 1, WallLinesY(2) - LineHeight + 1), WallColor
End Sub
Function CanMove(ByVal X As Integer, ByVal Y As Integer) As Boolean
'// this is a collision check if theres a furniture in the way or a dead space
    If FurniData(Y, X) = "0" And MapData(Y, X) = "O" Then
        CanMove = True
    Else
        CanMove = False
    End If
End Function
Sub BehindFurni(ByVal UserX As Integer, ByVal UserY As Integer)
On Error Resume Next
'//this sub is used to create the illusion that the character is "behind" an item
    frmMain.formcaption.Caption = "Y=" & UserY & " X=" & UserX
    '//checks if the user is at the end of the maps width
    '// if not then check what tile where theres a furni in front back or at the side
    '// so i can then get the points color it in and make the illusion
    If FurniData(UserY, UserX + 1) <> "0" And UserX + 1 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 0, 1, True, False, False)
    End If
    If FurniData(UserY + 1, UserX + 1) <> "0" And UserX + 1 <= LengthX And UserY + 1 <= UBound(MapData()) Then
        Call modEngine.PasteFurniTop(UserX, UserY, 1, 1, False, True, False)
    End If
    If FurniData(UserY + 1, UserX) <> "0" And UserY + 1 <= UBound(MapData()) Then
        Call modEngine.PasteFurniTop(UserX, UserY, 1, 0, False, False, True)
    End If
    If FurniData(UserY + 1, UserX + 2) <> "0" And UserY + 1 <= UBound(MapData()) And UserX + 2 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 1, 2, True, False, False)
    End If
    If FurniData(UserY + 2, UserX + 2) <> "0" And UserY + 2 <= UBound(MapData()) And UserX + 2 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 2, 2, False, True, False)
    End If
    If FurniData(UserY + 2, UserX + 1) <> "0" And UserY + 2 <= UBound(MapData()) And UserX + 1 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 2, 1, False, False, True)
    End If
    If FurniData(UserY + 2, UserX + 3) <> "0" And UserY + 2 <= UBound(MapData()) And UserX + 3 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 2, 3, True, False, False)
    End If
    If FurniData(UserY + 3, UserX + 3) <> "0" And UserY + 3 <= UBound(MapData()) And UserX + 3 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 3, 3, False, True, False)
    End If
    If FurniData(UserY + 3, UserX + 2) <> "0" And UserY + 3 <= UBound(MapData()) And UserX + 2 <= LengthX Then
        Call modEngine.PasteFurniTop(UserX, UserY, 3, 2, False, False, True)
    End If
    frmMain.Picture1.Refresh
End Sub
Sub GetFurniInfo(FurniChar As String, PieceWidth As Integer, PieceHeight As Integer, PicHdc As Long, MaskHdc As Long)
'// this will input the furnis width and height
    If FurniChar = "I" Then
        PieceWidth = Memory.picItemBox.Width
        PieceHeight = Memory.picItemBox.Height
        PicHdc = Memory.picItemBox.hdc
        MaskHdc = Memory.picItemBoxMask.hdc
    ElseIf FurniChar = "D" Then
        PieceWidth = Memory.picMountinDew.Width
        PieceHeight = Memory.picMountinDew.Height
        PicHdc = Memory.picMountinDew.hdc
        MaskHdc = Memory.picMountinDewMask.hdc
    End If
End Sub
Function FindTile(ByVal Y As Integer, ByVal X As Integer) As Integer
'// function thats finds then number of the tile by the y and x axis
    FindTile = (Y - 1) * LengthX + X
End Function
Sub UpdateChar()
'// updates the char height and width
    If Direction = Up Then
        If Drink = False Then
            CharWidth = Memory.picBodyUp.Width
            CharHeight = Memory.picBodyUp.Height
        Else
            CharWidth = Memory.picBodyUpDrink.Width
            CharHeight = Memory.picBodyUpDrink.Height
        End If
    ElseIf Direction = down Then
        If Drink = False Then
            CharWidth = Memory.picBodyDown.Width
            CharHeight = Memory.picBodyDown.Height
        Else
            CharWidth = Memory.picBodyDrinkDown.Width
            CharHeight = Memory.picBodyDrinkDown.Height
        End If
    ElseIf Direction = Left Then
        If Drink = False Then
            CharWidth = Memory.picBodyLeft.Width
            CharHeight = Memory.picBodyLeft.Height
        Else
            CharWidth = Memory.picBodyLeftDrink.Width
            CharHeight = Memory.picBodyLeftDrink.Height
        End If
    Else '// Right
        If Drink = False Then
            CharWidth = Memory.picBodyRight.Width
            CharHeight = Memory.picBodyRight.Height
        Else
            CharWidth = Memory.picBodyDrinkRight.Width
            CharHeight = Memory.picBodyDrinkRight.Height
        End If
    End If
End Sub
Sub CopyFurniPiece(ByVal Y As Long, ByVal X As Long, Width As Integer, Height As Integer)
    '//  copies the furni piece from the real background to the original
    Call BitBlt(frmMain.Picture1.hdc, X, Y, Width, Height, frmMain.PicRealBck.hdc, X, Y, vbSrcCopy)
End Sub
Sub PasteFurniTop(ByVal UserX, ByVal UserY, ByVal AddY, ByVal AddX, Left As Boolean, Middle As Boolean, Right As Boolean)
 '// This sub will recreate the furni image paste on top of the picture making the character seem its behind
Dim PictureHDC As Long
Dim PSetX As Long
Dim PSetY As Long
Dim MaskHdc As Long
Dim BckColor As Long
Dim PicColor As Long
Dim MaskColor As Long
Dim Tiles As Integer
Dim PieceWidth As Integer
Dim PieceHeight As Integer
Dim Difference As Long
    If Left = True Then
        '// get furni info
        Call GetFurniInfo(FurniData(UserY + AddY, UserX + AddX), PieceWidth, PieceHeight, PictureHDC, MaskHdc)
        '// loop through the part of the image that needs coloring
        For Y = 1 To Abs(CharY + CharHeight - FurniY(FindTile(UserY + AddY, UserX + AddX)))
            For X = 1 To CharX + CharWidth - FurniX(FindTile(UserY + AddY, UserX + AddX))
                '//get colors of the pic and mask
                PicColor = GetPixel(PictureHDC, X - 1, Y)
                MaskColor = GetPixel(MaskHdc, X - 1, Y)
                '//get the points on the users map where to color in
                PSetX = FurniX(FindTile(UserY + AddY, UserX + AddX)) + X - 1
                PSetY = FurniY(FindTile(UserY + AddY, UserX + AddX)) + Y
                '// if piccolor = 0 which means its black and mask color =16777215 which is white
                '// it means that that point needs to be transparent so we dont
                '// color anything in
                If PicColor = 0 And MaskColor = 16777215 Then
                Else
                    BckColor = GetPixel(frmMain.PicRealBck.hdc, PSetX, PSetY)
                    Call SetPixel(frmMain.Picture1.hdc, PSetX, PSetY, BckColor)
                End If
                DoEvents
            Next X
        Next Y
    ElseIf Middle = True Then
        '// get furni info
        Call GetFurniInfo(FurniData(UserY + AddY, UserX + AddX), PieceWidth, PieceHeight, PictureHDC, MaskHdc)
        '// loop through the part of the image that needs coloring
        For Y = 1 To Abs(CharY + CharHeight - FurniY(FindTile(UserY + AddY, UserX + AddX)))
            For X = 1 To CharWidth
                '// we find the difference between the point of the furniture
                '// and the point where the character stands
                '// we find this difference because we need a starting point
                '// where we would loop the masks that everything will correspond
                Difference = CharX - FurniX(FindTile(UserY + AddY, UserX + AddX))
                '//get colors of the pic and mask
                PicColor = GetPixel(PictureHDC, X + Difference - 1, Y - 1)
                MaskColor = GetPixel(MaskHdc, X + Difference - 1, Y - 1)
                '//get the points on the users map where to color in
                PSetX = CharX + X - 1
                PSetY = FurniY(FindTile(UserY + AddY, UserX + AddX)) + Y - 1
                '// if piccolor = 0 which means its black and mask color =16777215 which is white
                '// it means that that point needs to be transparent so we dont
                '// color anything in
                If PicColor = 0 And MaskColor = 16777215 Then
                Else
                    BckColor = GetPixel(frmMain.PicRealBck.hdc, PSetX, PSetY)
                    Call SetPixel(frmMain.Picture1.hdc, PSetX, PSetY, BckColor)
                End If
                DoEvents
            Next X
        Next Y
    ElseIf Right = True Then
        '// get furni info
        Call GetFurniInfo(FurniData(UserY + AddY, UserX + AddX), PieceWidth, PieceHeight, PictureHDC, MaskHdc)
        '// loop through the part of the image that needs coloring
        For Y = 1 To Abs((CharY + CharHeight) - FurniY(FindTile(UserY + AddY, UserX + AddX)))
            For X = 1 To (FurniX(FindTile(UserY + AddY, UserX + AddX)) + PieceWidth) - CharX
                '//get colors of the pic and mask
                PicColor = GetPixel(PictureHDC, PieceWidth - X - 4, Y)
                MaskColor = GetPixel(MaskHdc, PieceWidth - X - 4, Y)
                '//get the points on the users map where to color in
                PSetX = FurniX(FindTile(UserY + AddY, UserX + AddX)) + PieceWidth - X - 4
                PSetY = FurniY(FindTile(UserY + AddY, UserX + AddX)) + Y
                '// if piccolor = 0 which means its black and mask color =16777215 which is white
                '// it means that that point needs to be transparent so we dont
                '// color anything in
                If PicColor = 0 And MaskColor = 16777215 Then
                Else
                    BckColor = GetPixel(frmMain.PicRealBck.hdc, PSetX, PSetY)
                    Call SetPixel(frmMain.Picture1.hdc, PSetX, PSetY, BckColor)
                End If
                DoEvents
            Next X
        Next Y
    End If
End Sub
Sub ColorBehindItem(ByVal Y As Integer, ByVal X As Integer, Height As Integer, Width As Integer, CharCode As String)
    If CharCode = "I" Then
        Call BitBlt(frmMain.Picture1.hdc, X, Y, Width, Height, Memory.picItemBoxMask.hdc, 0, 0, vbSrcAnd)
        Call BitBlt(frmMain.Picture1.hdc, X, Y, Width, Height, Memory.picItemBox.hdc, 0, 0, vbSrcPaint)
    ElseIf CharCode = "D" Then
        Call BitBlt(frmMain.Picture1.hdc, X, Y, Width, Height, Memory.picMountinDewMask.hdc, 0, 0, vbSrcAnd)
        Call BitBlt(frmMain.Picture1.hdc, X, Y, Width, Height, Memory.picMountinDew.hdc, 0, 0, vbSrcPaint)
    End If
End Sub
Sub MoveBodyDiagnole(LeftUp As Boolean, LeftDown As Boolean, RightUp As Boolean, RightDown As Boolean, LastKey As Integer)
On Error Resume Next
Dim Tile As Integer
Dim MapX As Integer
Dim MapY As Integer
    '// find the user's chars points
    For Y = 1 To UBound(MapData())
        For X = 1 To LengthX
            If CharData(Y, X) = "O" Then
                '// update the coordinates
                If LeftUp = True Then
                    MapX = X - 1
                    MapY = Y - 1
                    If MapX > 0 And MapY > 0 Then
                        If CanMove(MapX, MapY) = True Then
                            CharData(Y, X) = "X"
                            CharData(MapY, MapX) = "O"
                            Tile = FindTile(MapY, MapX)
                            '// cut the image and restore the background
                            BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                            CharY = TilesY(Tile) + Memory.picTile1.Height / 2 - CharHeight + 7
                        Else
                            Exit Sub
                        End If
                        Y = UBound(MapData())
                        X = LengthX
                    Else
                        Exit Sub
                    End If
                ElseIf LeftDown = True Then
                    MapX = X - 1
                    MapY = Y + 1
                    If MapX > 0 And MapY < UBound(MapData()) Then
                        If CanMove(MapX, MapY) = True Then
                            CharData(Y, X) = "X"
                            CharData(MapY, MapX) = "O"
                            Tile = FindTile(MapY, MapX)
                            '// cut the image and restore the background
                            BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                            CharX = TilesX(Tile) + 7
                        Else
                            Exit Sub
                        End If
                        Y = UBound(MapData())
                        X = LengthX
                    Else
                        Exit Sub
                    End If
                ElseIf RightDown = True Then
                    MapX = X + 1
                    MapY = Y + 1
                    If MapX < LengthX And MapY < UBound(MapData()) Then
                        If CanMove(MapX, MapY) = True Then
                            CharData(Y, X) = "X"
                            CharData(MapY, MapX) = "O"
                            Tile = FindTile(MapY, MapX)
                            '// cut the image and restore the background
                            BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                            CharY = TilesY(Tile) + 7 + Memory.picTile1.Height / 2 - CharHeight
                        Else
                            Exit Sub
                        End If
                        Y = UBound(MapData())
                        X = LengthX
                    Else
                        Exit Sub
                    End If
                ElseIf RightUp = True Then
                    MapX = X + 1
                    MapY = Y - 1
                    If MapX < LengthX And MapY > 0 Then
                        If CanMove(MapX, MapY) = True Then
                            CharData(Y, X) = "X"
                            CharData(MapY, MapX) = "O"
                            Tile = FindTile(MapY, MapX)
                            '// cut the image and restore the background
                            BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
                            CharX = TilesX(Tile) + 7
                        Else
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                    Y = UBound(MapData())
                    X = LengthX
                End If
            End If
        Next X
    Next Y
    If LastKey = 40 Then
        '// paste the image with the down direction
        Direction = down
        Call UpdateChar
        If Drink = False Then
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDownMask.Width, Memory.picBodyDownMask.Height, Memory.picBodyDownMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDown.Width, Memory.picBodyDown.Height, Memory.picBodyDown.hdc, 0, 0, vbSrcPaint
        Else
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDownMask.Width, Memory.picBodyDrinkDownMask.Height, Memory.picBodyDrinkDownMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDown.Width, Memory.picBodyDrinkDown.Height, Memory.picBodyDrinkDown.hdc, 0, 0, vbSrcPaint
            '// paste the image with drink
        End If
    ElseIf LastKey = 39 Then
        '// paste the image with the right direction
        Direction = Right
        Call UpdateChar
        If Drink = False Then
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRightMask.Width, Memory.picBodyRightMask.Height, Memory.picBodyRightMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyRight.Width, Memory.picBodyRight.Height, Memory.picBodyRight.hdc, 0, 0, vbSrcPaint
        Else
            '// paste the image with the drink
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRightMask.Width, Memory.picBodyDrinkRightMask.Height, Memory.picBodyDrinkRightMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRight.Width, Memory.picBodyDrinkRight.Height, Memory.picBodyDrinkRight.hdc, 0, 0, vbSrcPaint
        End If
    ElseIf LastKey = 37 Then
        '// paste the image with the left direction
        Direction = Left
        Call UpdateChar
        If Drink = False Then
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftMask.Width, Memory.picBodyLeftMask.Height, Memory.picBodyLeftMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeft.Width, Memory.picBodyLeft.Height, Memory.picBodyLeft.hdc, 0, 0, vbSrcPaint
        Else
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrinkMask.Width, Memory.picBodyLeftDrinkMask.Height, Memory.picBodyLeftDrinkMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrink.Width, Memory.picBodyLeftDrink.Height, Memory.picBodyLeftDrink.hdc, 0, 0, vbSrcPaint
            '// paste the image with the drink
        End If
    ElseIf LastKey = 38 Then
        '// paste the image with the up direction
        Direction = Up
        Call UpdateChar
        If Drink = False Then
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpMask.Width, Memory.picBodyUpMask.Height, Memory.picBodyUpMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUp.Width, Memory.picBodyUp.Height, Memory.picBodyUp.hdc, 0, 0, vbSrcPaint
        Else
            '// paste the image with the drink
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrinkMask.Width, Memory.picBodyUpDrinkMask.Height, Memory.picBodyUpDrinkMask.hdc, 0, 0, vbSrcAnd
            BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrink.Width, Memory.picBodyUpDrink.Height, Memory.picBodyUpDrink.hdc, 0, 0, vbSrcPaint
        End If
    End If
    Call BehindFurni(MapX, MapY)
    frmMain.Picture1.Refresh
End Sub
Sub InteractCommand()
Dim CharMapX As Integer
Dim CharMapY As Integer
    '// Algorithm
    '====================
    '//Will check what direction the user is facing
    '// then we find the object on his left/right/up/down side
    '// see if it's interactble then after that run animation corresponding
    '// to that object and then its done
    
    '//Find where the user is standing
    For Y = 1 To UBound(MapData())
        For X = 1 To LengthX
            If CharData(Y, X) = "O" Then
                '// record the coordination
                CharMapX = X
                CharMapY = Y
                '// end the loop
                Y = UBound(MapData())
                X = LengthX
            End If
        Next X
    Next Y
    If Direction = Left Then
        '// check if the furni is interactble
        If CharMapX - 1 > 0 Then '// check for boundary offsets
            If CheckInteract(FurniData(CharMapY, CharMapX - 1)) = True Then
                Call InteractFurni(FurniData(CharMapY, CharMapX - 1))
            End If
        End If
    ElseIf Direction = Right Then
        '// check if the furni is interactble
        If CharMapX + 1 < LengthX Then '// check for boundary offsets
            If CheckInteract(FurniData(CharMapY, CharMapX + 1)) = True Then
                Call InteractFurni(FurniData(CharMapY, CharMapX + 1))
            End If
        End If
    ElseIf Direction = Up Then
        '// check if the furni is interactble
        If CharMapY - 1 > 0 Then '// check for boundary offsets
            If CheckInteract(FurniData(CharMapY - 1, CharMapX)) = True Then
                Call InteractFurni(FurniData(CharMapY - 1, CharMapX))
            End If
        End If
    Else '//else if direction = down
        '// check if the furni is interactble
        If CharMapY + 1 < UBound(MapData()) Then '// check for boundary offsets
            If CheckInteract(FurniData(CharMapY + 1, CharMapX)) = True Then
                Call InteractFurni(FurniData(CharMapY + 1, CharMapX))
            End If
        End If
    End If
    Call UpdateChar
    Call BehindFurni(CharMapX, CharMapY)
    frmMain.Picture1.Refresh
End Sub
Function CheckInteract(ObjData As String) As Boolean
    '// check which object
    If ObjData = "D" Then
        CheckInteract = True
    End If
End Function
Sub InteractFurni(ObjData As String)
    '// Check which object animation to run
    If ObjData = "D" Then
        If Drink = False Then
            Drink = True
            '// animation loop for the fridge door to open
            '// then cut out the image of the user and replace it with
            '// with an image of holding the drink but first check what direction
            BitBlt frmMain.Picture1.hdc, CharX, CharY, CharWidth, CharHeight, frmMain.PicRealBck.hdc, CharX, CharY, vbSrcCopy
            If Direction = Right Then
                '// paste the image cooresponding to the direction
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRightMask.Width, Memory.picBodyDrinkRightMask.Height, Memory.picBodyDrinkRightMask.hdc, 0, 0, vbSrcAnd
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkRight.Width, Memory.picBodyDrinkRight.Height, Memory.picBodyDrinkRight.hdc, 0, 0, vbSrcPaint
            ElseIf Direction = Left Then
                '// paste the image cooresponding to the direction
                CharX = CharX - 8
                CharY = CharY
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrinkMask.Width, Memory.picBodyLeftDrinkMask.Height, Memory.picBodyLeftDrinkMask.hdc, 0, 0, vbSrcAnd
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyLeftDrink.Width, Memory.picBodyLeftDrink.Height, Memory.picBodyLeftDrink.hdc, 0, 0, vbSrcPaint
            ElseIf Direction = down Then
                 '// paste the image cooresponding to the direction
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDownMask.Width, Memory.picBodyDrinkDownMask.Height, Memory.picBodyDrinkDownMask.hdc, 0, 0, vbSrcAnd
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyDrinkDown.Width, Memory.picBodyDrinkDown.Height, Memory.picBodyDrinkDown.hdc, 0, 0, vbSrcPaint
            ElseIf Direction = Up Then
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrinkMask.Width, Memory.picBodyUpDrinkMask.Height, Memory.picBodyUpDrinkMask.hdc, 0, 0, vbSrcAnd
                BitBlt frmMain.Picture1.hdc, CharX, CharY, Memory.picBodyUpDrink.Width, Memory.picBodyUpDrink.Height, Memory.picBodyUpDrink.hdc, 0, 0, vbSrcPaint
            End If
        End If
    End If
End Sub
