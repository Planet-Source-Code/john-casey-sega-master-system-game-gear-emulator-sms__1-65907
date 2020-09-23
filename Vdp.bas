Attribute VB_Name = "Vdp"
Option Explicit

'//     (1) Include GG Palette.
'//     (2) Are all &HFF& Gates needed?
'//     (3) Re-enable sprite collision?

Private VRAM()                  As Long             'Video RAM.
Private CRAM()                  As Long             'Colour RAM.

Private vdpreg()                As Long             'VDP Registers.
Private status                  As Long             'Status Register.

Private first_byte              As Long             'First or Second Byte of Command Word.
Private command_byte            As Long             'Command Word First Byte Latch.
Private location                As Long             'Location in VRAM.
Private operation               As Long             'Store type of operation taking place.
Private read_buffer             As Long             'Buffer VRAM Reads.

Private line                    As Long             'Current Line Number.
Private counter                 As Long             'Vertical Line Interrupt Counter.
Private lineint                 As Long             'Line interrupt Pending.
Private frameint                As Long             'Frame interrupt Pending.

Private bg_priority()           As Long             'Background Priorites.
'Private spritecol()            As Long             'Sprite Collisions.
Public display()                As Long             'Pointer to Current Display.
Private SMS_PALETTE(63)         As Long             'SMS Color palette

Private Const H_START           As Long = 0&        'Horizontal viewport start
Private Const H_END             As Long = SMS_WIDTH 'Horizontal viewport end

'--------------~
'   Reset VDP.
'--------------~
Public Sub reset()

    ReDim VRAM(16385&) As Long              'Clear VRAM memory.
    ReDim CRAM(64&) As Long                 'Clear CRAM memory.
    ReDim vdpreg(16&) As Long               'Clear VDPREG.
    ReDim bg_priority(270) As Long          'Clear bg priorities.
    'ReDim spritecol(SMS_WIDTH) As Long     'Clear sprite collisions
    ReDim display(49408) As Long


    first_byte = True
       lineint = False
      frameint = False
      location = 0&
       counter = 0&
        status = 0&
    
     vdpreg(2) = 14&                            'SMS Bios Value.
     vdpreg(5) = 126&                           'SMS Bios Value.
     vdpreg(6) = 255&                           'SMS Bios Value.
    vdpreg(10) = 1&                             'SMS Bios Value.

    irqsetLine = False

End Sub

'---------------------~
'   Read Vertical Port.
'---------------------~
Public Function getVCount() As Long

    'If (ntsc) Then
        If (line > &HDA&) Then getVCount = line - 6 Else getVCount = line
    'Else
        'If (line > &HF2&) Then getVCount = line - &H39& Else getVCount = line
    'End If
    
    'MAKE SURE TO RETURN LINE!!!
End Function

'---------------------------------~
'   Write to VDP Control Port (BF).
'---------------------------------~
Public Sub controlWrite(value As Long)

    '// Store First Byte of Command Word
    If (first_byte) Then
    
        first_byte = False
        command_byte = value
        Exit Sub
        
    End If
    
        first_byte = True
        
        '// Set VDP Register
        If (ShiftRight(value, 4&) = 8&) Then
            
            vdpreg(value And 15&) = command_byte

            If (lineint) Then
                If (((value And 15) = 0) And ((command_byte And 16) = 0)) Then
                    irqsetLine = False
                ElseIf (((value And 15) = 0) And ((command_byte And 16) = 16)) Then
                    irqsetLine = True
                End If
            End If

            If (frameint) Then
                If (((value And 15) = 1) And ((command_byte And 32) = 0)) Then
                    irqsetLine = False
                ElseIf (((value And 15) = 1) And ((command_byte And 32) = 32)) Then
                    irqsetLine = True
                End If
            End If
                
        Else

            '// Operation from B6 + B7
            operation = ShiftRight(value, 6&)
            '// Set location in VRAM
            location = command_byte + ShiftLeft(value And 63&, 8&)
            
            If (operation = 0&) Then
                
                read_buffer = VRAM(location)
                If (location > 16383&) Then location = 0& Else location = location + 1&
            
            End If
        End If
     
End Sub

'------------------------------~
'   Read VDP Control Port (BF).
'------------------------------~
Public Function controlRead() As Long

    first_byte = True
    
    Dim statuscopy As Long
    statuscopy = status
    
    status = status And Not &H80& And Not &H40& And Not &H20&
    
    '// Clear pending interrupts
    lineint = False
    frameint = False
    
    '// Clear IRQ Line
    irqsetLine = False
    
    controlRead = statuscopy

End Function

'------------------------------~
'   Write to VDP Data Port (BE).
'------------------------------~
Public Sub dataWrite(value As Long)

    first_byte = True
    
    Select Case (operation)
        
        '           // VRAM Write
        Case 0& To 2&:  VRAM(location) = value
        
        '           // CRAM Write
        Case 3&
                    CRAM(location And &H1F&) = SMS_PALETTE(value And &H3F&)
                    
        'Case Else:  Debug.Print "Unknown dataWrite " & Hex(operation)
    End Select

    If (location < 16383&) Then location = location + 1& Else location = 0&
    
End Sub

'---------------------------~
'   Read VDP Data Port (BE).
'---------------------------~
Public Function dataRead() As Long

    first_byte = True
    
    'VRAM Read
    If (operation < 2&) Then
        dataRead = read_buffer
        read_buffer = VRAM(location)
    'Else
        'Debug.Print "Unsupported dataRead " & Hex(operation)
    End If
    
    If (location < 16383&) Then location = location + 1& Else location = 0&
                        
End Function

'--------------------------------~
'   Render Line of SMS/GG Display.
'--------------------------------~
Public Sub drawLine(lineno As Long)
    
    'ReDim spritecol(SMS_WIDTH) As Long
    
    If (Not (vdpreg(1) And 64) = 0) Then
        drawBg (lineno)
        drawSprite (lineno)
    End If
    
    If ((vdpreg(0) And &H20&) = &H20&) Then blankColumn (lineno)

End Sub

'----------------------------------~
'   Render Line of Background Layer.
'----------------------------------~
Private Sub drawBg(lineno As Long)

    Dim X               As Long
    Dim Y               As Long:    Y = lineno
    Dim bgt             As Long:    bgt = ShiftLeft(vdpreg(2&) And 15& And Not 1&, 10&)
    Dim pattern         As Long
    Dim pal             As Long
    Dim hscroll         As Long
    Dim start_column    As Long
    Dim start_row       As Long
    Dim tilex           As Long
    Dim tiley           As Long
    Dim tile_props      As Long
    Dim firstbyte       As Long
    Dim secondbyte      As Long
    Dim priority        As Long
    Dim address         As Long
    Dim xpos            As Long
    Dim row_precal      As Long
    Dim address0        As Long
    Dim address1        As Long
    Dim address2        As Long
    Dim address3        As Long
    Dim bit             As Long
    Dim a               As Long
    Dim colour          As Long


    '// Vertical Scroll Fine Tune
    If (Not vdpreg(9&) = 0&) Then lineno = lineno + (vdpreg(9&) And 7&)


    '// Vertical Scroll; Row of Tile to Plot.
    start_row = lineno: tiley = (lineno And 7&)


    '// Top Two Rows Not Affected by Horizontal Scrolling (SMS Only)
    If (((vdpreg(0&) And 64&) = 64&) And (Y < 16&)) Then
        hscroll = 0&
    ElseIf (Not vdpreg(8&) = 0&) Then
        hscroll = vdpreg(8&) And 7&
        start_column = 32& - ShiftRight(vdpreg(8&), 3&)
    End If


    'If (setup.is_gg) Then... [Add coding for GameGear Here]


    '// Adjust Vertical Scroll
    If (Not vdpreg(9&) = 0&) Then
        start_row = start_row + (vdpreg(9&) And 248&)
        If (start_row > 223&) Then start_row = start_row - 224&
    End If


    '// Cycle through background table
    For X = H_START To H_END Step 8&
    
    
        '// wraps at 32
        If (start_column = 32&) Then start_column = 0&
        
        
        '// Rightmost 8 columns Not Affected by Vertical Scrolling
        If (((vdpreg(0&) And 128&) = 128&) And (X > 184&)) Then
            start_row = Y
            tiley = (Y And 7&)
        End If
        

        '// Get the two bytes from VRAM containing the tile's properties
        tile_props = bgt + ShiftLeft(start_column, 1&) + ShiftLeft(start_row - tiley, 3&)
        firstbyte = VRAM(tile_props) ' And 255&
        secondbyte = VRAM(tile_props + 1) ' And 255&
        
        
        '// Priority of tile
        priority = secondbyte And 16&
        
        
        '// Select Palette (Extended Colors).
        If ((secondbyte And 8&) = 8&) Then pal = 16& Else pal = 0&
        
        
        '// Pattern Number
        pattern = ShiftLeft(firstbyte + ShiftLeft(secondbyte And 1&, 8&), 5&)
        

        '// Vertical Tile Flip
        If ((secondbyte And 4&) = 0&) Then address = ShiftLeft(tiley, 2&) + pattern Else address = ShiftLeft(7& - tiley, 2&) + pattern
        
        
        '// Rowcount
        tilex = 0&
        
        
        '// Precalculate Y Position
        row_precal = ShiftLeft(Y, 8&)


        address0 = VRAM(address)
        address1 = VRAM(address + 1&)
        address2 = VRAM(address + 2&)
        address3 = VRAM(address + 3&)


        '//Plots row of 8 pixels
        bit = 128&
        
        Do

            '// Horizontal Tile Flip
            If ((secondbyte And 2&) = 0&) Then xpos = tilex + hscroll + X Else xpos = 7& - tilex + hscroll + X
                colour = 0 '// Set Colour of Pixel (0-15)
                If (Not (address0 And bit) = 0&) Then colour = colour Or 1&
                If (Not (address1 And bit) = 0&) Then colour = colour Or 2&
                If (Not (address2 And bit) = 0&) Then colour = colour Or 4&
                If (Not (address3 And bit) = 0&) Then colour = colour Or 8&
                
            
                '// Set Priority Array (Sprites over background tile)
                If (priority = 16&) And (Not colour = 0&) Then bg_priority(xpos) = True Else bg_priority(xpos) = False

                
                display(xpos + row_precal) = CRAM(colour + pal)
                tilex = tilex + 1&
                bit = ShiftRight(bit, 1&)
        Loop Until bit = 0&

        start_column = start_column + 1&
    Next X

End Sub

'-------------------------------~
'   Render Line of Sprite Layer.
'-------------------------------~
Private Sub drawSprite(lineno As Long)


    Dim sat         As Long:    sat = ShiftLeft(vdpreg(5&) And Not 1& And Not &H80&, 7&)
    Dim count       As Long
    Dim height      As Long:    height = 8&
    Dim zoomed      As Boolean
    Dim spriteno    As Long
    Dim Y           As Long
    Dim address     As Long
    Dim X           As Long
    Dim i           As Long
    Dim row_precal  As Long
    Dim adr         As Long
    Dim bit         As Long
    Dim pixel       As Long
    
    '// Enabled 8x16 Sprites
    If ((vdpreg(1) And 2&) = 2&) Then height = 16&


    '// Enable Zoomed Sprites
    If ((vdpreg(1&) And 1&) = 1&) Then
        height = ShiftLeft(height, 1&)
        zoomed = True
    End If
    
    '// Search Sprite Attribute Table (64 Bytes)
    For spriteno = 0& To 63&
    
        'If (count >= 8) Then
        '    status = status Or &H40&
        '    Exit Sub
        'end if
    
        Y = VRAM(sat + spriteno)
        address = sat + ShiftLeft(spriteno, 1&)
        X = VRAM(address + &H80&)
        i = VRAM(address + &H81&)
        
        If (Y = 208&) Then Exit Sub
        
        Y = Y + 1&
        
        If ((vdpreg(6&) And 4&) = 4&) Then i = i Or &H100&
        If ((vdpreg(1&) And 2&) = 2&) Then i = i And Not 1&
        If ((vdpreg(0&) And 8&) = 8&) Then X = X - 8&
        If (Y > 240&) Then Y = Y - 256&
        
        If ((lineno >= Y) And ((lineno - Y) < height)) Then
        
            row_precal = ShiftLeft(lineno, 8&)
            
            If zoomed = False Then  '// Normal sprites (Width = 8)
            
                adr = ShiftLeft(i, 5&) + ShiftLeft(lineno - Y, 2&)
                bit = &H80&
                
                Do
                
                    If ((X >= 0&) And (Y >= 0&) And (X <= 255&)) Then
                        plotSpritePixel X, lineno, X + row_precal, adr, bit
                    End If
                    
                    X = X + 1&
                    bit = ShiftRight(bit, 1&)
                    
                Loop Until bit = 0&
                
            Else                    '// Zoomed sprites (Width = 16)
            
                adr = ShiftLeft(i, 5&) + ShiftLeft(ShiftRight(lineno - Y, 1&), 2&)
                bit = &H80&
                
                Do

                    If ((X + pixel + 1& >= 0) And (Y >= 0&) And (X <= 255&)) Then
                        '// Plot Two Pixels
                        plotSpritePixel X + pixel, lineno, X + pixel + row_precal, adr, bit
                        plotSpritePixel X + pixel + 1&, lineno, X + pixel + row_precal, adr, bit
                    End If
                    
                    X = X + 1&
                    bit = ShiftRight(bit, 1&)
                
                Loop Until bit = 0&
            End If
            count = count + 1&
        End If
    Next spriteno

End Sub

'------------------------------~
'   Plot a single sprite pixel.
'------------------------------~
Private Sub plotSpritePixel(X As Long, Y As Long, location As Long, address As Long, bit As Long)

    Dim colour As Long: colour = 0&
    
    If (Not (VRAM(address + 0&) And bit) = 0&) Then colour = colour Or 1&
    If (Not (VRAM(address + 1&) And bit) = 0&) Then colour = colour Or 2&
    If (Not (VRAM(address + 2&) And bit) = 0&) Then colour = colour Or 4&
    If (Not (VRAM(address + 3&) And bit) = 0&) Then colour = colour Or 8&
    
    
    If (Not (bg_priority(X)) And (Not (colour) = 0&)) Then
        
        'If (Not spritecol(X)) Then
            
        '    spritecol(X) = True
            display(location) = CRAM(colour + 16&)
        
        'Else
            
        '    status = status Or 32
        'End If
    End If
    

End Sub

'---------------------------------------~
'   Blank leftmost column of a scanline.
'   Replace with flllRect API.
'---------------------------------------~
Private Sub blankColumn(lineno As Long)

    Dim colour      As Long:    colour = CRAM(16 + (vdpreg(7) And &HF&))
    Dim row_precal  As Long:    row_precal = ShiftLeft(lineno, 8)
    Dim X           As Long
    
    For X = 8 To 0 Step -1
        display(X + row_precal) = colour
    Next X
    
End Sub

'---------------------------~
'   Generate VDP Interrupts.
'---------------------------~
Public Sub interrupts(lineno As Long)

    line = lineno
    
    If (lineno <= 192&) Then
    
        '// Frame Interrupt Pending
        If (lineno = 192&) Then
            status = status Or &H80&
            frameint = True
        End If
        
        '// Counter Expired = Line Interrupt Pendind
        If (counter = 0) Then
            counter = vdpreg(10)
            lineint = True
        Else
        '// Otherwise Decrement Counter
            counter = counter - 1
        End If
        
        '// Line Interrupts Enabled and Pending. Assert IRQ Line.
        If (lineint And ((vdpreg(0) And &H10&) = &H10&)) Then
            irqsetLine = True
        End If
    
    '// lineno >= 193
    Else
        
        '// Reload counter on every line outside active display + 1
        counter = vdpreg(10)
        
        '// Frame Interrupts Enabled and Pending. Assert IRQ Line.
        If (frameint And ((vdpreg(1) And &H20&) = &H20&) And (lineno < 224&)) Then irqsetLine = True
        
    End If
End Sub

'-----------------------------------~
'   Load Custom .pal (Palette) File.
'-----------------------------------~
Public Function loadPalette() As Boolean

    Dim X As Long, Y As Long, bytArray(192) As Byte


    '// Check Palette File Exists!
    If Dir$(App.Path & "\sms.pal", vbArchive) = vbNullString Then
        loadPalette = False
        MsgBox "Master System Color Palette File 'sms.pal' not found!", vbCritical, "Error"
        Exit Function
    End If


    '// Store SMS Palette File, into bytArray.
    Open App.Path & "\sms.pal" For Binary As #1
    Get #1, , bytArray()
    Close #1
    
    
    '// Convert Palette File to Long Color values.
    For X = 0 To 63
        SMS_PALETTE(X) = RGB(bytArray(Y), bytArray(Y + 1), bytArray(Y + 2))
        Y = Y + 3
    Next X
    
    
    '// Return True.
    loadPalette = True

End Function
