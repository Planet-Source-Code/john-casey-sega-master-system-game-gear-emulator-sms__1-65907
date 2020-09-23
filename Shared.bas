Attribute VB_Name = "Shared"

'=========================================================================+
'                             vbSMS Todo List                             |
'=========================================================================+
'
'   Make changes to incorporate GGear roms, this will involve detecting if
'   a certain rom is GG and using different VDP Palette and Display Size.
'
'   Fix existing bugs, SRAM Writes are Buggy "The Flash"
'   Vdp.GetVCount > is this function 100% correct "Spiderman Sinister 6"?
'   Add other VDP Modes "F-16 Fighter".
'   Fix CPU errors, (Doesn't pass ZEXALL).
'   Fix "Back to the Future III" (PAL Timing), maybe Spiderman Sinister 6.
'   Extended display - "Not only words".
'   Add Codemaster Mappers.
'   "Chuck Rock" Has incorrect colors at the start?
'   "Galactic Protector" Does the controller work?
'
'   Pause Button, Sound, Interface.
'   Improve Speed, Add CPU timing (So emulation runs at correct speed).
'   Are lookup tables any faster than the actually math?
'
'=========================================================================+
'                  [vbSMS, by John Casey <xshifu@msn.com>]                |
'=========================================================================+

Option Explicit

Public glMemAddrDiv256(65535)   As Long             'Lookup table for speed division.
Public ShiftRight(600, 1 To 10) As Long             'Lookup table for speed ShiftRight.
Public ShiftLeft(600, 1 To 10)  As Long             'Lookup table for speed ShiftLeft.

Public cartRom(49152)           As Byte             'ROM Pages, 0,1,2 (SaveRAM)
Public cartRam(8192)            As Byte             'Cart RAM.  (8K)

Public Const Page1              As Long = &H4000&   'Value For Page 1, (Memory Mapping).
Public Const Page2              As Long = &H8000&   'Value For Page 2, (Memory Mapping).
Public Mul4000()                As Long             'Lookup table for * &H4000&

Public pages()                  As Byte             'Holds ROM file data.
Public number_of_pages          As Long             'How many pages exist in a ROM.
Public frame_two_rom            As Long             'Does frame 2 use ROM or SRAM.
Public SRAM(16384)              As Byte             'Cartridge Ram Page 1 (Save RAM).

Public Const SMS_WIDTH          As Long = 256&      'SMS Screen width
Public Const SMS_HEIGHT         As Long = 192&      'SMS Screen height

Public controller1              As Long             'SMS Controller 1.



'------------------------------+
'          API CALLS           |
'------------------------------+
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb As Long)
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Type BITMAPINFOHEADER
        biSize                  As Long
        biWidth                 As Long
        biHeight                As Long
        biPlanes                As Integer
        biBitCount              As Integer
        biCompression           As Long
        biSizeImage             As Long
        biXPelsPerMeter         As Long
        biYPelsPerMeter         As Long
        biClrUsed               As Long
        biClrImportant          As Long
End Type

Public Type BITMAPINFO
        bmiHeader               As BITMAPINFOHEADER
        bmiColor0               As Long
        bmiColor1               As Long
        bmiColor2               As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public myBMP                    As BITMAPINFO
'------------------------------+
'           End Block          +
'------------------------------+




'--------------------------~
'   // Program Entry Point.
'--------------------------~
Private Sub Main()

    Dim X               As Long
    Dim Y               As Long
    Dim Z               As Long
    Dim bytArray(192)   As Byte


    If Not Vdp.loadPalette Then End     '// Load SMS Palette.
    controller1 = 255&                  '// Default Controller Value.
    frame_two_rom = True


    '// Calculate Div \ 256.
    For X = 0& To 65535
    glMemAddrDiv256(X) = (X \ 256&)
    Next X


    '// Calculate Flag Parity.
    For X = 0& To 255&
        Y = True
        For Z = 0& To 7&
            If (X And (2& ^ Z)) <> 0& Then Y = Not Y
        Next Z
        Parity(X) = Y
    Next X


    '// Calculate Bit Shifts.
    For X = 1& To 600&
        Z = 2
        For Y = 1& To 10&
                ShiftRight(X, Y) = (X \ Z)
                ShiftLeft(X, Y) = (X * Z)
                Z = Z * 2
        Next Y
    Next X

    
    '// Calculate Page Lookup. Upto 1MB.
    ReDim Mul4000(64)
    For X = 0 To 64
        Mul4000(X) = X * Page1
    Next X


    '// SMS Screen Setup.
    With myBMP.bmiHeader
        .biWidth = SMS_WIDTH
        .biHeight = -SMS_HEIGHT
        .biSize = 40
        .biBitCount = 32
        .biPlanes = 1
    End With
    
    
    Form1.Show                          '// Display Main Screen.

End Sub

'-----------------------------------~
'   // Read a cartridge into memory.
'-----------------------------------~
Public Function readCart(url As String) As Long

    Dim X           As Long
    Dim Y           As Long

    '// Does SMS Cart Exist?
    If Dir$(url, vbArchive) = vbNullString Then
        readCart = False
        Exit Function
    End If


    '// Read Cart into Memory.
    Open url For Binary As #1
    
        Y = LOF(1)                          '// Store File Length.
        number_of_pages = (Y / &H4000&)     '// Calculate Number of Pages.
        ReDim pages(Y)                      '// Redim Paging Array Accordingly.
        Get #1, , pages()                   '// Retrieve file data.

    Close #1
    
    
    '// Default ROM Mapping.
    If (Y >= &HC000&) Then Y = &HC000&
    CopyMemory cartRom(0), pages(0), Y


    readCart = True                         '// Return True.

End Function

'------------------------~
'   Output to a Z80 Port.
'------------------------~
Public Sub out(port As Long, value As Long)

    '----------------------------------------~
    '   Not all ports are coded yet, there's
    '   still Sound, and Auto. Nationalisation.
    '----------------------------------------~

    Select Case (port)
    
        Case &HBE&:     Vdp.dataWrite (value)       '// VDP Data Port.
        Case &HBF&:     Vdp.controlWrite (value)    '// VDP Control Port.
        Case &HBD&:     Vdp.controlWrite (value)    '// VDP Control Port.
        Case 0& To 5&:  Vdp.dataWrite (value)       '// GG Serial Ports.
        'Case &H7F&
        'Case Else:      Debug.Print "Unknown Port Out " & Hex(port)
        
    End Select

End Sub

'------------------------~
'   Read from a Z80 Port.
'------------------------~
Public Function inn(port As Long) As Long

    '----------------------------------------~
    '   Not all ports are coded, there's still
    '   Controller 2, Horizontal Port..
    '----------------------------------------~

    Select Case (port)
        
        Case &H7E&:     inn = Vdp.getVCount         '// Vertical Port.
        Case &HDC&:     inn = controller1           '// Controller 1.
        Case &HC0&:     inn = controller1           '// Controller 1. (Mirrored)
        Case &HBE&:     inn = Vdp.dataRead          '// VDP Data Port.
        Case &HBF&:     inn = Vdp.controlRead       '// VDP Control Port.
        
        Case Else:      inn = 255&                  '// Default Value.
                        'Debug.Print "!Unknown Port In " & Hex(port)
    End Select

End Function
