Attribute VB_Name = "Z80"
Option Explicit


'
'   Name:       Z80.bas (Zilog Z80 CPU Emulation).
'   Author:     Chris Cowley.
'               Modified for vbSMS by John Casey.
'   Version:    12th July 2006. Beta #3
'


'   //Registers.
Private regA                        As Long         'Accumulator.
Private regHL                       As Long         'General Purpose.
Private regB                        As Long         'Loop Counter.
Private regC                        As Long         'Loop Counter.
Private regDE                       As Long         'General Purpose.


'   //Alternate Registers.
Private regAF_                      As Long         'Alternate AF.
Private regHL_                      As Long         'Alternate HL.
Private regBC_                      As Long         'Alternate BC.
Private regDE_                      As Long         'Alternate DE.


'   //Index Registers.
Private regIX                       As Long         'Index Page Address Register.
Private regIY                       As Long         'Index Page Address Register.
Private regID                       As Long         'Temp for IX/IY.


'   //Flag Register.
Private fC                          As Long         'Carry.
Private fN                          As Long         'Negative.
Private fPV                         As Long         'Parity.
Private f3                          As Long         'Bit3.
Private fH                          As Long         'Half-Carry.
Private f5                          As Long         'Bit5.
Private fZ                          As Long         'Zero.
Private fS                          As Long         'Sign.


'   //Flag Bit Positions.
Private Const F_C                   As Long = 1     'Bit0 Carry.
Private Const F_N                   As Long = 2     'Bit1 Negative.
Private Const F_PV                  As Long = 4     'Bit2 Parity.
Private Const F_3                   As Long = 8     'Bit3.
Private Const F_H                   As Long = 16    'Bit4 Half-Carry.
Private Const F_5                   As Long = 32    'Bit5.
Private Const F_Z                   As Long = 64    'Bit6 Zero.
Private Const F_S                   As Long = 128   'Bit7 Sign.


'   //Interrupt / Refresh.
Private intI                        As Long         'Interrupt Page Address Register.
Private intR                        As Long         'Memory Refresh Register.
Private intRTemp                    As Long
Private intIFF1                     As Long         'Interrupt Flip Flop 1.
Private intIFF2                     As Long         'Interrupt Flip Flop 2.
Private intIM                       As Long         'Interrupt Mode (0,1,2).
Private halt                        As Long         'Cpu Is Currently In HALT.


'   //The Rest...
Private regSP                       As Long         'Stack Pointer.
Private regPC                       As Long         'Program Counter.
Private lineno                      As Long         'Current VDP Scanline.
Private Const Tcycles               As Long = 228   'Number of Tcycles Per Second.
Public Parity(256)                  As Long         'Parity Flag Lookup.
Public irqsetLine                   As Long         'Interrupt Line.


'   //API Declares.
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const QS_KEY As Long = &H1
Private Const QS_MOUSEBUTTON As Long = &H4
Private Const QS_SENDMESSAGE = &H40


Private Sub adc_a(b As Long)
    Dim wans As Long, ans As Long, c As Long
    
    If fC Then c = 1
    
    wans = regA + b + c
    ans = wans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fC = (wans And &H100&) <> 0
    fPV = ((regA Xor ((Not b) And &HFFFF&)) And (regA Xor ans) And &H80&) <> 0

    fH = (((regA And &HF&) + (b And &HF&) + c) And F_H) <> 0
    fN = False
     
    regA = ans
End Sub

Private Sub add_a(b As Long)
    Dim wans As Long, ans As Long
    
    wans = regA + b
    ans = wans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fC = (wans And &H100&) <> 0
    fPV = ((regA Xor ((Not (b)) And &HFFFF&)) And (regA Xor ans) And &H80&) <> 0
    fH = (((regA And &HF&) + (b And &HF&)) And F_H) <> 0
    fN = False
       
    regA = ans
End Sub
Private Function adc16(a As Long, b As Long) As Long
    Dim c As Long, lans As Long, ans As Long
    
    If fC Then c = 1
    
    lans = a + b + c
    ans = lans And &HFFFF&

    fS = (ans And (F_S * 256&)) <> 0
    f3 = (ans And (F_3 * 256&)) <> 0
    f5 = (ans And (F_5 * 256&)) <> 0
    fZ = (ans = 0)
    fC = (lans And &H10000) <> 0
    fPV = ((a Xor ((Not b) And &HFFFF&)) And (a Xor ans) And &H8000&) <> 0
    fH = (((a And &HFFF&) + (b And &HFFF&) + c) And &H1000&) <> 0
    fN = False
    
    adc16 = ans
End Function
Private Function add16(a As Long, b As Long) As Long
    Dim lans As Long
    Dim ans As Long
        
    lans = a + b
    ans = lans And &HFFFF&

    f3 = (ans And (F_3 * 256&)) <> 0
    f5 = (ans And (F_5 * 256&)) <> 0
    fC = (lans And &H10000) <> 0
    fH = (((a And &HFFF&) + (b And &HFFF&)) And &H1000&) <> 0
    fN = False
    
    add16 = ans
End Function
Private Sub and_a(b As Long)
    regA = (regA And b)
    
    fS = (regA And F_S) <> 0
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fH = True
    fPV = Parity(regA)
    fZ = (regA = 0)
    fN = False
    fC = False
End Sub
Private Sub bit(b As Long, r As Long)
    Dim IsbitSet As Long
    
    IsbitSet = (r And b) <> 0
    fN = False
    fH = True
    f3 = (r And F_3) <> 0
    f5 = (r And F_5) <> 0
    
    If b = F_S Then fS = IsbitSet Else fS = False
    
    fZ = Not IsbitSet
    fPV = fZ
End Sub
Private Function bitRes(bit As Long, val As Long) As Long
    bitRes = val And (Not (bit) And &HFFFF&)
End Function
Public Function bitSet(bit As Long, val As Long) As Long
    bitSet = val Or bit
End Function

Private Sub ccf()
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fH = fC
    fN = False
    fC = Not fC
End Sub
Public Sub cp_a(b As Long)
    Dim a As Long, wans As Long, ans As Long
    
    a = regA
    wans = a - b
    ans = wans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (b And F_3) <> 0
    f5 = (b And F_5) <> 0
    fN = True
    fZ = (ans = 0)
    fC = (wans And &H100&) <> 0
    fH = (((a And &HF&) - (b And &HF&)) And F_H) <> 0
    fPV = ((a Xor b) And (a Xor ans) And &H80&) <> 0
End Sub

Private Sub cpl_a()
    regA = (regA Xor &HFF&) And &HFF&
    
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fH = True
    fN = True
End Sub
Private Sub daa_a()
    Dim ans As Long, incr As Long, carry As Long
    
    ans = regA
    carry = fC
    
    If (fH = True) Or ((ans And &HF&) > &H9&) Then
        incr = incr Or &H6&
    End If
    
    If (carry = True) Or (ans > &H9F&) Then
        incr = incr Or &H60&
    End If
    
    If ((ans > &H8F&) And ((ans And &HF&) > 9&)) Then
        incr = incr Or &H60&
    End If
    
    If (ans > &H99&) Then
        carry = True
    End If
    If (fN = True) Then
        sub_a incr
    Else
        add_a incr
    End If
    
    ans = regA
    fC = carry
    fPV = Parity(ans)
End Sub
Private Function dec16(a As Long) As Long
    dec16 = (a - 1) And &HFFFF&
End Function
Private Sub ex_af_af()
    Dim t As Long
    
    t = getAF
    setAF regAF_
    regAF_ = t
End Sub
Private Function execute_cb() As Long
    Dim xxx As Long
    
    ' // Yes, I appreciate that GOTO's and labels are a hideous blashphemy!
    ' // However, this code is the fastest possible way of fetching and handling
    ' // Z80 instructions I could come up with. There are only 8 compares per
    ' // instruction fetch rather than between 1 and 255 as required in
    ' // the previous version of vb81 with it's huge Case statement.
    ' //
    ' // I know it's slightly harder to follow the new code, but I think the
    ' // speed increase justifies it. <CC>
    
    
    ' // REFRESH 1
    intRTemp = intRTemp + 1
    
    xxx = nxtpcb

    If (xxx And 128) Then GoTo ex_cb128_255 Else GoTo ex_cb0_127
    
ex_cb0_127:
    If (xxx And 64) Then GoTo ex_cb64_127 Else GoTo ex_cb0_63
    
ex_cb0_63:
    If (xxx And 32) Then GoTo ex_cb32_63 Else GoTo ex_cb0_31
    
ex_cb0_31:
    If (xxx And 16) Then GoTo ex_cb16_31 Else GoTo ex_cb0_15
    
ex_cb0_15:
    If (xxx And 8) Then GoTo ex_cb8_15 Else GoTo ex_cb0_7
    
ex_cb0_7:
    If (xxx And 4) Then GoTo ex_cb4_7 Else GoTo ex_cb0_3
    
ex_cb0_3:
    If (xxx And 2) Then GoTo ex_cb2_3 Else GoTo ex_cb0_1
    
ex_cb0_1:
    If xxx = 0 Then
        ' 000 RLC B
        regB = rlc(regB)
        execute_cb = 8
    Else
        ' 001 RLC C
        regC = rlc(regC)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb2_3:
    If xxx = 2 Then
        ' 002 RLC D
        setD rlc(getD)
        execute_cb = 8
    Else
        ' 003 RLC E
        setE rlc(getE)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb4_7:
    If (xxx And 2) Then GoTo ex_cb6_7 Else GoTo ex_cb4_5
    
ex_cb4_5:
    If xxx = 4 Then
        ' 004 RLC H
        setH rlc(getH)
        execute_cb = 8
    Else
        ' 005 RLC L
        setL rlc(getL)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb6_7:
    If xxx = 6 Then
        ' 006 RLC (HL)
        pokeb regHL, rlc(peekb(regHL))
        execute_cb = 15
    Else
        ' 007 RLC A
        regA = rlc(regA)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb8_15:
    If (xxx And 4) Then GoTo ex_cb12_15 Else GoTo ex_cb8_11
    
ex_cb8_11:
    If (xxx And 2) Then GoTo ex_cb10_11 Else GoTo ex_cb8_9
    
ex_cb8_9:
    If xxx = 8 Then
        ' 008 RRC B
        regB = rrc(regB)
        execute_cb = 8
    Else
        ' 009 RRC C
        regC = rrc(regC)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb10_11:
    If xxx = 10 Then
        ' 010 RRC D
        setD rrc(getD)
        execute_cb = 8
    Else
        ' 011 RRC E
        setE rrc(getE)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb12_15:
    If (xxx And 2) Then GoTo ex_cb14_15 Else GoTo ex_cb12_13
    
ex_cb12_13:
    If xxx = 12 Then
        ' 012 RRC H
        setH rrc(getH)
        execute_cb = 8
    Else
        ' 013 RRC L
        setL rrc(getL)
        execute_cb = 8
    End If
    Exit Function

ex_cb14_15:
    If xxx = 14 Then
        ' 014 RRC (HL)
        pokeb regHL, rrc(peekb(regHL))
        execute_cb = 15
    Else
        ' 015 RRC A
        regA = rrc(regA)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb16_31:
    If (xxx And 8) Then GoTo ex_cb24_31 Else GoTo ex_cb16_23
    
ex_cb16_23:
    If (xxx And 4) Then GoTo ex_cb20_23 Else GoTo ex_cb16_19
    
ex_cb16_19:
    If (xxx And 2) Then GoTo ex_cb18_19 Else GoTo ex_cb16_17
    
ex_cb16_17:
    If xxx = 16 Then
        ' 016 RL B
        regB = rl(regB)
        execute_cb = 8
    Else
        ' 017 RL C
        regC = rl(regC)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb18_19:
    If xxx = 18 Then
        ' 018 RL D
        setD rl(getD)
        execute_cb = 8
    Else
        ' 019 RL E
        setE rl(getE)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb20_23:
    If (xxx And 2) Then GoTo ex_cb22_23 Else GoTo ex_cb20_21
    
ex_cb20_21:
    If xxx = 20 Then
        ' 020 RL H
        setH rl(getH)
        execute_cb = 8
    Else
        ' 021 RL L
        setL rl(getL)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb22_23:
    If xxx = 22 Then
        ' 022 RL (HL)
        pokeb regHL, rl(peekb(regHL))
        execute_cb = 15
    Else
        ' 023 RL A
        regA = rl(regA)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb24_31:
    If (xxx And 4) Then GoTo ex_cb28_31 Else GoTo ex_cb24_27

ex_cb24_27:
    If (xxx And 2) Then GoTo ex_cb26_27 Else GoTo ex_cb24_25
    
ex_cb24_25:
    If xxx = 24 Then
        ' 024 RR B
        regB = rr(regB)
        execute_cb = 8
    Else
        ' 025 RR C
        regC = rr(regC)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb26_27:
    If xxx = 26 Then
        ' 026 RR D
        setD rr(getD)
        execute_cb = 8
    Else
        ' 027 RR E
        setE rr(getE)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb28_31:
    If (xxx And 2) Then GoTo ex_cb30_31 Else GoTo ex_cb28_29
    
ex_cb28_29:
    If xxx = 28 Then
        ' 028 RR H
        setH rr(getH)
        execute_cb = 8
    Else
        ' 029 RR L
        setL rr(getL)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb30_31:
    If xxx = 30 Then
        ' 030 RR (HL)
        pokeb regHL, rr(peekb(regHL))
        execute_cb = 15
    Else
        ' 031 RR A
        regA = rr(regA)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb32_63:
    Select Case xxx
    Case 32 ' SLA B
        regB = sla(regB)
        execute_cb = 8
    Case 33 ' SLA C
        regC = sla(regC)
        execute_cb = 8
    Case 34 ' SLA D
        setD sla(getD)
        execute_cb = 8
    Case 35 ' SLA E
        setE sla(getE)
        execute_cb = 8
    Case 36 ' SLA H
        setH sla(getH)
        execute_cb = 8
    Case 37 ' SLA L
        setL sla(getL)
        execute_cb = 8
    Case 38 ' SLA (HL)
        pokeb regHL, sla(peekb(regHL))
        execute_cb = 15
    Case 39 ' SLA A
        regA = sla(regA)
        execute_cb = 8
    Case 40 ' SRA B
        regB = sra(regB)
        execute_cb = 8
    Case 41 ' SRA C
        regC = sra(regC)
        execute_cb = 8
    Case 42 ' SRA D
        setD sra(getD)
        execute_cb = 8
    Case 43 ' SRA E
        setE sra(getE)
        execute_cb = 8
    Case 44 ' SRA H
        setH sra(getH)
        execute_cb = 8
    Case 45  ' SRA L
        setL sra(getL)
        execute_cb = 8
    Case 46 ' SRA (HL)
        pokeb regHL, sra(peekb(regHL))
        execute_cb = 15
    Case 47 ' SRA A
        regA = sra(regA)
        execute_cb = 8
    Case 48 ' SLS B
        regB = sls(regB)
        execute_cb = 8
    Case 49 ' SLS C
        regC = sls(regC)
        execute_cb = 8
    Case 50 ' SLS D
        setD sls(getD)
        execute_cb = 8
    Case 51 ' SLS E
        setE sls(getE)
        execute_cb = 8
    Case 52 ' SLS H
        setH sls(getH)
        execute_cb = 8
    Case 53 ' SLS L
        setL sls(getL)
        execute_cb = 8
    Case 54 ' SLS (HL)
        pokeb regHL, sls(peekb(regHL))
        execute_cb = 15
    Case 55 ' SLS A
        regA = sls(regA)
        execute_cb = 8
    Case 56 ' SRL B
        regB = srl(regB)
        execute_cb = 8
    Case 57 ' SRL C
        regC = srl(regC)
        execute_cb = 8
    Case 58 ' SRL D
        setD srl(getD)
        execute_cb = 8
    Case 59 ' SRL E
        setE srl(getE)
        execute_cb = 8
    Case 60 ' SRL H
        setH srl(getH)
        execute_cb = 8
    Case 61 ' SRL L
        setL srl(getL)
        execute_cb = 8
    Case 62 ' SRL (HL)
        pokeb regHL, srl(peekb(regHL))
        execute_cb = 15
    Case 63 ' SRL A
        regA = srl(regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb64_127:
    If (xxx And 32) Then GoTo ex_cb96_127 Else GoTo ex_cb64_95
    
ex_cb64_95:
    If (xxx And 16) Then GoTo ex_cb80_95 Else GoTo ex_cb64_79
    
ex_cb64_79:
    If (xxx And 8) Then GoTo ex_cb72_79 Else GoTo ex_cb64_71
    
ex_cb64_71:
    If (xxx And 4) Then GoTo ex_cb68_71 Else GoTo ex_cb64_67
    
ex_cb64_67:
    If (xxx And 2) Then GoTo ex_cb66_67 Else GoTo ex_cb64_65
    
ex_cb64_65:
    If xxx = 64 Then
        ' 064 BIT 0,B
        bit &H1&, regB
        execute_cb = 8
    Else
        ' 065 ' BIT 0,C
        bit 1&, regC
        execute_cb = 8
    End If
    Exit Function
    
ex_cb66_67:
    If xxx = 66 Then
        ' 066 BIT 0,D
        bit 1&, getD
        execute_cb = 8
    Else
        ' 067 BIT 0,E
        bit 1&, getE
        execute_cb = 8
    End If
    Exit Function
    
ex_cb68_71:
    If (xxx And 2) Then GoTo ex_cb70_71 Else GoTo ex_cb68_69
    
ex_cb68_69:
    If xxx = 68 Then
        ' 068 BIT 0,H
        bit 1&, getH
        execute_cb = 8
    Else
        ' 069 BIT 0,L
        bit 1&, getL
        execute_cb = 8
    End If
    Exit Function
    
ex_cb70_71:
    If xxx = 70 Then
        ' 070 BIT 0,(HL)
        bit 1&, peekb(regHL)
        execute_cb = 12
    Else
        ' 071 BIT 0,A
        bit 1&, regA
        execute_cb = 8
    End If
    Exit Function
    
ex_cb72_79:
    Select Case xxx
    Case 72 ' BIT 1,B
        bit 2&, regB
        execute_cb = 8
    Case 73 ' BIT 1,C
        bit 2&, regC
        execute_cb = 8
    Case 74 ' BIT 1,D
        bit 2&, getD
        execute_cb = 8
    Case 75 ' BIT 1,E
        bit 2&, getE
        execute_cb = 8
    Case 76 ' BIT 1,H
        bit 2&, getH
        execute_cb = 8
    Case 77 ' BIT 1,L
        bit 2&, getL
        execute_cb = 8
    Case 78 ' BIT 1,(HL)
        bit 2&, peekb(regHL)
        execute_cb = 12
    Case 79 ' BIT 1,A
        bit 2&, regA
        execute_cb = 8
    End Select
    Exit Function

ex_cb80_95:
    Select Case xxx
    Case 80 ' BIT 2,B
        bit 4&, regB
        execute_cb = 8
    Case 81 ' BIT 2,C
        bit 4&, regC
        execute_cb = 8
    Case 82 ' BIT 2,D
        bit 4&, getD
        execute_cb = 8
    Case 83 ' BIT 2,E
        bit 4&, getE
        execute_cb = 8
    Case 84 ' BIT 2,H
        bit 4&, getH
        execute_cb = 8
    Case 85 ' BIT 2,L
        bit 4&, getL
        execute_cb = 8
    Case 86 ' BIT 2,(HL)
        bit 4&, peekb(regHL)
        execute_cb = 12
    Case 87 ' BIT 2,A
        bit 4&, regA
        execute_cb = 8
    Case 88 ' BIT 3,B
        bit 8&, regB
        execute_cb = 8
    Case 89 ' BIT 3,C
        bit 8&, regC
        execute_cb = 8
    Case 90 ' BIT 3,D
        bit 8&, getD
        execute_cb = 8
    Case 91 ' BIT 3,E
        bit 8&, getE
        execute_cb = 8
    Case 92 ' BIT 3,H
        bit 8&, getH
        execute_cb = 8
    Case 93 ' BIT 3,L
        bit 8&, getL
        execute_cb = 8
    Case 94 ' BIT 3,(HL)
        bit 8&, peekb(regHL)
        execute_cb = 12
    Case 95 ' BIT 3,A
        bit 8&, regA
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb96_127:
    If (xxx And 16) Then GoTo ex_cb112_127 Else GoTo ex_cb96_111
    
ex_cb96_111:
    If (xxx And 8) Then GoTo ex_cb104_111 Else GoTo ex_cb96_103
    
ex_cb96_103:
    Select Case xxx
    Case 96 ' BIT 4,B
        bit &H10&, regB
        execute_cb = 8
    Case 97 ' BIT 4,C
        bit &H10&, regC
        execute_cb = 8
    Case 98 ' BIT 4,D
        bit &H10&, getD
        execute_cb = 8
    Case 99 ' BIT 4,E
        bit &H10&, getE
        execute_cb = 8
    Case 100 ' BIT 4,H
        bit &H10&, getH
        execute_cb = 8
    Case 101 ' BIT 4,L
        bit &H10&, getL
        execute_cb = 8
    Case 102 ' BIT 4,(HL)
        bit &H10&, peekb(regHL)
        execute_cb = 12
    Case 103 ' BIT 4,A
        bit &H10&, regA
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb104_111:
    Select Case xxx
    Case 104 ' BIT 5,B
        bit &H20&, regB
        execute_cb = 8
    Case 105 ' BIT 5,C
        bit &H20&, regC
        execute_cb = 8
    Case 106 ' BIT 5,D
        bit &H20&, getD
        execute_cb = 8
    Case 107 ' BIT 5,E
        bit &H20&, getE
        execute_cb = 8
    Case 108 ' BIT 5,H
        bit &H20&, getH
        execute_cb = 8
    Case 109 ' BIT 5,L
        bit &H20&, getL
        execute_cb = 8
    Case 110 ' BIT 5,(HL)
        bit &H20&, peekb(regHL)
        execute_cb = 12
    Case 111 ' BIT 5,A
        bit &H20&, regA
        execute_cb = 8
    End Select
    Exit Function

ex_cb112_127:
    If (xxx And 8) Then GoTo ex_cb120_127 Else GoTo ex_cb112_119
    
ex_cb112_119:
    If (xxx And 4) Then GoTo ex_cb116_119 Else GoTo ex_cb112_115
    
ex_cb112_115:
    If (xxx And 2) Then GoTo ex_cb114_115 Else GoTo ex_cb112_113
    
ex_cb112_113:
    If xxx = 112 Then
        ' 112 BIT 6,B
        bit &H40&, regB
        execute_cb = 8
    Else
        ' 113 BIT 6,C
        bit &H40&, regC
        execute_cb = 8
    End If
    Exit Function

ex_cb114_115:
    If xxx = 114 Then
        ' 114 BIT 6,D
        bit &H40&, getD
        execute_cb = 8
    Else
        ' 115 BIT 6,E
        bit &H40&, getE
        execute_cb = 8
    End If
    Exit Function

ex_cb116_119:
    If (xxx And 2) Then GoTo ex_cb118_119 Else GoTo ex_cb116_117
    
ex_cb116_117:
    If xxx = 116 Then
        ' 116 BIT 6,H
        bit &H40&, getH
        execute_cb = 8
    Else
        ' 117 BIT 6,L
        bit &H40&, getL
        execute_cb = 8
    End If
    Exit Function

ex_cb118_119:
    If xxx = 118 Then
        ' 118 BIT 6,(HL)
        bit &H40&, peekb(regHL)
        execute_cb = 12
    Else
        ' 119 ' BIT 6,A
        bit &H40&, regA
        execute_cb = 8
    End If
    Exit Function
    
ex_cb120_127:
    If (xxx And 4) Then GoTo ex_cb124_127 Else GoTo ex_cb120_123
    
ex_cb120_123:
    If (xxx And 2) Then GoTo ex_cb122_123 Else GoTo ex_cb120_121
    
ex_cb120_121:
    If xxx = 120 Then
        ' 120 BIT 7,B
        bit &H80&, regB
        execute_cb = 8
    Else
        ' 121 BIT 7,C
        bit &H80&, regC
        execute_cb = 8
    End If
    Exit Function
    
ex_cb122_123:
    If xxx = 122 Then
        ' 122 BIT 7,D
        bit &H80&, getD
        execute_cb = 8
    Else
        ' 123 BIT 7,E
        bit &H80&, getE
        execute_cb = 8
    End If
    Exit Function

ex_cb124_127:
    If (xxx And 2) Then GoTo ex_cb126_127 Else GoTo ex_cb124_125
    
ex_cb124_125:
    If xxx = 124 Then
        ' 124 BIT 7,H
        bit &H80&, getH
        execute_cb = 8
    Else
        ' 125 BIT 7,L
        bit &H80&, getL
        execute_cb = 8
    End If
    Exit Function

ex_cb126_127:
    If xxx = 126 Then
        ' 126 BIT 7,(HL)
        bit &H80&, peekb(regHL)
        execute_cb = 12
    Else
        ' 127 BIT 7,A
        bit &H80&, regA
        execute_cb = 8
    End If
    Exit Function
    
ex_cb128_255:
    If (xxx And 64) Then GoTo ex_cb192_255 Else GoTo ex_cb128_191
    
ex_cb128_191:
    If (xxx And 32) Then GoTo ex_cb160_191 Else GoTo ex_cb128_159
    
ex_cb128_159:
    If (xxx And 16) Then GoTo ex_cb144_159 Else GoTo ex_cb128_143
    
ex_cb128_143:
    Select Case xxx
    Case 128 ' RES 0,B
        regB = bitRes(1&, regB)
        execute_cb = 8
    Case 129 ' RES 0,C
        regC = bitRes(1&, regC)
        execute_cb = 8
    Case 130 ' RES 0,D
        setD bitRes(1&, getD)
        execute_cb = 8
    Case 131 ' RES 0,E
        setE bitRes(1&, getE)
        execute_cb = 8
    Case 132 ' RES 0,H
        setH bitRes(1&, getH)
        execute_cb = 8
    Case 133 ' RES 0,L
        setL bitRes(1&, getL)
        execute_cb = 8
    Case 134 ' RES 0,(HL)
        pokeb regHL, bitRes(&H1&, peekb(regHL))
        execute_cb = 15
    Case 135 ' RES 0,A
        regA = bitRes(1&, regA)
        execute_cb = 8
    Case 136 ' RES 1,B
        regB = bitRes(2&, regB)
        execute_cb = 8
    Case 137 ' RES 1,C
        regC = bitRes(2&, regC)
        execute_cb = 8
    Case 138 ' RES 1,D
        setD bitRes(2&, getD)
        execute_cb = 8
    Case 139 ' RES 1,E
        setE bitRes(2&, getE)
        execute_cb = 8
    Case 140 ' RES 1,H
        setH bitRes(2&, getH)
        execute_cb = 8
    Case 141 ' RES 1,L
        setL bitRes(2&, getL)
        execute_cb = 8
    Case 142 ' RES 1,(HL)
        pokeb regHL, bitRes(2&, peekb(regHL))
        execute_cb = 15
    Case 143 ' RES 1,A
        regA = bitRes(2&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb144_159:
    Select Case xxx
    Case 144 ' RES 2,B
        regB = bitRes(4&, regB)
        execute_cb = 8
    Case 145 ' RES 2,C
        regC = bitRes(4&, regC)
        execute_cb = 8
    Case 146 ' RES 2,D
        setD bitRes(4&, getD)
        execute_cb = 8
    Case 147 ' RES 2,E
        setE bitRes(4&, getE)
        execute_cb = 8
    Case 148 ' RES 2,H
        setH bitRes(4&, getH)
        execute_cb = 8
    Case 149 ' RES 2,L
        setL bitRes(4&, getL)
        execute_cb = 8
    Case 150 ' RES 2,(HL)
        pokeb regHL, bitRes(4&, peekb(regHL))
        execute_cb = 15
    Case 151 ' RES 2,A
        regA = bitRes(4&, regA)
        execute_cb = 8
    Case 152 ' RES 3,B
        regB = bitRes(8&, regB)
        execute_cb = 8
    Case 153 ' RES 3,C
        regC = bitRes(8&, regC)
        execute_cb = 8
    Case 154 ' RES 3,D
        setD bitRes(8&, getD)
        execute_cb = 8
    Case 155 ' RES 3,E
        setE bitRes(8&, getE)
        execute_cb = 8
    Case 156 ' RES 3,H
        setH bitRes(8&, getH)
        execute_cb = 8
    Case 157 ' RES 3,L
        setL bitRes(8&, getL)
        execute_cb = 8
    Case 158 ' RES 3,(HL)
        pokeb regHL, bitRes(8&, peekb(regHL))
        execute_cb = 15
    Case 159 ' RES 3,A
        regA = bitRes(8&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb160_191:
    If (xxx And 16) Then GoTo ex_cb176_191 Else GoTo ex_cb160_175
    
ex_cb160_175:
    If (xxx And 8) Then GoTo ex_cb168_175 Else GoTo ex_cb160_167
    
ex_cb160_167:
    Select Case xxx
    Case 160 ' RES 4,B
        regB = bitRes(&H10&, regB)
        execute_cb = 8
    Case 161 ' RES 4,C
        regC = bitRes(&H10&, regC)
        execute_cb = 8
    Case 162 ' RES 4,D
        setD bitRes(&H10&, getD)
        execute_cb = 8
    Case 163 ' RES 4,E
        setE bitRes(&H10&, getE)
        execute_cb = 8
    Case 164 ' RES 4,H
        setH bitRes(&H10&, getH)
        execute_cb = 8
    Case 165 ' RES 4,L
        setL bitRes(&H10&, getL)
        execute_cb = 8
    Case 166 ' RES 4,(HL)
        pokeb regHL, bitRes(&H10&, peekb(regHL))
        execute_cb = 15
    Case 167 ' RES 4,A
        regA = bitRes(&H10&, regA)
        execute_cb = 8
    End Select
    Exit Function

ex_cb168_175:
    If (xxx And 4) Then GoTo ex_cb172_175 Else GoTo ex_cb168_171
    
ex_cb168_171:
    If (xxx And 2) Then GoTo ex_cb170_171 Else GoTo ex_cb168_169
    
ex_cb168_169:
    If xxx = 168 Then
        ' 168 RES 5,B
        regB = bitRes(&H20&, regB)
        execute_cb = 8
    Else
        ' 169 RES 5,C
        regC = bitRes(&H20&, regC)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb170_171:
    If xxx = 170 Then
        ' 170 RES 5,D
        setD bitRes(&H20&, getD)
        execute_cb = 8
    Else
        ' 171 RES 5,E
        setE bitRes(&H20&, getE)
        execute_cb = 8
    End If
    Exit Function
    
ex_cb172_175:
    Select Case xxx
    Case 172 ' RES 5,H
        setH bitRes(&H20&, getH)
        execute_cb = 8
    Case 173 ' RES 5,L
        setL bitRes(&H20&, getL)
        execute_cb = 8
    Case 174 ' RES 5,(HL)
        pokeb regHL, bitRes(&H20&, peekb(regHL))
        execute_cb = 15
    Case 175 ' RES 5,A
        regA = bitRes(&H20&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb176_191:
    Select Case xxx
    Case 176 ' RES 6,B
        regB = bitRes(&H40&, regB)
        execute_cb = 8
    Case 177 ' RES 6,C
        regC = bitRes(&H40&, regC)
        execute_cb = 8
    Case 178 ' RES 6,D
        setD bitRes(&H40&, getD)
        execute_cb = 8
    Case 179 ' RES 6,E
        setE bitRes(&H40&, getE)
        execute_cb = 8
    Case 180 ' RES 6,H
        setH bitRes(&H40&, getH)
        execute_cb = 8
    Case 181 ' RES 6,L
        setL bitRes(&H40&, getL)
        execute_cb = 8
    Case 182 ' RES 6,(HL)
        pokeb regHL, bitRes(&H40&, peekb(regHL))
        execute_cb = 15
    Case 183 ' RES 6,A
        regA = bitRes(&H40&, regA)
        execute_cb = 8
    Case 184 ' RES 7,B
        regB = bitRes(&H80&, regB)
        execute_cb = 8
    Case 185 ' RES 7,C
        regC = bitRes(&H80&, regC)
        execute_cb = 8
    Case 186 ' RES 7,D
        setD bitRes(&H80&, getD)
        execute_cb = 8
    Case 187 ' RES 7,E
        setE bitRes(&H80&, getE)
        execute_cb = 8
    Case 188 ' RES 7,H
        setH bitRes(&H80&, getH)
        execute_cb = 8
    Case 189 ' RES 7,L
        setL bitRes(&H80&, getL)
        execute_cb = 8
    Case 190 ' RES 7,(HL)
        pokeb regHL, bitRes(&H80&, peekb(regHL))
        execute_cb = 15
    Case 191 ' RES 7,A
        regA = bitRes(&H80&, regA)
        execute_cb = 8
    End Select
    Exit Function

ex_cb192_255:
    If (xxx And 32) Then GoTo ex_cb224_255 Else GoTo ex_cb192_223
    
ex_cb192_223:
    If (xxx And 16) Then GoTo ex_cb208_223 Else GoTo ex_cb192_207
    
ex_cb192_207:
    If (xxx And 8) Then GoTo ex_cb200_207 Else GoTo ex_cb192_199
    
ex_cb192_199:
    Select Case xxx
    Case 192 ' SET 0,B
        regB = bitSet(1&, regB)
        execute_cb = 8
    Case 193 ' SET 0,C
        regC = bitSet(1&, regC)
        execute_cb = 8
    Case 194 ' SET 0,D
        setD bitSet(1&, getD)
        execute_cb = 8
    Case 195 ' SET 0,E
        setE bitSet(1&, getE)
        execute_cb = 8
    Case 196 ' SET 0,H
        setH bitSet(1&, getH)
        execute_cb = 8
    Case 197 ' SET 0,L
        setL bitSet(1&, getL)
        execute_cb = 8
    Case 198 ' SET 0,(HL)
        pokeb regHL, bitSet(1&, peekb(regHL))
        execute_cb = 15
    Case 199 ' SET 0,A
        regA = bitSet(1&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb200_207:
    Select Case xxx
    Case 200 ' SET 1,B
        regB = bitSet(2&, regB)
        execute_cb = 8
    Case 201 ' SET 1,C
        regC = bitSet(2&, regC)
        execute_cb = 8
    Case 202 ' SET 1,D
        setD bitSet(2&, getD)
        execute_cb = 8
    Case 203 ' SET 1,E
        setE bitSet(2&, getE)
        execute_cb = 8
    Case 204 ' SET 1,H
        setH bitSet(2&, getH)
        execute_cb = 8
    Case 205 ' SET 1,L
        setL bitSet(2&, getL)
        execute_cb = 8
    Case 206 ' SET 1,(HL)
        pokeb regHL, bitSet(2&, peekb(regHL))
        execute_cb = 15
    Case 207 ' SET 1,A
        regA = bitSet(2&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb208_223:
    Select Case xxx
    Case 208 ' SET 2,B
        regB = bitSet(4&, regB)
        execute_cb = 8
    Case 209 ' SET 2,C
        regC = bitSet(4&, regC)
        execute_cb = 8
    Case 210 ' SET 2,D
        setD bitSet(4&, getD)
        execute_cb = 8
    Case 211 ' SET 2,E
        setE bitSet(4&, getE)
        execute_cb = 8
    Case 212 ' SET 2,H
        setH bitSet(4&, getH)
        execute_cb = 8
    Case 213 ' SET 2,L
        setL bitSet(4&, getL)
        execute_cb = 8
    Case 214 ' SET 2,(HL)
        pokeb regHL, bitSet(&H4&, peekb(regHL))
        execute_cb = 15
    Case 215 ' SET 2,A
        regA = bitSet(4&, regA)
        execute_cb = 8
    Case 216 ' SET 3,B
        regB = bitSet(8&, regB)
        execute_cb = 8
    Case 217 ' SET 3,C
        regC = bitSet(8&, regC)
        execute_cb = 8
    Case 218 ' SET 3,D
        setD bitSet(8&, getD)
        execute_cb = 8
    Case 219 ' SET 3,E
        setE bitSet(8&, getE)
        execute_cb = 8
    Case 220 ' SET 3,H
        setH bitSet(8&, getH)
        execute_cb = 8
    Case 221 ' SET 3,L
        setL bitSet(8&, getL)
        execute_cb = 8
    Case 222 ' SET 3,(HL)
        pokeb regHL, bitSet(&H8&, peekb(regHL))
        execute_cb = 15
    Case 223 ' SET 3,A
        regA = bitSet(8&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb224_255:
    If (xxx And 16) Then GoTo ex_cb240_255 Else GoTo ex_cb224_239
    
ex_cb224_239:
    Select Case xxx
    Case 224 ' SET 4,B
        regB = bitSet(&H10&, regB)
        execute_cb = 8
    Case 225 ' SET 4,C
        regC = bitSet(&H10&, regC)
        execute_cb = 8
    Case 226 ' SET 4,D
        setD bitSet(&H10&, getD)
        execute_cb = 8
    Case 227 ' SET 4,E
        setE bitSet(&H10&, getE)
        execute_cb = 8
    Case 228 ' SET 4,H
        setH bitSet(&H10&, getH)
        execute_cb = 8
    Case 229 ' SET 4,L
        setL bitSet(&H10&, getL)
        execute_cb = 8
    Case 230 ' SET 4,(HL)
        pokeb regHL, bitSet(&H10&, peekb(regHL))
        execute_cb = 15
    Case 231 ' SET 4,A
        regA = bitSet(&H10&, regA)
        execute_cb = 8
    Case 232 ' SET 5,B
        regB = bitSet(&H20&, regB)
        execute_cb = 8
    Case 233 ' SET 5,C
        regC = bitSet(&H20&, regC)
        execute_cb = 8
    Case 234 ' SET 5,D
        setD bitSet(&H20&, getD)
        execute_cb = 8
    Case 235 ' SET 5,E
        setE bitSet(&H20&, getE)
        execute_cb = 8
    Case 236 ' SET 5,H
        setH bitSet(&H20&, getH)
        execute_cb = 8
    Case 237 ' SET 5,L
        setL bitSet(&H20&, getL)
        execute_cb = 8
    Case 238 ' SET 5,(HL)
        pokeb regHL, bitSet(&H20&, peekb(regHL))
        execute_cb = 15
    Case 239 ' SET 5,A
        regA = bitSet(&H20&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb240_255:
    If (xxx And 8) Then GoTo ex_cb248_255 Else GoTo ex_cb240_247
    
ex_cb240_247:
    Select Case xxx
    Case 240 ' SET 6,B
        regB = bitSet(&H40&, regB)
        execute_cb = 8
    Case 241 ' SET 6,C
        regC = bitSet(&H40&, regC)
        execute_cb = 8
    Case 242 ' SET 6,D
        setD bitSet(&H40&, getD)
        execute_cb = 8
    Case 243 ' SET 6,E
        setE bitSet(&H40&, getE)
        execute_cb = 8
    Case 244 ' SET 6,H
        setH bitSet(&H40&, getH)
        execute_cb = 8
    Case 245 ' SET 6,L
        setL bitSet(&H40&, getL)
        execute_cb = 8
    Case 246 ' SET 6,(HL)
        pokeb regHL, bitSet(&H40&, peekb(regHL))
        execute_cb = 15
    Case 247 ' SET 6,A
        regA = bitSet(&H40&, regA)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb248_255:
    If (xxx And 4) Then GoTo ex_cb252_255 Else GoTo ex_cb248_251
    
ex_cb248_251:
    Select Case xxx
    Case 248 ' SET 7,B
        regB = bitSet(&H80&, regB)
        execute_cb = 8
    Case 249 ' SET 7,C
        regC = bitSet(&H80&, regC)
        execute_cb = 8
    Case 250 ' SET 7,D
        setD bitSet(&H80&, getD)
        execute_cb = 8
    Case 251 ' SET 7,E
        setE bitSet(&H80&, getE)
        execute_cb = 8
    End Select
    Exit Function
    
ex_cb252_255:
    If (xxx And 2) Then GoTo ex_cb254_255 Else GoTo ex_cb252_253
    
ex_cb252_253:
    If xxx = 252 Then
        ' 252 SET 7,H
        setH bitSet(&H80&, getH)
        execute_cb = 8
    Else
        ' 253 SET 7,L
        setL bitSet(&H80&, getL)
        execute_cb = 8
    End If
    Exit Function

ex_cb254_255:
    If xxx = 254 Then
        ' 254 SET 7,(HL)
        pokeb regHL, bitSet(&H80&, peekb(regHL))
        execute_cb = 15
    Else
        ' 255 SET 7,A
        regA = bitSet(&H80&, regA)
        execute_cb = 8
    End If
End Function

Function sla(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H80&) <> 0
    ans = (ans * 2) And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
    
    sla = ans
End Function

Private Function execute_ed(ByVal local_tstates As Long) As Long
    Dim xxx As Long, count As Long, dest As Long, from As Long
    Dim TempLocal_tstates As Long, c As Long, b As Long
        
    ' // Yes, I appreciate that GOTO's and labels are a hideous blashphemy!
    ' // However, this code is the fastest possible way of fetching and handling
    ' // Z80 instructions I could come up with. There are only 8 compares per
    ' // instruction fetch rather than between 1 and 255 as required in
    ' // the previous version of vb81 with it's huge Case statement.
    ' //
    ' // I know it's slightly harder to follow the new code, but I think the
    ' // speed increase justifies it. <CC>
    
    
    ' // REFRESH 1
    intRTemp = intRTemp + 1
    
    xxx = nxtpcb
    
    If (xxx And 128) Then GoTo ex_ed128_255 Else GoTo ex_ed0_127

ex_ed0_127:
    If (xxx And 64) Then
        GoTo ex_ed64_127
    Else
        ' 000 to 063 = NOP
        execute_ed = 8
        Exit Function
    End If
    
ex_ed64_127:
    If (xxx And 32) Then GoTo ex_ed96_127 Else GoTo ex_ed64_95
    
ex_ed64_95:
    If (xxx And 16) Then GoTo ex_ed80_95 Else GoTo ex_ed64_79
    
ex_ed64_79:
    If (xxx And 8) Then GoTo ex_ed72_79 Else GoTo ex_ed64_71
    
ex_ed64_71:
    If (xxx And 4) Then GoTo ex_ed68_71 Else GoTo ex_ed64_67
    
ex_ed64_67:
    If (xxx And 2) Then GoTo ex_ed66_67 Else GoTo ex_ed64_65

ex_ed64_65:
    If xxx = 64 Then
        ' 064 IN B,(c)
        regB = inn(regC) 'in_bc()
        execute_ed = 12
    Else
        ' 065 OUT (c),B
        out regC, regB
        execute_ed = 12
    End If
    Exit Function
    
ex_ed66_67:
    If xxx = 66 Then
        ' 066 SBC HL,BC
        regHL = sbc16(regHL, getBC)
        execute_ed = 15
    Else
        ' 067 LD (nn),BC
        pokew nxtpcw(), getBC
        execute_ed = 20
    End If
    Exit Function
    
ex_ed68_71:
    If (xxx And 2) Then GoTo ex_ed70_71 Else GoTo ex_ed68_69
    
ex_ed68_69:
    If xxx = 68 Then
        ' 068 NEG
        neg_a
        execute_ed = 8
    Else
        ' 069 RETn
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    End If
    Exit Function
    
ex_ed70_71:
    If xxx = 70 Then
        ' 070 IM 0
        intIM = 0
        execute_ed = 8
    Else
        ' 071 LD I,A
        intI = regA
        execute_ed = 9
    End If
    Exit Function
    
ex_ed72_79:
    If (xxx And 4) Then GoTo ex_ed76_79 Else GoTo ex_ed72_75
    
ex_ed72_75:
    If (xxx And 2) Then GoTo ex_ed74_75 Else GoTo ex_ed72_73

ex_ed72_73:
    If xxx = 72 Then
        ' 072 IN C,(c)
        regC = inn(regC) 'in_bc()
        execute_ed = 12
    Else
        ' 073 OUT (c),C
        out regC, regC
        execute_ed = 12
    End If
    Exit Function
    
ex_ed74_75:
    If xxx = 74 Then
        ' 074 ADC HL,BC
        regHL = adc16(regHL, getBC)
        execute_ed = 15
    Else
        ' 075 LD BC,(nn)
        setBC peekw(nxtpcw())
        execute_ed = 20
    End If
    Exit Function
    
ex_ed76_79:
    If (xxx And 2) Then GoTo ex_ed78_79 Else GoTo ex_ed76_77
    
ex_ed76_77:
    If xxx = 76 Then
        ' 076 NEG
        neg_a
        execute_ed = 8
    Else
        ' 077 RETI
        ' // TOCHECK: according to the official Z80 docs, IFF2 does not get
        ' //          copied to IFF1 for RETI - but in a real Z80 it is
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    End If
    Exit Function
    
ex_ed78_79:
    If xxx = 78 Then
        ' 078 IM 0
        intIM = 0
        execute_ed = 8
    Else
        ' 079 LD R,A
        intR = (regA And 128)
        intRTemp = intR
        execute_ed = 9
    End If
    Exit Function
    
ex_ed80_95:
    If (xxx And 8) Then GoTo ex_ed88_95 Else GoTo ex_ed80_87
    
ex_ed80_87:
    If (xxx And 4) Then GoTo ex_ed84_87 Else GoTo ex_ed80_83
    
ex_ed80_83:
    If (xxx And 2) Then GoTo ex_ed82_83 Else GoTo ex_ed80_81
    
ex_ed80_81:
    If xxx = 80 Then
        ' 080 IN D,(c)
        setD inn(regC) 'in_bc()
        execute_ed = 12
    Else
        ' 081 OUT (c),D
        out regC, getD
        execute_ed = 12
    End If
    Exit Function

ex_ed82_83:
    If xxx = 82 Then
        ' 082 SBC HL,DE
        regHL = sbc16(regHL, regDE)
        execute_ed = 15
    Else
        ' 083 LD (nn),DE
        pokew nxtpcw(), regDE
        execute_ed = 20
    End If
    Exit Function

ex_ed84_87:
    If (xxx And 2) Then GoTo ex_ed86_87 Else GoTo ex_ed84_85
    
ex_ed84_85:
    If xxx = 84 Then
        ' NEG
        neg_a
        execute_ed = 8
    Else
        ' 85 RETn
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    End If
    Exit Function
    
ex_ed86_87:
    If xxx = 86 Then
        ' 86 ' IM 1
        intIM = 1
        execute_ed = 8
    Else
        ' 87 ' LD A,I
        ld_a_i
        execute_ed = 9
    End If
    Exit Function
    
ex_ed88_95:
    If (xxx And 4) Then GoTo ex_ed92_95 Else GoTo ex_ed88_91
    
ex_ed88_91:
    If (xxx And 2) Then GoTo ex_ed90_91 Else GoTo ex_ed88_89
    
ex_ed88_89:
    If xxx = 88 Then
        ' 088 IN E,(c)
        setE inn(regC) 'in_bc()
        execute_ed = 12
    Else
        ' 089 OUT (c),E
        out regC, getE
        execute_ed = 12
    End If
    Exit Function
    
ex_ed90_91:
    If xxx = 90 Then
        ' 090 ADC HL,DE
        regHL = adc16(regHL, regDE)
        execute_ed = 15
    Else
        ' 091 LD DE,(nn)
        regDE = peekw(nxtpcw())
        execute_ed = 20
    End If
    Exit Function
    
ex_ed92_95:
    If (xxx And 2) Then GoTo ex_ed94_95 Else GoTo ex_ed92_93
    
ex_ed92_93:
    If xxx = 92 Then
        ' NEG
        neg_a
        execute_ed = 8
    Else
        ' 93 RETI
        ' // TOCHECK: according to the official Z80 docs, IFF2 does not get
        ' //          copied to IFF1 for RETI - but in a real Z80 it is
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    End If
    Exit Function
    
ex_ed94_95:
    If xxx = 94 Then
        ' IM 2
        intIM = 2
        execute_ed = 8
    Else
        ' 95 LD A,R
        ld_a_r
        execute_ed = 9
    End If
    Exit Function
    
ex_ed96_127:
    If (xxx And 16) Then GoTo ex_ed112_127 Else GoTo ex_ed96_111
    
ex_ed96_111:
    If (xxx And 8) Then GoTo ex_ed104_111 Else GoTo ex_ed96_103
    
ex_ed96_103:
    Select Case xxx
    Case 96 ' IN H,(c)
        setH inn(regC) 'in_bc()
        execute_ed = 12
    Case 97 ' OUT (c),H
        out regC, getH
        execute_ed = 12
    Case 98 ' SBC HL,HL
        regHL = sbc16(regHL, regHL)
        execute_ed = 15
    Case 99 ' LD (nn),HL
        pokew nxtpcw(), regHL
        execute_ed = 20
    Case 100 ' NEG
        neg_a
        execute_ed = 8
    Case 101 ' RETn
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    Case 102 ' IM 0
        intIM = 0
        execute_ed = 8
    Case 103 ' RRD
        rrd_a
        execute_ed = 18
    End Select
    Exit Function
    
ex_ed104_111:
    Select Case xxx
    Case 104 ' IN L,(c)
        setL inn(regC) 'in_bc()
        execute_ed = 12
    Case 105 ' OUT (c),L
        out regC, getL
        execute_ed = 12
    Case 106 ' ADC HL,HL
        regHL = adc16(regHL, regHL)
        execute_ed = 15
    Case 107 ' LD HL,(nn)
        regHL = peekw(nxtpcw())
        execute_ed = 20
    Case 108 ' NEG
        neg_a
        execute_ed = 8
    Case 109 ' RETI
        ' // TOCHECK: according to the official Z80 docs, IFF2 does not get
        ' //          copied to IFF1 for RETI - but in a real Z80 it is
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    Case 110 ' IM 0
        intIM = 0
        execute_ed = 8
    Case 111  ' RLD
        rld_a
        execute_ed = 18
    End Select
    Exit Function
    
ex_ed112_127:
    If (xxx And 8) Then GoTo ex_ed120_127 Else GoTo ex_ed112_119
    
ex_ed112_119:
    Select Case xxx
    Case 112 ' IN (c)
        'RAWR
        inn regC
        Debug.Print "IN C???"
        'in_bc
        execute_ed = 12
    Case 113 ' OUT (c),0
        out regC, 0
        execute_ed = 12
    Case 114 ' SBC HL,SP
        regHL = sbc16(regHL, regSP)
        execute_ed = 15
    Case 115 ' LD (nn),SP
        pokew nxtpcw(), regSP
        execute_ed = 20
    Case 116 ' NEG
        neg_a
        execute_ed = 8
    Case 117 ' RETn
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    Case 118 ' IM 1
        intIM = 1
        execute_ed = 8
    Case 119
        MsgBox "Unknown opcode 0xED 119"
    End Select
    Exit Function
    
ex_ed120_127:
    Select Case xxx
    Case 120 ' IN A,(c)
        regA = inn(regC) 'in_bc
        execute_ed = 12
    Case 121 ' OUT (c),A
        out regC, regA
        execute_ed = 12
    Case 122 ' ADC HL,SP
        regHL = adc16(regHL, regSP)
        execute_ed = 15
    Case 123 ' LD SP,(nn)
        regSP = peekw(nxtpcw())
        execute_ed = 20
    Case 124 ' NEG
        neg_a
        execute_ed = 8
    Case 125 ' RETI
        ' // TOCHECK: according to the official Z80 docs, IFF2 does not get
        ' //          copied to IFF1 for RETI - but in a real Z80 it is
        intIFF1 = intIFF2
        poppc
        execute_ed = 14
    Case 126 ' IM 2
        intIM = 2
        execute_ed = 8
    Case 127 ' NOP
        execute_ed = 8
    End Select
    Exit Function
    
ex_ed128_255:
    If (xxx And 64) Then GoTo ex_ed192_255 Else GoTo ex_ed128_191
    
ex_ed128_191:
    If (xxx And 32) Then
        GoTo ex_ed160_191
    Else
        ' NOP
        execute_ed = 8
        Exit Function
    End If

ex_ed160_191:
    Select Case xxx
    ' // xxI
    Case 160 ' LDI
        pokeb regDE, peekb(regHL)
        
        f3 = (F_3 And (peekb(regHL) + regA)) ' // TOCHECK: Is this correct?
        f5 = (2 And (peekb(regHL) + regA))   ' // TOCHECK: Is this correct?
        
        regDE = inc16(regDE)
        regHL = inc16(regHL)
        setBC dec16(getBC)

        fPV = (getBC <> 0)
        fH = False
        fN = False

        execute_ed = 16
    Case 161 ' CPI
        c = fC

        cp_a peekb(regHL)
        regHL = inc16(regHL)
        setBC dec16(getBC)

        fPV = (getBC <> 0)
        fC = c

        execute_ed = 16
    Case 162 ' INI
        pokeb regHL, inn(regC) 'inn(getBC)
        b = qdec8(regB)
        regB = b
        regHL = inc16(regHL)

        fZ = (b = 0)
        fN = True

        execute_ed = 16
    Case 163 ' OUTI
        b = qdec8(regB)
        regB = b
        out regC, peekb(regHL)
        regHL = inc16(regHL)

        fZ = (b = 0)
        fN = True

        execute_ed = 16
    
    ' /* xxD */
    Case 168 ' LDD
        pokeb regDE, peekb(regHL)
        
        f3 = (F_3 And (peekb(regHL) + regA)) ' // TOCHECK: Is this correct?
        f5 = (2 And (peekb(regHL) + regA))   ' // TOCHECK: Is this correct?
        
        regDE = dec16(regDE)
        regHL = dec16(regHL)
        setBC dec16(getBC())

        fPV = (getBC() <> 0)
        fH = False
        fN = False
        

        execute_ed = 16
    Case 169 ' CPD
        c = fC

        cp_a peekb(regHL)
        regHL = dec16(regHL)
        setBC dec16(getBC)

        fPV = (getBC <> 0)
        fC = c

        execute_ed = 16
    Case 170 ' IND
        pokeb regHL, inn(regC) 'inn(getBC)
        b = qdec8(regB)
        regB = b
        regHL = dec16(regHL)

        fZ = (b = 0)
        fN = True

        execute_ed = 16
    Case 171 ' OUTD
        count = qdec8(regB)
        regB = count
        out regC, peekb(regHL)
        regHL = dec16(regHL)

        fZ = (count = 0)
        fN = True

        execute_ed = 16

    ' // xxIR
    Case 176 ' LDIR
    
    'THINK THERES AN ERROR HERE...
        'Load location (DE) with location (HL), incr DE,HL; decr
        TempLocal_tstates = 0
        count = getBC
        dest = regDE
        from = regHL
        
        ' // REFRESH -2
        intRTemp = intRTemp - 2
        Do
            pokeb dest, peekb(from)
            from = (from + 1) And 65535
            dest = (dest + 1) And 65535
            count = count - 1
            
            TempLocal_tstates = TempLocal_tstates + 21
            ' // REFRESH (2)
            intRTemp = intRTemp + 2
            If (TempLocal_tstates >= 0) Then
                ' // interruptTriggered
                Exit Do
            End If
        Loop While count <> 0
        
        regPC = regPC - 2
        fH = False
        fN = False
        fPV = True
        f3 = (F_3 And (peekb(from - 1) + regA)) ' // TOCHECK: Is this correct?
        f5 = (2 And (peekb(from - 1) + regA))   ' // TOCHECK: Is this correct?
        
        If count = 0 Then
            regPC = regPC + 2
            TempLocal_tstates = TempLocal_tstates - 5
            fPV = False
        End If
        regDE = dest
        regHL = from
        setBC count
            
        execute_ed = TempLocal_tstates
    Case 177 ' CPIR
        c = fC
        
        cp_a peekb(regHL)
        regHL = inc16(regHL)
        setBC dec16(getBC)
        
        fC = c
        c = getBC <> 0
        fPV = c
        If (fPV) And (fZ = False) Then
            regPC = regPC - 2
            execute_ed = 21
        Else
            execute_ed = 16
        End If
    Case 178 ' INIR
        pokeb regHL, inn(regC) 'inn(getBC)
        b = qdec8(regB)
        regB = b
        regHL = inc16(regHL)

        fZ = True
        fN = True
        If (b <> 0) Then
            regPC = regPC - 2
            execute_ed = 21
        Else
            execute_ed = 16
        End If
    Case 179 ' OTIR
        b = qdec8(regB)
        regB = b
        out regC, peekb(regHL)
        regHL = inc16(regHL)

        fZ = True
        fN = True
        If (b <> 0) Then
            regPC = regPC - 2
            execute_ed = 21
        Else
            execute_ed = 16
        End If

    ' // xxDR
    Case 184 ' LDDR
        TempLocal_tstates = 0
        count = getBC
        dest = regDE
        from = regHL
        
        ' // REFRESH -2
        intRTemp = intRTemp - 2
        Do
            pokeb dest, peekb(from)
            from = (from - 1) And 65535
            dest = (dest - 1) And 65535
            count = count - 1
            
            TempLocal_tstates = TempLocal_tstates + 21
            
            ' // REFRESH (2)
            intRTemp = intRTemp + 2
            
            If (TempLocal_tstates >= 0) Then
                ' // interruptTriggered
                Exit Do
            End If
        Loop While count <> 0
        regPC = regPC - 2
        fH = False
        fN = False
        fPV = True

        f3 = (F_3 And (peekb(from - 1) + regA)) ' // TOCHECK: Is this correct?
        f5 = (2 And (peekb(from - 1) + regA))   ' // TOCHECK: Is this correct?

        If count = 0 Then
            regPC = regPC + 2
            TempLocal_tstates = TempLocal_tstates - 5
            fPV = False
        End If
        
        regDE = dest
        regHL = from
        setBC count
            
        execute_ed = TempLocal_tstates
    Case 185 ' CPDR
        c = fC

        cp_a peekb(regHL)
        regHL = dec16(regHL)
        setBC dec16(getBC)

        fPV = getBC <> 0
        fC = c
        If (fPV) And (fZ = False) Then
            regPC = regPC - 2
            execute_ed = 21
        Else
            execute_ed = 16
        End If
    Case 186 ' INDR
        pokeb regHL, inn(regC) 'inn(getBC)
        b = qdec8(regB)
        regB = b
        regHL = dec16(regHL)

        fZ = True
        fN = True
        If (b <> 0) Then
            regPC = regPC - 2
            execute_ed = 21
        Else
            execute_ed = 16
        End If
    Case 187 ' OTDR
        b = qdec8(regB)
        regB = b
        out regC, peekb(regHL)
        regHL = dec16(regHL)

        fZ = True
        fN = True
        If (b <> 0) Then
            regPC = regPC - 2
            execute_ed = 21
        Else
            execute_ed = 16
        End If
    Case 187 To 191
        MsgBox "Unknown ED instruction " & xxx & " at " & regPC
    Case Else ' (164 To 167, 172 To 175, 180 To 183)
        ' NOP
        execute_ed = 8
        
    End Select
    Exit Function
    
ex_ed192_255:
    Debug.Print "Unknown ED instruction " & xxx
    'Select Case xxx
    'Case 252
        ' // Patched tape LOAD routine
    '    TapeLoad regHL
    'Case 253
        ' // Patched tape SAVE routine
    '    TapeSave regHL
    'Case Else
    '    MsgBox "Unknown ED instruction " & xxx & " at " & regPC
    '    execute_ed = 8
    'End Select
End Function
Private Sub rld_a()
    Dim ans As Long, t As Long, q As Long
    
    ans = regA
    t = peekb(regHL)
    q = t

    t = (t * 16) Or (ans And &HF&)
    ans = (ans And &HF0&) Or q \ 16&
    pokeb regHL, (t And &HFF&)

    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = intIFF2
    fH = False
    fN = False
    
    regA = ans
End Sub
Private Sub rrd_a()
    Dim ans As Long, t As Long, q As Long
    
    ans = regA
    t = peekb(regHL)
    q = t

    t = (t \ 16&) Or (ans * 16&)
    ans = (ans And &HF0&) Or (q And &HF&)
    pokeb regHL, t

    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = intIFF2
    fH = False
    fN = False
    
    regA = ans
End Sub
Private Sub neg_a()
    Dim t As Long
    
    t = regA
    regA = 0
    sub_a t
End Sub
Private Sub execute_id_cb(op As Long, ByVal Z As Long)
    Select Case op
    Case 0 ' RLC B
        op = rlc(peekb(Z))
        regB = op
        pokeb Z, op
    Case 1 ' RLC C
        op = rlc(peekb(Z))
        regC = op
        pokeb Z, op
    Case 2 ' RLC D
        op = rlc(peekb(Z))
        setD op
        pokeb Z, op
    Case 3 ' RLC E
        op = rlc(peekb(Z))
        setE op
        pokeb Z, op
    Case 4 ' RLC H
        op = rlc(peekb(Z))
        setH op
        pokeb Z, op
    Case 5 ' RLC L
        op = rlc(peekb(Z))
        setL op
        pokeb Z, op
    Case 6 ' RLC (HL)
        pokeb Z, rlc(peekb(Z))
    Case 7 ' RLC A
        op = rlc(peekb(Z))
        regA = op
        pokeb Z, op
    
    Case 8 ' RRC B
        op = rrc(peekb(Z))
        regB = op
        pokeb Z, op
    Case 9 ' RRC C
        op = rrc(peekb(Z))
        regC = op
        pokeb Z, op
    Case 10 ' RRC D
        op = rrc(peekb(Z))
        setD op
        pokeb Z, op
    Case 11 ' RRC E
        op = rrc(peekb(Z))
        setE op
        pokeb Z, op
    Case 12 ' RRC H
        op = rrc(peekb(Z))
        setH op
        pokeb Z, op
    Case 13 ' RRC L
        op = rrc(peekb(Z))
        setL op
        pokeb Z, op
    Case 14 ' RRC (HL)
        pokeb Z, rrc(peekb(Z))
    Case 15 ' RRC A
        op = rrc(peekb(Z))
        regA = op
        pokeb Z, op
    Case 16 ' RL B
        op = rl(peekb(Z))
        regB = op
        pokeb Z, op
    Case 17 ' RL C
        op = rl(peekb(Z))
        regC = op
        pokeb Z, op
    Case 18 ' RL D
        op = rl(peekb(Z))
        setD op
        pokeb Z, op
    Case 19 ' RL E
        op = rl(peekb(Z))
        setE op
        pokeb Z, op
    Case 20 ' RL H
        op = rl(peekb(Z))
        setH op
        pokeb Z, op
    Case 21 ' RL L
        op = rl(peekb(Z))
        setL op
        pokeb Z, op
    Case 22 ' RL (HL)
        pokeb Z, rl(peekb(Z))
    Case 23 ' RL A
        op = rl(peekb(Z))
        regA = op
        pokeb Z, op
    Case 24 ' RR B
        op = rr(peekb(Z))
        regB = op
        pokeb Z, op
    Case 25 ' RR C
        op = rr(peekb(Z))
        regC = op
        pokeb Z, op
    Case 26 ' RR D
        op = rr(peekb(Z))
        setD op
        pokeb Z, op
    Case 27 ' RR E
        op = rr(peekb(Z))
        setE op
        pokeb Z, op
    Case 28 ' RR H
        op = rr(peekb(Z))
        setH op
        pokeb Z, op
    Case 29 ' RR L
        op = rr(peekb(Z))
        setL op
        pokeb Z, op
    Case 30 ' RR (HL)
        pokeb Z, rl(peekb(Z))
    Case 31 ' RR A
        op = rr(peekb(Z))
        regA = op
        pokeb Z, op
    Case 32 ' SLA B
        op = sla(peekb(Z))
        regB = op
        pokeb Z, op
    Case 33 ' SLA C
        op = sla(peekb(Z))
        regC = op
        pokeb Z, op
    Case 34 ' SLA D
        op = sla(peekb(Z))
        setD op
        pokeb Z, op
    Case 35 ' SLA E
        op = sla(peekb(Z))
        setE op
        pokeb Z, op
    Case 36 ' SLA H
        op = sla(peekb(Z))
        setH op
        pokeb Z, op
    Case 37 ' SLA L
        op = sla(peekb(Z))
        setL op
        pokeb Z, op
    Case 38 ' SLA (HL)
        pokeb Z, sla(peekb(Z))
    Case 39 ' SLA A
        op = sla(peekb(Z))
        regA = op
        pokeb Z, op
    Case 40 ' SRA B
        op = sra(peekb(Z))
        regB = op
        pokeb Z, op
    Case 41 ' SRA C
        op = sra(peekb(Z))
        regC = op
        pokeb Z, op
    Case 42 ' SRA D
        op = sra(peekb(Z))
        setD op
        pokeb Z, op
    Case 43 ' SRA E
        op = sra(peekb(Z))
        setE op
        pokeb Z, op
    Case 44 ' SRA H
        op = sra(peekb(Z))
        setH op
        pokeb Z, op
    Case 45 ' SRA L
        op = sra(peekb(Z))
        setL op
        pokeb Z, op
    Case 46 ' SRA (HL)
        pokeb Z, sra(peekb(Z))
    Case 47 ' SRA A
        op = sra(peekb(Z))
        regA = op
        pokeb Z, op
    Case 48 ' SLS B
        op = sls(peekb(Z))
        regB = op
        pokeb Z, op
    Case 49 ' SLS C
        op = sls(peekb(Z))
        regC = op
        pokeb Z, op
    Case 50 ' SLS D
        op = sls(peekb(Z))
        setD op
        pokeb Z, op
    Case 51 ' SLS E
        op = sls(peekb(Z))
        setE op
        pokeb Z, op
    Case 52 ' SLS H
        op = sls(peekb(Z))
        setH op
        pokeb Z, op
    Case 53 ' SLS L
        op = sls(peekb(Z))
        setL op
        pokeb Z, op
    Case 54 ' SLS (HL)
        pokeb Z, sls(peekb(Z))
    Case 55 ' SLS A
        op = sls(peekb(Z))
        regA = op
        pokeb Z, op
        
        
    Case 62 ' SRL (HL)
        pokeb Z, srl(peekb(Z))
    Case 63 ' SRL A
        op = srl(peekb(Z))
        regA = op
        pokeb Z, op
    Case 64 To 71 ' BIT 0,B
        bit &H1&, peekb(Z)
    Case 72 To 79 ' BIT 1,B
        bit &H2&, peekb(Z)
    Case 80 To 87 ' BIT 2,B
        bit &H4&, peekb(Z)
    Case 88 To 95 ' BIT 3,B
        bit &H8&, peekb(Z)
    
    Case 96 To 103 ' BIT 4,B
        bit &H10&, peekb(Z)
    Case 104 To 111 ' BIT 5,B
        bit &H20&, peekb(Z)
    Case 112 To 119 ' BIT 6,B
        bit &H40&, peekb(Z)
    Case 120 To 127 ' BIT 7,B
        bit &H80&, peekb(Z)
    Case 134 ' RES 0,(HL)
        pokeb Z, bitRes(&H1&, peekb(Z))
    Case 142 ' RES 1,(HL)
        pokeb Z, bitRes(&H2&, peekb(Z))
    Case 150 ' RES 2,(HL)
        pokeb Z, bitRes(&H4&, peekb(Z))
    Case 158 ' RES 3,(HL)
        pokeb Z, bitRes(&H8&, peekb(Z))
    Case 166 ' RES 4,(HL)
        pokeb Z, bitRes(&H10&, peekb(Z))
    Case 172 ' RES 5,H
        setH bitRes(&H20&, peekb(Z))
        pokeb Z, getH
    Case 174 ' RES 5,(HL)
        pokeb Z, bitRes(&H20&, peekb(Z))
    Case 175 ' RES 5,A
        regA = bitRes(&H20&, peekb(Z))
        pokeb Z, regA
    Case 182 ' RES 6,(HL)
        pokeb Z, bitRes(&H40&, peekb(Z))
    Case 190 ' RES 7,(HL)
        pokeb Z, bitRes(&H80&, peekb(Z))
    Case 198 ' SET 0,(HL)
        pokeb Z, bitSet(&H1&, peekb(Z))
    Case 206 ' SET 1,(HL)
        pokeb Z, bitSet(&H2&, peekb(Z))
    Case 214 ' SET 2,(HL)
        pokeb Z, bitSet(&H4&, peekb(Z))
    Case 222 ' SET 3,(HL)
        pokeb Z, bitSet(&H8&, peekb(Z))
    Case 230 ' SET 4,(HL)
        pokeb Z, bitSet(&H10&, peekb(Z))
    Case 238 ' SET 5,(HL)
        pokeb Z, bitSet(&H20&, peekb(Z))
    Case 246 ' SET 6,(HL)
        pokeb Z, bitSet(&H40&, peekb(Z))
    Case 254 ' SET 7,(HL)
        pokeb Z, bitSet(&H80&, peekb(Z))
    Case 255 ' SET 7,A
        regA = bitSet(&H80&, peekb(Z))
        pokeb Z, regA
    Case Else
        MsgBox "Invalid ID CB op=" & op & " z=" & Z
    End Select
End Sub

Private Sub exx()
    Dim t As Long
    
    t = regHL
    regHL = regHL_
    regHL_ = t
    
    t = regDE
    regDE = regDE_
    regDE_ = t
    
    t = getBC
    setBC regBC_
    regBC_ = t
End Sub
Private Function getAF() As Long
    getAF = (regA * 256&) Or getF
End Function
Private Function getBC() As Long
    getBC = (regB * 256&) Or regC
End Function
Private Function getD() As Long
    getD = glMemAddrDiv256(regDE)
End Function
Private Function getE() As Long
    getE = regDE And &HFF&
End Function
Private Function getF() As Long
    If fS Then getF = getF Or F_S
    If fZ Then getF = getF Or F_Z
    If f5 Then getF = getF Or F_5
    If fH Then getF = getF Or F_H
    If f3 Then getF = getF Or F_3
    If fPV Then getF = getF Or F_PV
    If fN Then getF = getF Or F_N
    If fC Then getF = getF Or F_C
End Function


Private Function getH() As Long
    getH = glMemAddrDiv256(regHL)
End Function
Private Function getL() As Long
    getL = regHL And &HFF&
End Function

Private Function getR() As Long
    getR = intR
End Function
Private Function id_d() As Long
    Dim d As Long
    
    d = nxtpcb()
    If d And 128 Then d = -(256 - d)
    id_d = (regID + d) And &HFFFF&
End Function
Private Sub ld_a_i()
    fS = (intI And F_S) <> 0
    f3 = (intI And F_3) <> 0
    f5 = (intI And F_5) <> 0
    fZ = (intI = 0)
    fPV = intIFF2
    fH = False
    fN = False

    regA = intI
End Sub
Private Sub ld_a_r()
'    If ((intRTemp \ 128) And 1) = 1 Then
'        intR = 0
'    Else
'        intR = 128
'    End If

    intRTemp = intRTemp And &H7F&
    regA = (intR And &H80&) Or intRTemp
    fS = (regA And F_S) <> 0
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fZ = (regA = 0)
    fPV = intIFF2
    fH = False
    fN = False
End Sub
Private Function inc8(ByVal ans As Long) As Long
    fPV = (ans = &H7F&)
    fH = (((ans And &HF&) + 1) And F_H) <> 0
    
    ans = (ans + 1) And &HFF&

    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fN = False
    
    inc8 = ans
End Function
Private Function dec8(ByVal ans As Long) As Long
    fPV = (ans = &H80&)
    fH = (((ans And &HF&) - 1) And F_H) <> 0
    
    ans = (ans - 1) And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    
    fN = True
    
    dec8 = ans
End Function

Private Function interruptTriggered(Tstates As Long) As Long
    interruptTriggered = (Tstates >= 0)
End Function
Private Sub or_a(b As Long)
    regA = (regA Or b)
    
    fS = (regA And F_S) <> 0
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fH = False
    fPV = Parity(regA)
    fZ = (regA = 0)
    fN = False
    fC = False
End Sub
Public Function peekw(Addr As Long) As Long
    peekw = peekb(Addr) Or (peekb(Addr + 1) * 256&)
End Function

Sub pokeb(ByVal address As Long, value As Long)

        Dim temp As Long

        If (address < &H8000&) Then
                                        Exit Sub
    ElseIf (address < &HC000&) Then     '// ROM (And Cart RAM).
                                        cartRom(address) = value
    ElseIf (address < &HE000&) Then     '// RAM (Onboard).
                                        cartRam(address - &HC000&) = value
    ElseIf (address < &HFFFC&) Then     '// RAM (Mirrored)
                                        cartRam(address - &HE000&) = value
    ElseIf (address < &H10000) Then
    
    '-----------------------------+
    '       Paging Registers      |
    '-----------------------------+
    
    '-----------------------------+
    '   SRAM/ROM Select Register. |
    '-----------------------------+
    If (address = &HFFFC&) Then
    
        If ((value And 8) = 8) Then     '// Uses SRAM.
        frame_two_rom = False
                
            If ((value And 4) = 4) Then '// Uses SRAM Page 0.
                '[Enter SRAM Page 0 Code]
            Else
                CopyMemory cartRom(Page2), SRAM(vbEmpty), Page1
            End If
            
        Else
            If frame_two_rom = False Then CopyMemory SRAM(vbEmpty), cartRom(Page2), Page1
            frame_two_rom = True
        End If
    
    '-----------------------------+
    '       Page 0 Rom Bank       |
    '-----------------------------+
    ElseIf (address = &HFFFD&) Then
    
        temp = Mul4000(value Mod number_of_pages) + &H400&
        CopyMemory cartRom(&H400&), pages(temp), &H3C00&
        
    '-----------------------------+
    '      Page 1 Rom Bank.       |
    '-----------------------------+
    ElseIf (address = &HFFFE&) Then
        
        temp = Mul4000(value Mod number_of_pages)
        CopyMemory cartRom(Page1), pages(temp), Page1
    
    '-----------------------------+
    '      Page 2 Rom Bank.       |
    '-----------------------------+
    ElseIf (address = &HFFFF&) Then
    
        If (frame_two_rom) Then
            temp = Mul4000(value Mod number_of_pages)
            CopyMemory cartRom(Page2), pages(temp), Page1
        End If
    End If

    cartRam(address - &HE000&) = value
    End If

End Sub

Sub pokew(ByVal Addr As Long, word As Long)
    pokeb Addr, word And &HFF&
       
    pokeb (Addr + 1) And &HFFFF&, glMemAddrDiv256(word And &HFF00&)

End Sub

Sub poppc()
    regPC = popw
End Sub
Private Function popw() As Long
    popw = peekb(regSP) Or (peekb(regSP + 1) * 256&)
    regSP = (regSP + 2 And &HFFFF&)
End Function
Sub pushpc()
    pushw regPC
End Sub

Private Sub pushw(word As Long)
    regSP = (regSP - 2) And &HFFFF&
    pokew regSP, word
End Sub
Private Function rl(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H80&) <> 0
    
    If fC Then
        ans = (ans * 2) Or &H1&
    Else
        ans = ans * 2
    End If
    ans = ans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
        
    rl = ans
End Function
Private Sub rl_a()
    Dim ans As Long, c As Long
    
    ans = regA
    c = (ans And &H80&) <> 0
    
    If fC Then
        ans = (ans * 2) Or &H1&
    Else
        ans = (ans * 2)
    End If
    ans = ans And &HFF&
    
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fN = False
    fH = False
    fC = c
    
    regA = ans
End Sub
Private Function rlc(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H80&) <> 0
    
    If c Then
        ans = (ans * 2) Or &H1&
    Else
        ans = (ans * 2)
    End If
    
    ans = ans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
    
    rlc = ans
End Function
Private Sub rlc_a()
    Dim c As Long
    
    c = (regA And &H80&) <> 0
    
    If c Then
        regA = (regA * 2) Or 1
    Else
        regA = (regA * 2)
    End If
    regA = regA And &HFF&
    
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fN = False
    fH = False
    fC = c
End Sub

Private Function rr(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H1&) <> 0
    
    If fC Then
        ans = ShiftRight(ans, 1) Or &H80&
    Else
        ans = ShiftRight(ans, 1) 'glMemAddrDiv2(ans)
    End If

    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
        
    rr = ans
End Function
Private Sub rr_a()
    Dim ans As Long, c As Long
    
    ans = regA
    c = (ans And &H1&) <> 0
    
    If fC Then
        ans = ShiftRight(ans, 1&) Or &H80& 'glMemAddrDiv2(ans) Or &H80&
    Else
        ans = ShiftRight(ans, 1&) 'glMemAddrDiv2(ans)
    End If
        
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fN = False
    fH = False
    fC = c
    
    regA = ans
End Sub

Private Function rrc(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H1&) <> 0
    
    If c Then
        ans = ShiftRight(ans, 1&) Or &H80& 'glMemAddrDiv2(ans) Or &H80&
    Else
        ans = ShiftRight(ans, 1&) 'glMemAddrDiv2(ans)
    End If
      
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
    
    rrc = ans
End Function

Private Sub rrc_a()
    Dim c As Long
    
    c = (regA And 1) <> 0
    
    If c Then
        regA = ShiftRight(regA, 1&) Or &H80& 'glMemAddrDiv2(regA) Or &H80&
    Else
        regA = ShiftRight(regA, 1&) 'glMemAddrDiv2(regA)
    End If
        
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fN = False
    fH = False
    fC = c
End Sub
Private Sub sbc_a(ByVal b As Long)
    Dim a As Long, wans As Long, ans As Long, c As Long
    
    a = regA
    
    If fC Then c = 1

    wans = a - b - c
    ans = wans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fC = (wans And &H100&) <> 0
    fPV = ((a Xor b) And (a Xor ans) And &H80&) <> 0
    fH = (((a And &HF&) - (b And &HF&) - c) And F_H) <> 0
    fN = True
    
    regA = ans
End Sub
Private Function sbc16(a As Long, b As Long) As Long
    Dim c As Long, lans As Long, ans As Long
    
    If fC Then c = 1
    
    lans = a - b - c
    ans = lans And &HFFFF&
    
    fS = (ans And (F_S * 256&)) <> 0
    f3 = (ans And (F_3 * 256&)) <> 0
    f5 = (ans And (F_5 * 256&)) <> 0
    fZ = (ans = 0)
    fC = (lans And &H10000) <> 0
    fPV = ((a Xor b) And (a Xor ans) And &H8000&) <> 0
    fH = (((a And &HFFF&) - (b And &HFFF&) - c) And &H1000&) <> 0
    fN = True
    
    sbc16 = ans
End Function

Private Sub scf()
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fN = False
    fH = False
    fC = True
End Sub
Public Sub setAF(v As Long)
    regA = glMemAddrDiv256(v And &HFF00&)
    setF (v And &HFF&)
End Sub
Sub setBC(nn As Long)
    regB = glMemAddrDiv256(nn And &HFF00&)
    regC = nn And &HFF&
End Sub

Public Sub execute()

    Dim Tstates As Long, lineno     As Long: Tstates = Tstates - Tcycles
    Dim opcode  As Long, oldTick    As Long: oldTick = timeGetTime
    Dim ltemp   As Long, d          As Long
    
fetch_next:

    If regPC = &H522 Then
        MsgBox lineno
        Beep
    End If
    
    '----------------------------------------------+
    '        <Start> Z80 Interrupt Routine         |
    '----------------------------------------------+
    If irqsetLine Then
        If halt Then regPC = (regPC + 1): halt = False
        If intIFF1 Then
        
            intIFF1 = False
            intIFF2 = False
                
            If (intIM < 2&) Then
                
                regSP = (regSP - 2) And &HFFFF&
                
                pokeb regSP, regPC And &HFF&
                pokeb (regSP + 1) And &HFFFF&, glMemAddrDiv256(regPC And &HFF00&)

                regPC = &H38&
                Tstates = Tstates + 13

            End If
        End If
    End If
    '----------------------------------------------+
    '         <End> Z80 Interrupt Routine          |
    '----------------------------------------------+


    '----------------------------------------------+
    '         <Start> VDP Line Interrupts          |
    '----------------------------------------------+
    If (Tstates >= 0) Then
        Tstates = Tstates - Tcycles                     '// Reset Interrupt Counter.
        Vdp.interrupts (lineno)                         '// Handle VDP Interrupt.
        If (lineno < 192&) Then Vdp.drawLine (lineno)   '// Draw Next VDP Scanline.
                
        If (lineno < 312&) Then
            lineno = (lineno + 1&)
        Else
            '--------------------------------------+
            '        CPU Clock Throttle (FPS).     |
            '--------------------------------------+
            ltemp = (timeGetTime - oldTick)
            If (ltemp < 19&) Then Sleep (19& - ltemp)
            oldTick = timeGetTime
            '-------------------------------------+
            '        1000 \ (50) = 20 (PAL).      |
            '-------------------------------------+
            lineno = vbEmpty
            If GetQueueStatus(QS_MOUSEBUTTON Or QS_KEY Or QS_SENDMESSAGE) Then DoEvents
            StretchDIBits Form1.hdc, 0&, 0&, SMS_WIDTH, SMS_HEIGHT, 0&, 0&, SMS_WIDTH, SMS_HEIGHT, display(0&), myBMP, 0&, vbSrcCopy
        End If
    End If
    '----------------------------------------------+
    '          <End> VDP Line Interrupts           |
    '----------------------------------------------+

        opcode = nxtpcb()

        If (opcode And 128) Then GoTo ex128_255 Else GoTo ex0_127
ex0_127:
        If (opcode And 64) Then GoTo ex64_127 Else GoTo ex0_63
ex0_63:
        If (opcode And 32) Then GoTo ex32_63 Else GoTo ex0_31
ex0_31:
        If (opcode And 16) Then GoTo ex16_31 Else GoTo ex0_15
ex0_15:
        If (opcode And 8) Then GoTo ex8_15 Else GoTo ex0_7

ex0_7:
        If (opcode And 4) Then GoTo ex4_7 Else GoTo ex0_3

ex0_3:
        If (opcode And 2) Then GoTo ex2_3 Else GoTo ex0_1

ex0_1:
        If opcode = 0 Then
            ' 000 NOP
            Tstates = Tstates + 4
        Else
            ' 001 LD BC,nn
            setBC nxtpcw()
            Tstates = Tstates + 10
        End If
        GoTo fetch_next

ex2_3:
        If opcode = 2 Then
            ' 002 LD (BC),A
            pokeb getBC, regA
            Tstates = Tstates + 7
        Else
            ' 003 INC BC
            setBC inc16(getBC)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex4_7:
    If (opcode And 2) Then GoTo ex6_7 Else GoTo ex4_5

ex4_5:
        If opcode = 4 Then
            ' 004 INC B
            regB = inc8(regB)
            Tstates = Tstates + 4
        Else
            ' 005 DEC B
            regB = dec8(regB)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex6_7:
        If opcode = 6 Then
            ' 006 LD B,n
            regB = nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 007 RLCA
            rlc_a
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex8_15:
        If (opcode And 4) Then GoTo ex12_15 Else GoTo ex8_11

ex8_11:
        If (opcode And 2) Then GoTo ex10_11 Else GoTo ex8_9

ex8_9:
        If opcode = 8 Then
            ' 008 EX AF,AF'
            ex_af_af
            Tstates = Tstates + 4
        Else
            '009 ADD HL,BC
            regHL = add16(regHL, getBC())
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex10_11:
        If opcode = 10 Then
            ' 010 LD A,(BC)
            regA = peekb(getBC)
            Tstates = Tstates + 7
        Else
            ' 011 DEC BC
            setBC dec16(getBC)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex12_15:
        If (opcode And 2) Then GoTo ex14_15 Else GoTo ex12_13

ex12_13:
        If opcode = 12 Then
            ' 012 INC C
            regC = inc8(regC)
            Tstates = Tstates + 4
        Else
            ' 013 DEC C
            regC = dec8(regC)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex14_15:
        If opcode = 14 Then
            ' 014 LD C,n
            regC = nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 015 RRCA
            rrc_a
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex16_31:
        If (opcode And 8) Then GoTo ex24_31 Else GoTo ex16_23

ex16_23:
        If (opcode And 4) Then GoTo ex20_23 Else GoTo ex16_19

ex16_19:
        If (opcode And 2) Then GoTo ex18_19 Else GoTo ex16_17

ex16_17:
        If opcode = 16 Then
            ' 016 DJNZ dis
            ltemp = qdec8(regB)

            regB = ltemp
            If ltemp <> 0 Then
                d = nxtpcb()
                If d And 128 Then d = -(256 - d)
                regPC = (regPC + d) And &HFFFF&
                Tstates = Tstates + 13
            Else
                regPC = inc16(regPC)
                Tstates = Tstates + 8
            End If
        Else
            ' 017 LD DE,nn
            regDE = nxtpcw()
            Tstates = Tstates + 10
        End If
        GoTo fetch_next

ex18_19:
        If opcode = 18 Then
            ' 018 LD (DE),A
            pokeb regDE, regA
            Tstates = Tstates + 7
        Else
            ' 019 INC DE
            regDE = inc16(regDE)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex20_23:
        If (opcode And 2) Then GoTo ex22_23 Else GoTo ex20_21

ex20_21:
        If opcode = 20 Then
        ' 020 INC D
            setD inc8(getD)
            Tstates = Tstates + 4
        Else
        ' 021 DEC D
            setD dec8(getD)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex22_23:
        If opcode = 22 Then
            ' 022 LD D,n
            setD nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 023 ' RLA
            rl_a
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex24_31:
        If (opcode And 4) Then GoTo ex28_31 Else GoTo ex24_27

ex24_27:
        If (opcode And 2) Then GoTo ex26_27 Else GoTo ex24_25

ex24_25:
        If opcode = 24 Then
        ' 024 JR dis
            d = nxtpcb()
            If d And 128 Then d = -(256 - d)
            regPC = (regPC + d) And &HFFFF&
            Tstates = Tstates + 12
        Else
            ' 025 ADD HL,DE
            regHL = add16(regHL, regDE)
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex26_27:
        If opcode = 26 Then
            ' 026 LD A,(DE)
            regA = peekb(regDE)
            Tstates = Tstates + 7
        Else
            ' 027 DEC DE
            regDE = dec16(regDE)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex28_31:
        If (opcode And 2) Then GoTo ex30_31 Else GoTo ex28_29

ex28_29:
        If opcode = 28 Then
            ' 028 INC E
            setE inc8(getE)
            Tstates = Tstates + 4
        Else
            ' 029 DEC E
            setE dec8(getE)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex30_31:
        If opcode = 30 Then
            ' 030 LD E,n
            setE nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 031 RRA
            rr_a
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex32_63:
        If (opcode And 16) Then GoTo ex48_63 Else GoTo ex32_47

ex32_47:
        If (opcode And 8) Then GoTo ex40_47 Else GoTo ex32_39

ex32_39:
        If (opcode And 4) Then GoTo ex36_39 Else GoTo ex32_35

ex32_35:
        If (opcode And 2) Then GoTo ex34_35 Else GoTo ex32_33

ex32_33:
        If opcode = 32 Then
            ' 032 JR NZ dis
            If fZ = False Then
                d = nxtpcb()
                If d And 128 Then d = -(256 - d)
                regPC = ((regPC + d) And &HFFFF&)
                Tstates = Tstates + 12
            Else
                regPC = inc16(regPC)
                Tstates = Tstates + 7
            End If
        Else
            ' 033 LD HL,nn
            regHL = nxtpcw()
            Tstates = Tstates + 10
        End If
        GoTo fetch_next

ex34_35:
        If opcode = 34 Then
            ' 034 LD (nn),HL
            pokew nxtpcw(), regHL
            Tstates = Tstates + 16
        Else
            ' 035 INC HL
            regHL = inc16(regHL)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex36_39:
        If (opcode And 2) Then GoTo ex38_39 Else GoTo ex36_37

ex36_37:
        If opcode = 36 Then
            ' 036 INC H
            setH inc8(getH)
            Tstates = Tstates + 4
        Else
            ' 037 DEC H
            setH dec8(getH)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex38_39:
        If opcode = 38 Then
            ' 038 LD H,n
            setH nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 039 DAA
            daa_a
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex40_47:
        If (opcode And 4) Then GoTo ex44_47 Else GoTo ex40_43

ex40_43:
        If (opcode And 2) Then GoTo ex42_43 Else GoTo ex40_41

ex40_41:
        If opcode = 40 Then
            ' 040 JR Z dis
            If fZ = True Then
                d = nxtpcb()
                If d And 128 Then d = -(256 - d)
                regPC = ((regPC + d) And &HFFFF&)
                Tstates = Tstates + 12
            Else
                regPC = inc16(regPC)
                Tstates = Tstates + 7
            End If
        Else
            ' 041 ADD HL,HL
            regHL = add16(regHL, regHL)
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex42_43:
        If opcode = 42 Then
            ' 042 LD HL,(nn)
            regHL = peekw(nxtpcw())
            Tstates = Tstates + 16
        Else
            ' 043 DEC HL
            regHL = dec16(regHL)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex44_47:
        If (opcode And 2) Then GoTo ex46_47 Else GoTo ex44_45

ex44_45:
        If opcode = 44 Then
            ' 044 INC L
            setL inc8(getL)
            Tstates = Tstates + 4
        Else
            ' 045 DEC L
            setL dec8(getL)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex46_47:
        If opcode = 46 Then
            ' 046 LD L,n
            setL nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 047 CPL
            cpl_a
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex48_63:
        If (opcode And 8) Then GoTo ex56_63 Else GoTo ex48_55

ex48_55:
        If (opcode And 4) Then GoTo ex52_55 Else GoTo ex48_51

ex48_51:
        If (opcode And 2) Then GoTo ex50_51 Else GoTo ex48_49

ex48_49:
        If opcode = 48 Then
            ' 048 JR NC dis
            If fC = False Then
                d = nxtpcb()
                If d And 128 Then d = -(256 - d)
                regPC = ((regPC + d) And &HFFFF&)
                Tstates = Tstates + 12
            Else
                regPC = inc16(regPC)
                Tstates = Tstates + 7
            End If
        Else
            ' 049 LD SP,nn
            regSP = nxtpcw()
            Tstates = Tstates + 10
        End If
        GoTo fetch_next

ex50_51:
        If opcode = 50 Then
            ' 050 LD (nn),A
            pokeb nxtpcw, regA
            Tstates = Tstates + 13
        Else
            ' 051 INC SP
            regSP = inc16(regSP)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex52_55:
        If (opcode And 2) Then GoTo ex54_55 Else GoTo ex52_53

ex52_53:
        If opcode = 52 Then
            ' 052 INC (HL)
            pokeb regHL, inc8(peekb(regHL))
            Tstates = Tstates + 11
        Else
            ' 053 DEC (HL)
            pokeb regHL, dec8(peekb(regHL))
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex54_55:
        If opcode = 54 Then
            ' 054 LD (HL),n
            pokeb regHL, nxtpcb()
            Tstates = Tstates + 10
        Else
            ' 055 SCF
            scf
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex56_63:
        If (opcode And 4) Then GoTo ex60_63 Else GoTo ex56_59

ex56_59:
        If (opcode And 2) Then GoTo ex58_59 Else GoTo ex56_57

ex56_57:
        If opcode = 56 Then
            ' 056 JR C dis
            If fC = True Then
                d = nxtpcb()
                If d And 128 Then d = -(256 - d)
                regPC = ((regPC + d) And &HFFFF&)
                Tstates = Tstates + 12
            Else
                regPC = inc16(regPC)
                Tstates = Tstates + 7
            End If
        Else
            ' 057 ADD HL,SP
            regHL = add16(regHL, regSP)
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex58_59:
        If opcode = 58 Then
            ' 058 LD A,(nn)
            regA = peekb(nxtpcw())
            Tstates = Tstates + 13
        Else
            ' 059 DEC SP
            regSP = dec16(regSP)
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex60_63:
        If (opcode And 2) Then GoTo ex62_63 Else GoTo ex60_61

ex60_61:
        If opcode = 60 Then
            ' 060 INC A
            regA = inc8(regA)
            Tstates = Tstates + 4
        Else
            ' 061 DEC A
            regA = dec8(regA)
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex62_63:
        If opcode = 62 Then
            ' 062 LD A,n
            regA = nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 063 CCF
            ccf
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex64_127:
        If (opcode And 32) Then GoTo ex96_127 Else GoTo ex64_95

ex64_95:
        If (opcode And 16) Then GoTo ex80_95 Else GoTo ex64_79

ex64_79:
        If (opcode And 8) Then GoTo ex72_79 Else GoTo ex64_71

ex64_71:
        If (opcode And 4) Then GoTo ex68_71 Else GoTo ex64_67

ex64_67:
        If (opcode And 2) Then GoTo ex66_67 Else GoTo ex64_65

ex64_65:
        If opcode = 64 Then
            ' LD B,B
            Tstates = Tstates + 4
        Else
            ' 65 ' LD B,C
            regB = regC
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex66_67:
        If opcode = 66 Then
            ' LD B,D
            regB = getD
            Tstates = Tstates + 4
        Else
            ' 67 ' LD B,E
            regB = getE
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex68_71:
        If (opcode And 2) Then GoTo ex70_71 Else GoTo ex68_69

ex68_69:
        If opcode = 68 Then
             ' LD B,H
            regB = getH
            Tstates = Tstates + 4
        Else
            ' 69 ' LD B,L
            regB = getL
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex70_71:
        If opcode = 70 Then
            ' LD B,(HL)
            regB = peekb(regHL)
            Tstates = Tstates + 7
        Else
            ' 71 ' LD B,A
            regB = regA
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex72_79:
        If (opcode And 4) Then GoTo ex76_79 Else GoTo ex72_75
        
ex72_75:
        If (opcode And 2) Then GoTo ex74_75 Else GoTo ex72_73
        
ex72_73:
        If opcode = 72 Then
            ' 72 ' LD C,B
            regC = regB
            Tstates = Tstates + 4
        Else
            ' 73 ' LD C,C
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex74_75:
        If opcode = 74 Then
            ' 74 ' LD C,D
            regC = getD
            Tstates = Tstates + 4
        Else
            ' 75 ' LD C,E
            regC = getE
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex76_79:
        If (opcode And 2) Then GoTo ex78_79 Else GoTo ex76_77
        
ex76_77:
        If opcode = 76 Then
            ' 76 ' LD C,H
            regC = getH
            Tstates = Tstates + 4
        Else
            ' 77 ' LD C,L
            regC = getL
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex78_79:
        If opcode = 78 Then
            ' 78 ' LD C,(HL)
            regC = peekb(regHL)
            Tstates = Tstates + 7
        Else
            ' 79 ' LD C,A
            regC = regA
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex80_95:
        Select Case opcode
        Case 80 ' LD D,B
            setD regB
            Tstates = Tstates + 4
        Case 81 ' LD D,C
            setD regC
            Tstates = Tstates + 4
        Case 82 ' LD D,D
            Tstates = Tstates + 4
        Case 83 ' LD D,E
            setD getE
            Tstates = Tstates + 4
        Case 84 ' LD D,H
            setD getH
            Tstates = Tstates + 4
        Case 85 ' LD D,L
            setD getL
            Tstates = Tstates + 4
        Case 86 ' LD D,(HL)
            setD peekb(regHL)
            Tstates = Tstates + 7
        Case 87 ' LD D,A
            setD regA
            Tstates = Tstates + 4
        ' // LD E,*
        Case 88 ' LD E,B
            setE regB
            Tstates = Tstates + 4
        Case 89 ' LD E,C
            setE regC
            Tstates = Tstates + 4
        Case 90 ' LD E,D
            setE getD
            Tstates = Tstates + 4
        Case 91 ' LD E,E
            Tstates = Tstates + 4
        Case 92 ' LD E,H
            setE getH
            Tstates = Tstates + 4
        Case 93 ' LD E,L
            setE getL
            Tstates = Tstates + 4
        Case 94 ' LD E,(HL)
            setE peekb(regHL)
            Tstates = Tstates + 7
        Case 95 ' LD E,A
            setE regA
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next

ex96_127:
        If (opcode And 16) Then GoTo ex112_127 Else GoTo ex96_111

ex96_111:
        Select Case opcode
        Case 96 ' LD H,B
            setH regB
            Tstates = Tstates + 4
        Case 97 ' LD H,C
            setH regC
            Tstates = Tstates + 4
        Case 98 ' LD H,D
            setH getD
            Tstates = Tstates + 4
        Case 99 ' LD H,E
            setH getE
            Tstates = Tstates + 4
        Case 100 ' LD H,H
            Tstates = Tstates + 4
        Case 101 ' LD H,L
            setH getL
            Tstates = Tstates + 4
        Case 102 ' LD H,(HL)
            setH peekb(regHL)
            Tstates = Tstates + 7
        Case 103 ' LD H,A
            setH regA
            Tstates = Tstates + 4
        ' // LD L,*
        Case 104 ' LD L,B
            setL regB
            Tstates = Tstates + 4
        Case 105 ' LD L,C
            setL regC
            Tstates = Tstates + 4
        Case 106 ' LD L,D
            setL getD
            Tstates = Tstates + 4
        Case 107 ' LD L,E
            setL getE
            Tstates = Tstates + 4
        Case 108 ' LD L,H
            setL getH
            Tstates = Tstates + 4
        Case 109 ' LD L,L
            Tstates = Tstates + 4
        Case 110 ' LD L,(HL)
            setL peekb(regHL)
            Tstates = Tstates + 7
        Case 111 ' LD L,A
            setL regA
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next

ex112_127:
        If (opcode And 8) Then GoTo ex120_127 Else GoTo ex112_119
        
ex112_119:
        If (opcode And 4) Then GoTo ex116_119 Else GoTo ex112_115
        
ex112_115:
        If (opcode And 2) Then GoTo ex114_115 Else GoTo ex112_113
        
ex112_113:
        If opcode = 112 Then
            ' 112 ' LD (HL),B
            pokeb regHL, regB
            Tstates = Tstates + 7
        Else
            ' 113 ' LD (HL),C
            pokeb regHL, regC
            Tstates = Tstates + 7
        End If
        GoTo fetch_next
        
ex114_115:
        If opcode = 114 Then
            ' 114 ' LD (HL),D
            pokeb regHL, getD
            Tstates = Tstates + 7
        Else
            ' 115 ' LD (HL),E
            pokeb regHL, getE
            Tstates = Tstates + 7
        End If
        GoTo fetch_next
        
ex116_119:
        Select Case opcode
        Case 116 ' LD (HL),H
            pokeb regHL, getH
            Tstates = Tstates + 7
        Case 117 ' LD (HL),L
            pokeb regHL, getL
            Tstates = Tstates + 7
        Case 118 ' HALT
        
        '--------------------------------------------------+
        '   Is this actually faster than decreasing the PC |
        '   and going to fetch_next?                       |
        '--------------------------------------------------+
    
    halt = True: regPC = (regPC - 1)
    Tstates = Tstates + 4
    
    'Tstates = -(Tcycles + (Tstates Mod 4))      '// Tstates Left After HALT Finished.
    
    'While (Not irqsetLine)                      '// HALT CPU Until Interrupt.

    '    Vdp.interrupts (lineno)
    '    If (lineno < 312&) Then
    '        If (lineno < 192&) Then Vdp.drawLine (lineno)
    '        lineno = lineno + 1
    '    Else
    '        lineno = vbEmpty
    '    End If

    'Wend

    'If GetQueueStatus(QS_MOUSEBUTTON Or QS_KEY Or QS_SENDMESSAGE) Then DoEvents
    'StretchDIBits Form1.hdc, 0&, 0&, SMS_WIDTH, SMS_HEIGHT, 0&, 0&, SMS_WIDTH, SMS_HEIGHT, display(0&), myBMP, 0&, vbSrcCopy

    'ltemp = (timeGetTime - oldTick)
    'If (ltemp < 19&) Then Sleep (19& - ltemp)
    'oldTick = timeGetTime




        Case 119 ' LD (HL),A
            pokeb regHL, regA
            Tstates = Tstates + 7
        End Select
        GoTo fetch_next
        
ex120_127:
        If (opcode And 4) Then GoTo ex124_127 Else GoTo ex120_123

ex120_123:
        If (opcode And 2) Then GoTo ex122_123 Else GoTo ex120_121
        
ex120_121:
        If opcode = 120 Then
            ' 120 ' LD A,B
            regA = regB
            Tstates = Tstates + 4
        Else
            ' 121 ' LD A,C
            regA = regC
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex122_123:
        If opcode = 122 Then
            ' 122 ' LD A,D
            regA = getD
            Tstates = Tstates + 4
        Else
            ' 123 ' LD A,E
            regA = getE
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex124_127:
        If (opcode And 2) Then GoTo ex126_127 Else GoTo ex124_125
        
ex124_125:
        If opcode = 124 Then
            ' 124 ' LD A,H
            regA = getH
            Tstates = Tstates + 4
        Else
            ' 125 ' LD A,L
            regA = getL
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex126_127:
        If opcode = 126 Then
            ' 126 ' LD A,(HL)
            regA = peekb(regHL)
            Tstates = Tstates + 7
        Else
            ' 127 ' LD A,A
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex128_255:
        If (opcode And 64) Then GoTo ex192_255 Else GoTo ex128_191

ex128_191:
        If (opcode And 32) Then GoTo ex160_191 Else GoTo ex128_159

ex128_159:
        Select Case opcode
        ' // ADD A,*
        Case 128 ' ADD A,B
            add_a regB
            Tstates = Tstates + 4
        Case 129 ' ADD A,C
            add_a regC
            Tstates = Tstates + 4
        Case 130 ' ADD A,D
            add_a getD
            Tstates = Tstates + 4
        Case 131 ' ADD A,E
            add_a getE
            Tstates = Tstates + 4
        Case 132 ' ADD A,H
            add_a getH
            Tstates = Tstates + 4
        Case 133 ' ADD A,L
            add_a getL
            Tstates = Tstates + 4
        Case 134 ' ADD A,(HL)
            add_a peekb(regHL)
            Tstates = Tstates + 7
        Case 135 ' ADD A,A
            add_a regA
            Tstates = Tstates + 4
        Case 136 ' ADC A,B
            adc_a regB
            Tstates = Tstates + 4
        Case 137 ' ADC A,C
            adc_a regC
            Tstates = Tstates + 4
        Case 138 ' ADC A,D
            adc_a getD
            Tstates = Tstates + 4
        Case 139 ' ADC A,E
            adc_a getE
            Tstates = Tstates + 4
        Case 140 ' ADC A,H
            adc_a getH
            Tstates = Tstates + 4
        Case 141 ' ADC A,L
            adc_a getL
            Tstates = Tstates + 4
        Case 142 ' ADC A,(HL)
            adc_a peekb(regHL)
            Tstates = Tstates + 7
        Case 143 ' ADC A,A
            adc_a regA
            Tstates = Tstates + 4
        Case 144 ' SUB B
            sub_a regB
            Tstates = Tstates + 4
        Case 145 ' SUB C
            sub_a regC
            Tstates = Tstates + 4
        Case 146 ' SUB D
            sub_a getD
            Tstates = Tstates + 4
        Case 147 ' SUB E
            sub_a getE
            Tstates = Tstates + 4
        Case 148 ' SUB H
            sub_a getH
            Tstates = Tstates + 4
        Case 149 ' SUB L
            sub_a getL
            Tstates = Tstates + 4
        Case 150 ' SUB (HL)
            sub_a peekb(regHL)
            Tstates = Tstates + 7
        Case 151 ' SUB A
            sub_a regA
            Tstates = Tstates + 4
        Case 152 ' SBC A,B
            sbc_a regB
            Tstates = Tstates + 4
        Case 153 ' SBC A,C
            sbc_a regC
            Tstates = Tstates + 4
        Case 154 ' SBC A,D
            sbc_a getD
            Tstates = Tstates + 4
        Case 155 ' SBC A,E
            sbc_a getE
            Tstates = Tstates + 4
        Case 156 ' SBC A,H
            sbc_a getH
            Tstates = Tstates + 4
        Case 157 ' SBC A,L
            sbc_a getL
            Tstates = Tstates + 4
        Case 158 ' SBC A,(HL)
            sbc_a peekb(regHL)
            Tstates = Tstates + 7
        Case 159 ' SBC A,A
            sbc_a regA
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next

ex160_191:
        If (opcode And 16) Then GoTo ex176_191 Else GoTo ex160_175

ex160_175:
        If (opcode And 8) Then GoTo ex168_175 Else GoTo ex160_167

ex160_167:
        If (opcode And 4) Then GoTo ex164_167 Else GoTo ex160_163
        
ex160_163:
        Select Case opcode
        Case 160 ' AND B
            and_a regB
            Tstates = Tstates + 4
        Case 161 ' AND C
            and_a regC
            Tstates = Tstates + 4
        Case 162 ' AND D
            and_a getD
            Tstates = Tstates + 4
        Case 163 ' AND E
            and_a getE
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next
        
ex164_167:
        Select Case opcode
        Case 164 ' AND H
            and_a getH
            Tstates = Tstates + 4
        Case 165 ' AND L
            and_a getL
            Tstates = Tstates + 4
        Case 166 ' AND (HL)
            and_a peekb(regHL)
            Tstates = Tstates + 7
        Case 167 ' AND A
            and_a regA
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next

ex168_175:
        If (opcode And 4) Then GoTo ex172_175 Else GoTo ex168_171
        
ex168_171:
        Select Case opcode
        Case 168 ' XOR B
            xor_a regB
            Tstates = Tstates + 4
        Case 169 ' XOR C
            xor_a regC
            Tstates = Tstates + 4
        Case 170 ' XOR D
            xor_a getD
            Tstates = Tstates + 4
        Case 171 ' XOR E
            xor_a getE
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next
        
ex172_175:
        Select Case opcode
        Case 172 ' XOR H
            xor_a getH
            Tstates = Tstates + 4
        Case 173 ' XOR L
            xor_a getL
            Tstates = Tstates + 4
        Case 174 ' XOR (HL)
            xor_a peekb(regHL)
            Tstates = Tstates + 7
        Case 175 ' XOR A
            regA = 0
            fS = False
            f3 = False
            f5 = False
            fH = False
            fPV = True
            fZ = True
            fN = False
            fC = False
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next

ex176_191:
        Select Case opcode
        Case 176 ' OR B
            or_a regB
            Tstates = Tstates + 4
        Case 177 ' OR C
            or_a regC
            Tstates = Tstates + 4
        Case 178 ' OR D'
            or_a getD
            Tstates = Tstates + 4
        Case 179 ' OR E
            or_a getE
            Tstates = Tstates + 4
        Case 180 ' OR H
            or_a getH
            Tstates = Tstates + 4
        Case 181 ' OR L
            or_a getL
            Tstates = Tstates + 4
        Case 182 ' OR (HL)
            or_a peekb(regHL)
            Tstates = Tstates + 7
        Case 183 ' OR A
            or_a regA
            Tstates = Tstates + 4
        ' // CP
        Case 184 ' CP B
            cp_a regB
            Tstates = Tstates + 4
        Case 185 ' CP C
            cp_a regC
            Tstates = Tstates + 4
        Case 186 ' CP D
            cp_a getD
            Tstates = Tstates + 4
        Case 187 ' CP E
            cp_a getE
            Tstates = Tstates + 4
        Case 188 ' CP H
            cp_a getH
            Tstates = Tstates + 4
        Case 189 ' CP L
            cp_a getL
            Tstates = Tstates + 4
        Case 190 ' CP (HL)
            cp_a peekb(regHL)
            Tstates = Tstates + 7
        Case 191 ' CP A
            cp_a regA
            Tstates = Tstates + 4
        End Select
        GoTo fetch_next

ex192_255:
        If (opcode And 32) Then GoTo ex224_255 Else GoTo ex192_223

ex192_223:
        If (opcode And 16) Then GoTo ex208_223 Else GoTo ex192_207

ex192_207:
        Select Case opcode
        Case 192 ' RET NZ
            If fZ = False Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Case 193 ' POP BC
            setBC popw
            Tstates = Tstates + 10
        Case 194 ' JP NZ,nn
            If fZ = False Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Case 195 ' JP nn
            regPC = peekw(regPC)
            Tstates = Tstates + 10
        Case 196 ' CALL NZ,nn
            If fZ = False Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Case 197 ' PUSH BC
            pushw getBC
            Tstates = Tstates + 11
        Case 198 ' ADD A,n
            add_a nxtpcb()
            Tstates = Tstates + 7
        Case 199 ' RST 0
            pushpc
            regPC = 0
            Tstates = Tstates + 11
        Case 200 ' RET Z
            If fZ Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Case 201 ' RET
            poppc
            Tstates = Tstates + 10
        Case 202 ' JP Z,nn
            If fZ Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Case 203 ' Prefix CB
            Tstates = Tstates + execute_cb
        Case 204 ' CALL Z,nn
            If fZ Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Case 205 ' CALL nn
            pushw regPC + 2
            regPC = nxtpcw
            Tstates = Tstates + 17
        Case 206 ' ADC A,n
            adc_a nxtpcb()
            Tstates = Tstates + 7
        Case 207 ' RST 8
            pushpc
            regPC = 8
            Tstates = Tstates + 11
        End Select
        GoTo fetch_next

ex208_223:
        If (opcode And 8) Then GoTo ex216_223 Else GoTo ex208_215
        
ex208_215:
        Select Case opcode
        Case 208 ' RET NC
            If fC = False Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Case 209 ' POP DE
            regDE = popw
            Tstates = Tstates + 10
        Case 210 '  JP NC,nn
            If fC = False Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Case 211 ' OUT (n),A
            out nxtpcb, regA
            Tstates = Tstates + 11
        Case 212 ' CALL NC,nn
            If fC = False Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Case 213 ' PUSH DE
            pushw regDE
            Tstates = Tstates + 11
        Case 214 ' SUB n
            sub_a nxtpcb()
            Tstates = Tstates + 7
        Case 215 ' RST 16
            pushpc
            regPC = 16
            Tstates = Tstates + 11
        End Select
        GoTo fetch_next
        
ex216_223:
        If (opcode And 4) Then GoTo ex220_223 Else GoTo ex216_219
        
ex216_219:
        Select Case opcode
        Case 216 ' RET C
            If fC Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Case 217 ' EXX
            exx
            Tstates = Tstates + 4
        Case 218 ' JP C,nn
            If fC Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Case 219 ' IN A,(n)
            'IS THIS CORRECT???
            regA = inn(nxtpcb)
            'regA = inn((regA * 256) Or nxtpcb)
            Tstates = Tstates + 11
        End Select
        GoTo fetch_next
        
ex220_223:
        Select Case opcode
        Case 220 ' CALL C,nn
            If fC Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Case 221 ' prefix IX
            regID = regIX
            Tstates = Tstates + execute_id
            ' // ZX81 Specific
            'If (regIX <> regID) And (regID > 8191) Then
                ' // IX has changed - looks like this is a hi-res graphics routine
                ' // 1. Search for hires screen and store it's location in lHiresLoc
            '    lHiresLoc = SearchHiresScreen()
            '    If lHiresLoc > 0 Then
            '        ReInitHiresScreen
            '        gpicDisplay.Cls
            '    End If
            'ElseIf lHiresLoc > 0 Then
            '    lHiresLoc = 0
            '    For lTemp = 0 To 767
            '        sLastScreen(lTemp) = 0
            '    Next lTemp
            '    For lTemp = 0 To 6143
            '        gcBufferBits(lTemp) = 0
            '    Next lTemp
            '    gpicDisplay.Cls
            'End If
            regIX = regID
        Case 222 ' SBC n
            sbc_a nxtpcb()
            Tstates = Tstates + 7
        Case 223 ' RST 24
            pushpc
            regPC = 24
            Tstates = Tstates + 11
        End Select
        GoTo fetch_next

ex224_255:
        If (opcode And 16) Then GoTo ex240_255 Else GoTo ex224_239

ex224_239:
        If (opcode And 8) Then GoTo ex232_239 Else GoTo ex224_231
        
ex224_231:
        If (opcode And 4) Then GoTo ex228_231 Else GoTo ex224_227
        
ex224_227:
        If (opcode And 2) Then GoTo ex226_227 Else GoTo ex224_225
        
ex224_225:
        If opcode = 224 Then
            ' 224 ' RET PO
            If fPV = False Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Else
            ' 225 ' POP HL
            regHL = popw
            Tstates = Tstates + 10
        End If
        GoTo fetch_next
        
ex226_227:
        If opcode = 226 Then
            ' 226 JP PO,nn
            If fPV = False Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Else
            ' 227 ' EX (SP),HL
            ltemp = regHL
            regHL = peekw(regSP)
            pokew regSP, ltemp
            Tstates = Tstates + 19
        End If
        GoTo fetch_next
        
ex228_231:
        If (opcode And 2) Then GoTo ex230_231 Else GoTo ex228_229
        
ex228_229:
        If opcode = 228 Then
            ' 228 ' CALL PO,nn
            If fPV = False Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Else
            ' 229 ' PUSH HL
            pushw regHL
            Tstates = Tstates + 11
        End If
        GoTo fetch_next
        
ex230_231:
        If opcode = 230 Then
            ' 230 ' AND n
            and_a nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 231 ' RST 32
            pushpc
            regPC = 32
            Tstates = Tstates + 11
        End If
        GoTo fetch_next
        
ex232_239:
        If (opcode And 4) Then GoTo ex236_239 Else GoTo ex232_235
        
ex232_235:
        If (opcode And 2) Then GoTo ex234_235 Else GoTo ex232_233
        
ex232_233:
        If opcode = 232 Then
            ' 232 ' RET PE
            If fPV Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Else
            ' 233 ' JP HL
            regPC = regHL
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex234_235:
        If opcode = 234 Then
            ' 234 ' JP PE,nn
            If fPV Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Else
            ' 235 ' EX DE,HL
            ltemp = regHL
            regHL = regDE
            regDE = ltemp
            Tstates = Tstates + 4
        End If
        GoTo fetch_next
        
ex236_239:
        If (opcode And 2) Then GoTo ex238_239 Else GoTo ex236_237
        
ex236_237:
        If opcode = 236 Then
            ' 236 ' CALL PE,nn
            If fPV Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Else
            ' 237 ' prefix ED
            Tstates = Tstates + execute_ed(Tstates)
        End If
        GoTo fetch_next
        
ex238_239:
        If opcode = 238 Then
            ' 238 ' XOR n
            xor_a nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 239 ' RST 40
            pushpc
            regPC = 40
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex240_255:
        If (opcode And 8) Then GoTo ex248_255 Else GoTo ex240_247

ex240_247:
        If (opcode And 4) Then GoTo ex244_247 Else GoTo ex240_243

ex240_243:
        If (opcode And 2) Then GoTo ex242_243 Else GoTo ex240_241

ex240_241:
        If opcode = 240 Then
            ' 240 RET P
            If fS = False Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Else
            ' 241 POP AF
            setAF popw
            Tstates = Tstates + 10
        End If
        GoTo fetch_next

ex242_243:
        If opcode = 242 Then
            ' 242 JP P,nn
            If fS = False Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Else
            ' 243 DI
            intIFF1 = False
            intIFF2 = False
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex244_247:
        If (opcode And 2) Then GoTo ex246_247 Else GoTo ex244_245

ex244_245:
        If opcode = 244 Then
        ' 244 CALL P,nn
            If fS = False Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Else
            ' 245 PUSH AF
            pushw getAF
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex246_247:
        If opcode = 246 Then
            ' 246 OR n
            or_a nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 247 RST 48
            pushpc
            regPC = 48
            Tstates = Tstates + 11
        End If
        GoTo fetch_next

ex248_255:
        If (opcode And 4) Then GoTo ex252_255 Else GoTo ex248_251

ex248_251:
        If (opcode And 2) Then GoTo ex250_251 Else GoTo ex248_249

ex248_249:
        If opcode = 248 Then
            ' 248 RET M
            If fS Then
                poppc
                Tstates = Tstates + 11
            Else
                Tstates = Tstates + 5
            End If
        Else
            ' 249 LD SP,HL
            regSP = regHL
            Tstates = Tstates + 6
        End If
        GoTo fetch_next

ex250_251:
        If opcode = 250 Then
            ' 250 JP M,nn
            If fS Then
                regPC = nxtpcw
            Else
                regPC = regPC + 2
            End If
            Tstates = Tstates + 10
        Else
            ' 251 EI
            intIFF1 = True
            intIFF2 = True
            Tstates = Tstates + 4
        End If
        GoTo fetch_next

ex252_255:
        If (opcode And 2) Then GoTo ex254_255 Else GoTo ex252_253

ex252_253:
        If opcode = 252 Then
            ' 252 CALL M,nn
            If fS Then
                pushw regPC + 2
                regPC = nxtpcw
                Tstates = Tstates + 17
            Else
                regPC = regPC + 2
                Tstates = Tstates + 10
            End If
        Else
            ' 253 prefix IY
            regID = regIY
            Tstates = Tstates + execute_id()
            regIY = regID
        End If
        GoTo fetch_next

ex254_255:
        If opcode = 254 Then
            ' 254 CP n
            cp_a nxtpcb()
            Tstates = Tstates + 7
        Else
            ' 255 RST 56
            pushpc
            regPC = 56
            Tstates = Tstates + 11
        End If
        GoTo fetch_next
End Sub
Private Function qdec8(a As Long) As Long
    qdec8 = (a - 1) And &HFF&
End Function
Private Function execute_id() As Long
    Dim xxx As Long, ltemp As Long, op As Long
    
    
    ' // Yes, I appreciate that GOTO's and labels are a hideous blashphemy!
    ' // However, this code is the fastest possible way of fetching and handling
    ' // Z80 instructions I could come up with. There are only 8 compares per
    ' // instruction fetch rather than between 1 and 255 as required in
    ' // the previous version of vb81 with it's huge Case statement.
    ' //
    ' // I know it's slightly harder to follow the new code, but I think the
    ' // speed increase justifies it. <CC>
    
    
    ' // REFRESH 1
    intRTemp = intRTemp + 1
    
    xxx = nxtpcb
    
    If (xxx And 128) Then GoTo ex_id128_255 Else GoTo ex_id0_127
    
ex_id0_127:
    If (xxx And 64) Then GoTo ex_id64_127 Else GoTo ex_id0_63
    
ex_id0_63:
    If (xxx And 32) Then GoTo ex_id32_63 Else GoTo ex_id0_31
    
ex_id0_31:
    Select Case xxx
    Case 0 To 8
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 9 ' ADD ID,BC
        regID = add16(regID, getBC)
        execute_id = 15
    Case 10 To 24
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 25 ' ADD ID,DE
        regID = add16(regID, regDE)
        execute_id = 15
    Case 26 To 31
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    End Select
    GoTo execute_id_end

ex_id32_63:
    Select Case xxx
    Case 32
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 33 ' LD ID,nn
        regID = nxtpcw
        execute_id = 14
    Case 34 ' LD (nn),ID
        pokew nxtpcw, regID
        execute_id = 20
    Case 35 ' INC ID
        regID = inc16(regID)
        execute_id = 10
    Case 36 ' INC IDH
        setIDH inc8(getIDH)
        execute_id = 9
    Case 37 ' DEC IDH
        setIDH dec8(getIDH)
        execute_id = 9
    Case 38 ' LD IDH,n
        setIDH nxtpcb()
        execute_id = 11
    Case 39, 40
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 41 ' ADD ID,ID
        ltemp = regID
        regID = add16(ltemp, ltemp)
        execute_id = 15
    Case 42 ' LD ID,(nn)
        regID = peekw(nxtpcw)
        execute_id = 20
    Case 43 ' DEC ID
        regID = dec16(regID)
        execute_id = 10
    Case 44 ' INC IDL
        setIDL inc8(getIDL)
        execute_id = 9
    Case 45 ' DEC IDL
        setIDL dec8(getIDL)
        execute_id = 9
    Case 46 ' LD IDL,n
        setIDL nxtpcb()
        execute_id = 11
    Case 47 To 51
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 52 ' INC (ID+d)
        ltemp = id_d
        pokeb ltemp, inc8(peekb(ltemp))
        execute_id = 23
    Case 53 ' DEC (ID+d)
        ltemp = id_d
        pokeb ltemp, dec8(peekb(ltemp))
        execute_id = 23
    Case 54 ' LD (ID+d),n
        ltemp = id_d
        pokeb ltemp, nxtpcb()
        execute_id = 19
    Case 55, 56
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 57 ' ADD ID,SP
        regID = add16(regID, regSP)
        execute_id = 15
    Case 58 To 63
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    End Select
    GoTo execute_id_end
    
ex_id64_127:
    Select Case xxx
    Case 64 To 67
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 68 ' LD B,IDH
        regB = getIDH
        execute_id = 9
    Case 69 ' LD B,IDL
        regB = getIDL
        execute_id = 9
    Case 70 ' LD B,(ID+d)
        regB = peekb(id_d)
        execute_id = 19
    Case 71 To 75
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 76 ' LD C,IDH
        regC = getIDH
        execute_id = 9
    Case 77 ' LD C,IDL
        regC = getIDL
        execute_id = 9
    Case 78 ' LD C,(ID+d)
        regC = peekb(id_d)
        execute_id = 19
    Case 79 To 83
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 84 ' LD D,IDH
        setD getIDH
        execute_id = 9
    Case 85 ' LD D,IDL
        setD getIDL
        execute_id = 9
    Case 86 ' LD D,(ID+d)
        setD peekb(id_d)
        execute_id = 19
    Case 87 To 91
            regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 92 ' LD E,IDH
        setE getIDH
        execute_id = 9
    Case 93 ' LD E,IDL
        setE getIDL
        execute_id = 9
    Case 94 ' LD E,(ID+d)
        setE peekb(id_d)
        execute_id = 19
    Case 95
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 96 ' LD IDH,B
        setIDH regB
        execute_id = 9
    Case 97 ' LD IDH,C
        setIDH regC
        execute_id = 9
    Case 98 ' LD IDH,D
        setIDH getD
        execute_id = 9
    Case 99 ' LD IDH,E
        setIDH getE
        execute_id = 9
    Case 100 ' LD IDH,IDH
        execute_id = 9
    Case 101 ' LD IDH,IDL
        setIDH getIDL
        execute_id = 9
    Case 102 ' LD H,(ID+d)
        setH peekb(id_d)
        execute_id = 19
    Case 103 ' LD IDH,A
        setIDH regA
        execute_id = 9
    Case 104 ' LD IDL,B
        setIDL regB
        execute_id = 9
    Case 105 ' LD IDL,C
        setIDL regC
        execute_id = 9
    Case 106 ' LD IDL,D
        setIDL getD
        execute_id = 9
    Case 107 ' LD IDL,E
        setIDL getE
        execute_id = 9
    Case 108 ' LD IDL,IDH
        setIDL getIDH
        execute_id = 9
    Case 109 ' LD IDL,IDL
        execute_id = 9
    Case 110 ' LD L,(ID+d)
        setL peekb(id_d)
        execute_id = 19
    Case 111 ' LD IDL,A
        setIDL regA
        execute_id = 9
    Case 112 ' LD (ID+d),B
        pokeb id_d, regB
        execute_id = 19
    Case 113 ' LD (ID+d),C
        pokeb id_d, regC
        execute_id = 19
    Case 114 ' LD (ID+d),D
        pokeb id_d, getD
        execute_id = 19
    Case 115 ' LD (ID+d),E
        pokeb id_d, getE
        execute_id = 19
    Case 116 ' LD (ID+d),H
        pokeb id_d, getH
        execute_id = 19
    Case 117 ' LD (ID+d),L
        pokeb id_d, getL
        execute_id = 19
    Case 118 ' UNKNOWN
        MsgBox "Unknown ID instruction " & xxx & " at " & regPC
    Case 119 ' LD (ID+d),A
        pokeb id_d, regA
        execute_id = 19
    Case 120 To 123
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 124 ' LD A,IDH
        regA = getIDH
        execute_id = 9
    Case 125 ' LD A,IDL
        regA = getIDL
        execute_id = 9
    Case 126 ' LD A,(ID+d)
        regA = peekb(id_d)
        execute_id = 19
    Case 127
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    End Select
    GoTo execute_id_end

ex_id128_255:
    If (xxx And 64) Then GoTo ex_id192_255 Else GoTo ex_id128_191
    
ex_id128_191:
    If (xxx And 32) Then GoTo ex_id160_191 Else GoTo ex_id128_159
    
ex_id128_159:
    Select Case xxx
    Case 128 To 131
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 132 ' ADD A,IDH
        add_a getIDH
        execute_id = 9
    Case 133 ' ADD A,IDL
        add_a getIDL
        execute_id = 9
    Case 134 ' ADD A,(ID+d)
        add_a peekb(id_d)
        execute_id = 19
    Case 135 To 139
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 140 ' ADC A,IDH
        adc_a getIDH
        execute_id = 9
    Case 141 ' ADC A,IDL
        adc_a getIDL
        execute_id = 9
    Case 142 ' ADC A,(ID+d)
        adc_a peekb(id_d)
        execute_id = 19
    Case 143 To 147
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 148 ' SUB IDH
        sub_a getIDH
        execute_id = 9
    Case 149 ' SUB IDL
        sub_a getIDL
        execute_id = 9
    Case 150 ' SUB (ID+d)
        sub_a peekb(id_d)
        execute_id = 19
    Case 151 To 155
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 156 ' SBC A,IDH
        sbc_a getIDH
        execute_id = 9
    Case 157 ' SBC A,IDL
        sbc_a getIDL
        execute_id = 9
    Case 158 ' SBC A,(ID+d)
        sbc_a peekb(id_d)
        execute_id = 19
    Case 159
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    End Select
    GoTo execute_id_end
    
ex_id160_191:
    Select Case xxx
    Case 160 To 163
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 164 ' AND IDH
        and_a getIDH
        execute_id = 9
    Case 165 ' AND IDL
        and_a getIDL
        execute_id = 9
    Case 166 ' AND (ID+d)
        and_a peekb(id_d)
        execute_id = 19
    Case 167 To 171
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 172 ' XOR IDH
        xor_a getIDH
        execute_id = 9
    Case 173 ' XOR IDL
        xor_a getIDL
        execute_id = 9
    Case 174 'XOR (ID+d)
        xor_a (peekb(id_d))
        execute_id = 19
    Case 175 To 179
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 180 ' OR IDH
        or_a getIDH
        execute_id = 9
    Case 181 ' OR IDL
        or_a getIDL
        execute_id = 9
    Case 182 ' OR (ID+d)
        or_a peekb(id_d)
        execute_id = 19
    Case 183 To 187
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 188 ' CP IDH
        cp_a getIDH
        execute_id = 9
    Case 189 ' CP IDL
        cp_a getIDL
        execute_id = 9
    Case 190 ' CP (ID+d)
        cp_a peekb(id_d)
        execute_id = 19
    Case 191
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    End Select
    GoTo execute_id_end
    
ex_id192_255:
    Select Case xxx
    Case 192 To 202
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 203 ' prefix CB
        ltemp = id_d
        op = nxtpcb()
        execute_id_cb op, ltemp
        If ((op And &HC0&) = &H40&) Then execute_id = 20 Else execute_id = 23
    Case 204 To 224
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 225 ' POP ID
        regID = popw()
        execute_id = 14
    Case 226
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 227 ' EX (SP),ID
        ltemp = regID
        regID = peekw(regSP)
        pokew regSP, ltemp
        execute_id = 23
    Case 228
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 229 ' PUSH ID
        pushw regID
        execute_id = 15
    Case 230 To 232
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 233 ' JP ID
        regPC = regID
        execute_id = 8
    Case 234 To 248
        regPC = dec16(regPC)
        ' // REFRESH -1
        intRTemp = intRTemp - 1
        execute_id = 4
    Case 249 ' LD SP,ID
        regSP = regID
        execute_id = 10
    Case Else
        MsgBox "Unknown ID instruction " & xxx & " at " & regPC
    End Select
    
execute_id_end:
End Function


Private Sub setIDH(byteval As Long)
    regID = ((byteval * 256&) And &HFF00&) Or (regID And &HFF&)
End Sub
Private Sub setIDL(byteval As Long)
    regID = (regID And &HFF00&) Or (byteval And &HFF&)
End Sub

Private Function getIDH() As Long
    getIDH = glMemAddrDiv256(regID) And &HFF&
End Function
Private Function getIDL() As Long
    getIDL = regID And &HFF&
End Function

Private Function inc16(a As Long) As Long
    inc16 = (a + 1) And &HFFFF&
End Function
Function nxtpcw() As Long
    nxtpcw = peekb(regPC) + (peekb(regPC + 1) * 256&)
    regPC = regPC + 2
End Function

Function nxtpcb() As Long
    nxtpcb = peekb(regPC)
    regPC = regPC + 1
End Function


Function peekb(ByVal address As Long) As Long


        If (address < &HC000&) Then     '// ROM (And Cart RAM).
                                        peekb = cartRom(address)
    ElseIf (address < &HE000&) Then     '// RAM (Onboard).
                                        peekb = cartRam(address - &HC000&)
    ElseIf (address < &H10000) Then     '// RAM (Mirrored)
                                        peekb = cartRam(address - &HE000&)
    End If


End Function

Sub setD(l As Long)
    regDE = (l * 256#) Or (regDE And &HFF&)
End Sub
Sub setE(l As Long)
    regDE = (regDE And &HFF00&) Or l
End Sub

Sub setF(b As Byte)
    fS = (b And F_S) <> 0
    fZ = (b And F_Z) <> 0
    f5 = (b And F_5) <> 0
    fH = (b And F_H) <> 0
    f3 = (b And F_3) <> 0
    fPV = (b And F_PV) <> 0
    fN = (b And F_N) <> 0
    fC = (b And F_C) <> 0
End Sub

Sub setH(l As Long)
    regHL = (l * 256#) Or (regHL And &HFF&)
End Sub


Sub setL(l As Long)
    regHL = (regHL And &HFF00&) Or l
End Sub




Function sra(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H1&) <> 0
    ans = ShiftRight(ans, 1&) Or (ans And &H80&)
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
    
    sra = ans
End Function

Function srl(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H1&) <> 0
    ans = ShiftRight(ans, 1&) 'glMemAddrDiv2(ans)
        
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
    
    srl = ans
End Function

Function sls(ByVal ans As Long) As Long
    Dim c As Long
    
    c = (ans And &H80&) <> 0
    ans = ((ans * 2) Or &H1) And &HFF
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fPV = Parity(ans)
    fH = False
    fN = False
    fC = c
    
    sls = ans
End Function


Sub sub_a(b As Long)
    Dim a As Long, wans As Long, ans As Long
    
    a = regA
    wans = a - b
    ans = wans And &HFF&
    
    fS = (ans And F_S) <> 0
    f3 = (ans And F_3) <> 0
    f5 = (ans And F_5) <> 0
    fZ = (ans = 0)
    fC = (wans And &H100&) <> 0
    fPV = ((a Xor b) And (a Xor ans) And &H80&) <> 0
    
    fH = (((a And &HF&) - (b And &HF)) And F_H) <> 0
    fN = True
    
    regA = ans
End Sub


Sub xor_a(b As Long)
    regA = (regA Xor b) And &HFF&
    
    fS = (regA And F_S) <> 0
    f3 = (regA And F_3) <> 0
    f5 = (regA And F_5) <> 0
    fH = False
    fPV = Parity(regA)
    fZ = (regA = 0)
    fN = False
    fC = False
End Sub

Public Sub Z80Reset()
    Dim iCounter As Integer
    

    regA = 0
    setF 0
    setBC 0
    regDE = 0
    regHL = 0
    
    regPC = 0
    regSP = &HDFF0&
    
    
    exx
    ex_af_af
    
    regA = 0
    setF 0
    setBC 0
    regDE = 0
    regHL = 0
    
    regIX = 0
    regIY = 0
    
    intR = 128
    intRTemp = 128
    
    intI = 0
    intIFF1 = False
    intIFF2 = False
    intIM = 0
    halt = False

End Sub


