Attribute VB_Name = "BASECONV"
'-----------------------------------------------------------------
'   Module:     mBases
'   ©2000 Trinet Ltd, http://www.trinet.co.uk
'   Author:     R. Deeming (richard@trinet.co.uk)
'
'   Purpose:    To provide simple conversion between different
'               number bases, including fractional parts.
'
'               An example of advanced conversion is included:
'               the sexagesimal (base 60) system. For this
'               example, the numbers are represented as a
'               2-digit number from 0 to 59, followed by a colon.
'               For example, ConvertBase("12:30:", ebSexagesimal)
'               will return 750 (12 * 60 + 30). Note that the
'               structure of the number is strict - "12:30" and
'               "12:5:" are not valid numbers, but "12:30:" and
'               "12:05:" are.
'
'               When there are errors in the number, the code
'               will raise error 13, with a meaningful message.
'               To change the error number, change the constant
'               ERROR_NUMBER, or modify the code.
'
'   You may use and distribute this code, but you may not charge
'   for it or present it as your own work. If you find any bugs
'   in this code, please notify the author.
'   This code is provided "As-Is" - if it doesn't work, we accept
'   no responsibility.
'-----------------------------------------------------------------

'-----------------------------------------------------------------
' A complete explanation of bases is beyond the scope of this
' code, but a brief explanation is included.
'
' A number is represented as a series of digits, arranged in
' columns. The number of available digits is the same as the
' base of the number. Each column is indexed, starting at 0 for
' the "units" column, and the value of the column is equal to
' the base raised to the power of the index. We are accustomed to
' seeing this in decimal (base 10):
'
'   Index   2   1   0   .   -1  -2
'   ------------------------------
'   Value   2   3   4   .    2   5
'
' representing two hundred and thirty four and one quarter. This
' also applies to other bases. For example, in binary (base 2):
'
'   Index   2   1   0   .   -1  -2
'   ------------------------------
'   Value   1   0   1   .    0   1
'
' represents five and one quarter -
' 1 * 2^2 + 0 * 2^1 + 1 * 2^0 + 0 * 2^-1 + 1 * 2^-2
' = 4 + 0 + 1 + 0 + 0.25 = 5.25
'
' This concept extends to bases of any whole number greater than
' 1, and is limited only by the number of available symbols.
'-----------------------------------------------------------------

'-----------------------------------------------------------------
' To add more bases, perform the following steps:
'   1. Add the base to the Bases enum, and set its value to
'      the base (e.g. for base 4, add "ebBase4 = 4&"
'
'   2. If the base is greater than 10, or requires non-standard
'      representation of the numbers, add a private const with
'      the characters used to represent the digits (see
'      hexadecimal for an example), or modify the DigitLength
'      function and other functions (see ebSexagesimal)
'
'   3. In the procedures ConvertDigit and DeconvertDigit, add
'      the new base to the select statement. If you are using a
'      standard representation of a base <= 10, add the new enum
'      value to the first select statement. Otherwise, copy and
'      modify the hexadecimal option.
'
' Example:
'   Add ebAlphabet, a base 26 option in which all digits are
'   represented as letters of the alphabet.
'
'   Private Const AlphaChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'
'   Public Enum Bases
'   ...
'       ebAlphabet = 26&
'
'   Private Function ConvertDigit(...)
'   ...
'       Case ebAlphabet:
'           ConvertDigit = Mid$(AlphaChars, lngDigit + 1, 1)
'
'   Private Function DeconvertDigit(...)
'   ...
'       Case ebAlphabet:
'           lngTemp = InStr(1, AlphaChars, strDigit)
'           If lngTemp = 0 Then
'               err.raise error_number, "DeconvertDigit", "Invalid Alpha Character"
'           Else
'               DeconvertDigit = lngTemp - 1
'           End If
'
'-----------------------------------------------------------------

'-----------------------------------------------------------------
' Bases    : OptikaBase42 & Base36
' Added by : Jacob Russo
' Date     : 1/1/2003
' Purpose  : To convert Optika Base 42 and Standard Base 36 numbers
'            to and from Decimal
' Date     : 7/24/2003
' Purpose  : Corrected OptikaBase42Chars by Adding the "@" at position 14
'            per Optika code
'-----------------------------------------------------------------

Option Explicit

Private Const ERROR_NUMBER = 13&

Private Const HexChars = "0123456789ABCDEF"
Private Const Base36Chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const OptikaBase42Chars = "!$%-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ_"
         
Public Enum Bases
    ebBinary = 2&
    ebOctal = 8&
    ebDecimal = 10&
    ebHexadecimal = 16&
    ebBase36 = 36&
    ebOptikaBase42 = 42&
    ebSexagesimal = 60&     'Base 60, e.g. time
End Enum

Private Function DigitLength(BASE As Bases) As Long
'Return the length of a digit in a given base
    Select Case BASE
        Case ebSexagesimal:
            DigitLength = 3
        
        'Add other special cases here
        
        Case Else   'ebBinary, ebOctal, ebDecimal, ebBase36, ebOptikaBase42, ebHexadecimal
            DigitLength = 1
    End Select
End Function

Private Function Floor(ByVal Number As Double) As Double
'Return the floor of the number
'(the highest whole number less than or equal to the number)
    If Int(Number) > Number Then
        Floor = Int(Number) - 1
    Else
        Floor = Int(Number)
    End If
End Function

Private Function GetNumDec(dblTemp As Double, PadTo As Long, BASE As Bases) As Long
'Return the number of digits required to represent
'the number dblTemp in the specified base, padded
'to the nearest multiple of PadTo

Dim lTemp As Long, lTempPad As Double

    lTemp = Floor(Log(dblTemp) / Log(BASE)) + 1
    If PadTo > 1 Then
        lTempPad = lTemp / CDbl(PadTo)
        If lTempPad > Floor(lTempPad) Then
            lTempPad = 1 + Floor(lTempPad) - lTempPad
            lTempPad = lTempPad * PadTo
            
            lTemp = lTemp + lTempPad
        End If
    End If
    GetNumDec = lTemp
End Function

Private Function ConvertDigit(lngDigit As Long, BASE As Bases) As String
'Convert a single digit to the specified base
    If lngDigit >= BASE Then
        Err.Raise ERROR_NUMBER, "ConvertDigit", "Invalid digit for base"
    Else
        Select Case BASE
            Case ebBinary, ebOctal, ebDecimal:
                ConvertDigit = CStr(lngDigit)
            
            Case ebHexadecimal:
                ConvertDigit = Mid$(HexChars, lngDigit + 1, 1)
            
            Case ebBase36:
                ConvertDigit = Mid$(Base36Chars, lngDigit + 1, 1)
            
            Case ebOptikaBase42:
                ConvertDigit = Mid$(OptikaBase42Chars, lngDigit + 1, 1)
            
            Case ebSexagesimal:
                ConvertDigit = Right$("00" & CStr(lngDigit), 2) & ":"
                
            'Add other bases here
            Case Else: Err.Raise ERROR_NUMBER, "ConvertDigit", "Unknown base"
        End Select
    End If
End Function

Private Function DeconvertDigit(strDigit As String, BASE As Bases) As Long
'Convert a single digit from the specified base to decimal
Dim lngTemp As Long
    Select Case BASE
        Case ebBinary, ebOctal, ebDecimal:
            If IsNumeric(strDigit) Then
                lngTemp = CDbl(strDigit)
                If lngTemp < BASE Then
                    DeconvertDigit = lngTemp
                Else
'                    Err.Raise ERROR_NUMBER, "DeconvertDigit", "Invalid digit for base"
                    MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
                End If
            Else
                    MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid character [" & BASE & "]"
            End If
            
        Case ebHexadecimal:
            lngTemp = InStr(1, HexChars, strDigit)
            If lngTemp = 0 Then
                    MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
            Else
                DeconvertDigit = lngTemp - 1
            End If
            
        Case ebBase36:
            lngTemp = InStr(1, Base36Chars, strDigit)
            If lngTemp = 0 Then
                    MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
            Else
                DeconvertDigit = lngTemp - 1
            End If
        
        Case ebOptikaBase42:
            lngTemp = InStr(1, OptikaBase42Chars, strDigit)
            If lngTemp = 0 Then
                    MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
            Else
                DeconvertDigit = lngTemp - 1
            End If
        
        Case ebSexagesimal:
            If Len(strDigit) = 3 Then
                If Right$(strDigit, 1) = ":" And IsNumeric(Left$(strDigit, 2)) Then
                    lngTemp = CDbl(Left$(strDigit, 2))
                    If lngTemp < BASE Then
                        DeconvertDigit = lngTemp
                    Else
                        MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
                    End If
                Else
                    MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
                End If
            Else
                MsgBox "ERROR: DeconvertDigit [" & strDigit & "] Invalid digit for base [" & BASE & "]"
        End If
        
        'Add other bases here
        
        Case Else:
                MsgBox "ERROR: DeconvertDigit [" & strDigit & "] UNKNOWN BASE [" & BASE & "]"
    End Select
End Function

Private Function ConvertDec2Base(ByVal Number, ByVal BASE As Bases, Optional NumDecimals As Long = -1, Optional Tolerance As Double = 1E-27, Optional PadTo As Long = 0) As String
'Convert Number from decimal to the specified base,
'with NumDecimals fractional digits (or to within tolerance),
'padded to the nearest multiple of PadTo

Dim dblTemp As Double
Dim lCDec As Long
Dim lDigit As Long
Dim dblPwr As Double
Dim strTemp As String

    If Not IsNumeric(Number) Then
        Err.Raise ERROR_NUMBER, "ConvertDec2Base", "Number must be decimal"
    ElseIf BASE < 2 Then
        Err.Raise ERROR_NUMBER, "ConvertDec2Base", "Invalid base"
    Else
        'Negative tolerance could cause loops
        Tolerance = Abs(Tolerance)
        
        dblTemp = CDbl(Number)
        If dblTemp < 0 Then
            strTemp = "-"
            dblTemp = -dblTemp
        End If
        
        lCDec = GetNumDec(dblTemp, PadTo, BASE)
        
        'Integer part
        If lCDec = 0 Then
            strTemp = strTemp & "0"
        Else
            Do Until lCDec = 0
                lCDec = lCDec - 1
                dblPwr = BASE ^ lCDec
                lDigit = 0
                Do While dblTemp >= dblPwr
                    lDigit = lDigit + 1
                    dblTemp = dblTemp - dblPwr
                Loop
                strTemp = strTemp & ConvertDigit(lDigit, BASE)
            Loop
        End If
        
        'Fractional part
        If dblTemp > Tolerance And (NumDecimals > 0 Or (NumDecimals = -1 And Tolerance > 0)) Then
            strTemp = strTemp & "."
            Do While dblTemp > Tolerance And (lCDec > (-NumDecimals) Or NumDecimals = -1)
                lCDec = lCDec - 1
                dblPwr = BASE ^ lCDec
                lDigit = 0
                Do While dblTemp >= dblPwr
                    lDigit = lDigit + 1
                    dblTemp = dblTemp - dblPwr
                Loop
                strTemp = strTemp & ConvertDigit(lDigit, BASE)
            Loop
        End If
        ConvertDec2Base = strTemp
    End If
End Function

Private Function ConvertBase2Dec(ByVal Number As String, ByVal BASE As Bases) As Double
'Convert the number from the specified base to decimal
Dim dblTemp As Double
Dim strDigit As String, lngDigit As Long, i As Long
Dim lngPwr As Long, lngSign As Long, lngDigitSize
    
    If BASE < 2 Then
        Err.Raise ERROR_NUMBER, "ConvertBase2Dec", "Invalid Base"
    Else
        lngDigitSize = DigitLength(BASE)
        lngPwr = 0
        lngSign = 1
        i = 1
        Do Until i > Len(Number)
            strDigit = Mid$(Number, i, lngDigitSize)
            If Left$(strDigit, 1) = "." Then
                i = i + 1
                If lngPwr = 0 Then
                    lngPwr = 1
                Else
                    Err.Raise ERROR_NUMBER, "ConvertBase2Dec", "More than one decimal point"
                End If
            ElseIf Left$(strDigit, 1) = "-" Then
                i = i + 1
                If lngPwr = 0 And dblTemp = 0 Then
                    lngSign = -lngSign
                Else
                    Err.Raise ERROR_NUMBER, "ConvertBase2Dec", "Invalid negation"
                End If
            Else
                i = i + lngDigitSize
                lngDigit = DeconvertDigit(strDigit, BASE)
                dblTemp = dblTemp * BASE + lngDigit
                lngPwr = lngPwr * BASE
            End If
        Loop
        If lngPwr > 1 Then
            ConvertBase2Dec = CDbl(lngSign) * (dblTemp / CDbl(lngPwr))
        Else
            ConvertBase2Dec = CDbl(lngSign) * dblTemp
        End If
    End If
End Function

Public Function ConvertBase(ByVal Number, ByVal FromBASE As Bases, Optional ByVal ToBASE As Bases = ebDecimal, Optional NumDecimals As Long = -1, Optional Tolerance As Double = 1E-27, Optional PadTo As Long = 0) As Variant
'Convert a number from one base to another.
'Parameters:
'       Number      A numeric value (when FromBASE = ebDecimal) or
'                   a string representing the number to convert
'
'       FromBASE    The base to convert from (enumeration)
'
'       ToBASE      (Optional) The base to convert to. Default = Decimal
'
'       NumDecimals (Optional) The number of decimal places to include
'                   when converting a fractional number to a non-decimal
'                   Specify 0 for integer only, or -1 to use tolerance.
'                   (This prevents problems with infinte loops)
'
'       Tolerance   (Optional) The value at which to terminate the
'                   fractional representation. If NumDecimals = -1
'                   and Tolerance = 0, no attempt will be made, to
'                   avoid an infinite loop. The sign of the tolerance
'                   is ignored.
'
'       PadTo       (Optional) Used to pad the non-decimal number
'                   to a given length. EG, binary numbers are
'                   normally shown with a multiple of 8 digits, so
'                   you would specify 8. Use 0 to avoid padding.
'
'Returns:
'       Either a double (if ToBASE = ebDecimal) or a string representing
'       the converted number.

Dim dblDec As Double
    If FromBASE = ebDecimal Then
        If IsNumeric(Number) Then
            dblDec = CDbl(Number)
        Else
            Err.Raise ERROR_NUMBER, "ConvertBase", "Not a decimal number"
        End If
    Else
        dblDec = ConvertBase2Dec(CStr(Number), FromBASE)
    End If
    If ToBASE = ebDecimal Then
        ConvertBase = dblDec
    Else
        ConvertBase = ConvertDec2Base(dblDec, ToBASE, NumDecimals, Tolerance, PadTo)
    End If
End Function
