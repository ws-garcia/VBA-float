VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Float"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
' GENERAL INFO:
' This class module was born with the purpose of incorporating in it large integer arithmetic functions such as addition, subtraction,
' multiplication and division. Although the class is called `Float`, it has focused its efforts on allowing the user to represent the same number
' using different cohorts.
'
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Private Enum enumErrorDescription
    ICI = 1 '"Invalid Character Input."
    IPP = 2 '"Invalid Point Position."
    SISI = 3 '"Significand Invalid Sign Input"
    EMISI = 4 '"Exponent Magnitude Invalid Sign Input"
    EMINSI = 5 '"Exponent Magnitude Invalid Notation Symbol Input"
    SSICI = 6 '"Significand Sign Invalid Character Input"
End Enum
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
Private c_ErrorDescription As String
Private c_ErrorDescriptionNumber As enumErrorDescription
Private c_IsZero As Boolean
Private c_LastRemainder As String
Private c_Magnitude As String
Private c_ProperlyRepresented As Boolean
Private c_Representation As String
Private c_Represented As Boolean
Private c_Sign As Boolean
Private c_Significand As String
Private c_SignificandDEC As String
Private c_SignificandINT As String
Private c_Value As String
Private StrSign As String
Private StrIntPart As String
Private StrDecPart As String
Private StrDecNotation As String
Private strExponent As String
Private SumFstOperandZero As Boolean
Private SumSndOperandZero As Boolean
Private Const strZero As String = "0"
Private Const strOne As String = "1"
Private Const strTwo As String = "2"
Private Const strNegativeOne As String = "-1"
Private Const strDot  As String = "."
Private Const Minus As String = "-"
Private Const Plus As String = "+"
Private Const SciNotSymbol As String = "E"
Private Const strDivError As String = "#NAN!"
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
Property Get ErrorDescription()
    ErrorDescription = c_ErrorDescription
End Property

Property Get IsZero() As Boolean
    IsZero = c_IsZero
End Property

Property Get LastRemainder() As String
    LastRemainder = c_LastRemainder
End Property

Property Get Magnitude() As String
    Magnitude = c_Magnitude
End Property

Property Get ProperlyRepresented() As Boolean
    ProperlyRepresented = c_ProperlyRepresented
End Property

Property Get Representation() As String
    Representation = c_Representation
End Property

Property Get Represented() As Boolean
    Represented = c_Represented
End Property

Property Get Sign() As Boolean
    Sign = c_Sign
End Property

Property Get Significand() As String
    Significand = c_Significand
End Property

Property Get SignificandDecimalPart() As String
    SignificandDecimalPart = c_SignificandDEC
End Property

Property Get SignificandIntegerPart() As String
    SignificandIntegerPart = c_SignificandINT
End Property

Property Get value() As String
    value = c_Value
End Property

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:

''' <summary>
''' Returns "True" if the absolute value of the first number is greater than the second.
''' </summary>
Private Function AbsGT(ByVal StrNum1 As String, ByVal StrNum2 As String) As Boolean
    Dim A As String, B As String
    Dim i As Long, n As Long
    
    If Mid(StrNum1, 1, 1) = Minus Then
        StrNum1 = Mid(StrNum1, 2)
    End If
    If Mid(StrNum2, 1, 1) = Minus Then
        StrNum2 = Mid(StrNum2, 2)
    End If
    StrNum1 = ZeroRemoval(StrNum1)
    StrNum2 = ZeroRemoval(StrNum2)
    If Len(StrNum1) > Len(StrNum2) Then
        A = StrNum1: B = StrNum2
    ElseIf Len(StrNum1) < Len(StrNum2) Then
        A = StrNum2: B = StrNum1
    Else
        n = Fix(Len(StrNum1) / 28)
        If 28 * n - Len(StrNum1) <> 0 Then
            n = n + 1
        End If
        Do While i <= n - 1 And Mid(StrNum1, 28 * i + 1, 28) = Mid(StrNum2, 28 * i + 1, 28)
            i = i + 1
        Loop
        If i < n Then
            If CDec(Mid(StrNum1, 28 * i + 1, 28)) > CDec(Mid(StrNum2, 28 * i + 1, 28)) Then
                A = StrNum1: B = StrNum2
            Else
                A = StrNum2: B = StrNum1
            End If
        End If
    End If
    If A = StrNum1 And A <> B Then
        AbsGT = True
    Else
        AbsGT = False
    End If
End Function

Private Function ApplyLawOfSigns(ByVal expression As String) As String
    Dim tmpResult As String
    
    tmpResult = expression
    Do While tmpResult Like "*[+-][+-]*"
        SignSubstitution tmpResult, "+-", Minus
        SignSubstitution tmpResult, "-+", Minus
        SignSubstitution tmpResult, "++", Plus
        SignSubstitution tmpResult, "--", Plus
    Loop
    ApplyLawOfSigns = tmpResult
End Function

''' <summary>
''' This source method computes the division (A / B) of two big integers A and B
''' by divide and correct method using chunks of size RadixBaseExponent selected
''' so that the radix(base) is B=10^RadixBaseExponent.
'''
''' This is an adaptation of Algorithm 1.6 from "Modern Computer Arithmetic"
''' by Richard P. Brent and Paul Zimmermann, 2010.
''' </summary>
Private Function BaseCaseDivRem(ByVal Dividend As String, ByVal Divisor As String, _
                                ByRef orem As String, Optional RadixBaseExponent As Integer = 14) As String
    Dim A As String, B As String
    Dim Base As Variant
    Dim BaseReducedByOne As Variant
    Dim bn As Variant
    Dim correctionInteger As String
    Dim correctionsCount As Long
    Dim Delta As Long
    Dim DividendSign As Boolean
    Dim DivisordSign As Boolean
    Dim DivSignFlag As Boolean
    Dim fisrtDivisorWord As Variant
    Dim jCounter As Long, m As Long
    Dim LessWordCount As Long
    Dim MaxWordCount As Long
    Dim mProd As Variant
    Dim Normalized As Boolean
    Dim qj_ As Variant, qj As String
    Dim scalingFactor As String
    Dim sfx As Long, k As Variant
    Dim TIntcheck As String
    Dim tmpRem As String
    Dim tmpResult As String
    
    If Not IsNumeric(RadixBaseExponent) Or RadixBaseExponent > 14 Or RadixBaseExponent < 0 Then RadixBaseExponent = 14
    A = Dividend
    B = Divisor
    DividendSign = GetNumberSign(A)
    DivisordSign = GetNumberSign(B)
    If Zero(A) Then
        BaseCaseDivRem = strZero
        orem = B
    Else
        If AbsGT(B, A) Then
            BaseCaseDivRem = strZero
        ElseIf Zero(B) Then
            BaseCaseDivRem = strDivError
        ElseIf A = B Then
            BaseCaseDivRem = strOne
            orem = strZero
        Else
            A = ZeroRemoval(A)
            B = ZeroRemoval(B)
            Base = CDec(10 ^ RadixBaseExponent)
            DivSignFlag = GetDivisionSign(DividendSign, DivisordSign)
            LessWordCount = WordCount(B, RadixBaseExponent)         'n
            '@-------------------------------------------------------------------------------
            ' Hacker’s delight / Henry S. Warren, Jr. -- 2nd ed., 2013
            If LessWordCount = 1 Then 'Take care of the case of a single-digit divisor here.
                MaxWordCount = WordCount(A, RadixBaseExponent)
                A = PrefixZeros(A, RadixBaseExponent * MaxWordCount - Len(A))
                k = CDec(0)
                For jCounter = 0 To MaxWordCount - 1
                    mProd = k * Base + CDec(MidB$(A, 2 * jCounter * RadixBaseExponent + 1, 2 * RadixBaseExponent))
                    qj_ = Fix(mProd / CDec(B))
                    k = mProd - qj_ * CDec(B)
                    tmpResult = tmpResult + PrefixZeros(CStr(qj_), RadixBaseExponent - Len(CStr(qj_)))
                Next
                orem = CStr(k)
            '@-------------------------------------------------------------------------------
            Else
                sfx = RadixBaseExponent * LessWordCount - Len(B)
                fisrtDivisorWord = CDec(Mid(B, 1, RadixBaseExponent - sfx))
                If (fisrtDivisorWord - Base / 2) < 0 Then
                    scalingFactor = CStr(Fix(CDec(Base) / CDec(Mid(B, 1, RadixBaseExponent - sfx) + 1)))
                    B = IntKBMult(B, scalingFactor)
                    A = IntKBMult(A, scalingFactor)
                    LessWordCount = WordCount(B, RadixBaseExponent)
                    sfx = RadixBaseExponent * LessWordCount - Len(B)
                    Normalized = True
                End If
                If sfx > 0 Then  'Scale to last word
                    B = SufixZeros(B, sfx)
                    A = SufixZeros(A, sfx)
                End If
                MaxWordCount = WordCount(A, RadixBaseExponent)
                A = PrefixZeros(A, RadixBaseExponent * MaxWordCount - Len(A))
                m = MaxWordCount - LessWordCount
                TIntcheck = SufixZeros(B, m * RadixBaseExponent)
                BaseReducedByOne = CDec(Base - 1)
                Delta = 0
                Select Case Compare2(A, TIntcheck)
                    Case strZero, strOne        'A >= B * 10^(m * radix)
                        tmpResult = strOne
                        A = IntStrSum(A, Minus + TIntcheck)
                        Delta = RadixBaseExponent * WordCount(A, RadixBaseExponent) - Len(A)
                    Case Else
                        tmpResult = strZero
                End Select
                bn = CDec(MidB$(B, 1, 2 * RadixBaseExponent))             'B last word
                For jCounter = m - 1 To 0 Step -1
                    qj_ = Fix(CDec(MidB$(A, 1, 4 * RadixBaseExponent - (2 * Delta))) / bn) '2 words by 1 word
                    If qj_ > BaseReducedByOne Then
                        qj = CStr(BaseReducedByOne)
                    Else
                        qj = CStr(qj_)
                    End If
                    A = IntStrSum(A, Minus + IntKBMult(SufixZeros(qj, jCounter * RadixBaseExponent), B))
                    Do While GT(strZero, A)
                        correctionsCount = correctionsCount + 1
                        If correctionInteger = vbNullString Then
                            correctionInteger = SufixZeros(B, jCounter * RadixBaseExponent)
                        End If
                        qj = IntStrSum(qj, strNegativeOne)
                        A = IntStrSum(A, correctionInteger)
                    Loop
                    correctionInteger = vbNullString
                    tmpResult = tmpResult + PrefixZeros(qj, RadixBaseExponent - Len(qj))
                    If jCounter Then
                        Delta = RadixBaseExponent * WordCount(A, RadixBaseExponent) - Len(A)
                    End If
                Next jCounter
                If Normalized Then
                    orem = BaseCaseDivRem(A, scalingFactor, tmpRem, RadixBaseExponent)
                Else
                    orem = A
                End If
                If Len(orem) > sfx Then   'Check last word scalling
                    orem = Mid(orem, 1, Len(orem) - sfx)
                End If
            End If
            tmpResult = ZeroRemoval(tmpResult)
            If DivSignFlag Then
                tmpResult = Minus + tmpResult
            End If
            If DividendSign Then
                orem = Minus + orem
            End If
            BaseCaseDivRem = tmpResult
        End If
        'Debug.Print "Corrections: "; correctionsCount
    End If
End Function

''' <summary>
''' Calculate the difference between the given LargerNumber and SmallerNumber.
''' The possible borrow is returned in the decBorrow parameter.
''' </summary>
Private Function ChunkDif(ByRef LargerNumber As String, ByRef SmallerNumber As String, _
                            ByRef decBorrow As String, ByRef DecimalRadixExponent As Integer, _
                            ByRef OrdinalWord As Long, Optional OnlyCarry As Boolean = False) As String
    Dim ShunkA As String, ShunkB As String, LargerLen As Double, SmallerLen As Double
    
    LargerLen = Len(LargerNumber)
    SmallerLen = Len(SmallerNumber)
    If Not OnlyCarry Then
        ShunkA = Mid(LargerNumber, LargerLen - (DecimalRadixExponent * OrdinalWord) + 1, DecimalRadixExponent)
        ShunkB = Mid(SmallerNumber, SmallerLen - (DecimalRadixExponent * OrdinalWord) + 1, DecimalRadixExponent)
        ChunkDif = CStr(CDec(ShunkA) - CDec(ShunkB) + CDec(decBorrow))
        If Mid(ChunkDif, 1, 1) = Minus Then
            ChunkDif = CStr(CDec(ChunkDif) + CDec(strOne + String(DecimalRadixExponent, 48)))
            decBorrow = strNegativeOne
        Else
            decBorrow = strZero
        End If
        If Len(ChunkDif) < DecimalRadixExponent Then
            ChunkDif = PrefixZeros(ChunkDif, (DecimalRadixExponent - Len(ChunkDif)))
        End If
    Else
        ShunkA = Mid(LargerNumber, LargerLen - (DecimalRadixExponent * OrdinalWord) + 1, DecimalRadixExponent)
        ChunkDif = CStr(CDec(ShunkA) + CDec(decBorrow))
        If Mid(ChunkDif, 1, 1) = Minus Then
            ChunkDif = CStr(CDec(ChunkDif) + CDec(strOne + String(DecimalRadixExponent, 48)))
            decBorrow = strNegativeOne
        Else
            decBorrow = strZero
        End If
        If Len(ChunkDif) < DecimalRadixExponent Then
            ChunkDif = PrefixZeros(ChunkDif, (DecimalRadixExponent - Len(ChunkDif)))
        End If
    End If
End Function

''' <summary>
''' Calculate the sum between the given LargerNumber and SmallerNumber.
''' The possible carry is returned in the decCarry parameter.
''' </summary>
Private Function ChunkSum(ByRef LargerNumber As String, ByRef SmallerNumber As String, _
                            ByRef decCarry As String, ByRef DecimalRadixExponent As Integer, _
                            ByRef OrdinalWord As Long, Optional OnlyCarry As Boolean = False) As String
    Dim ShunkA As String, ShunkB As String, LargerLen As Double, SmallerLen As Double
    
    LargerLen = Len(LargerNumber)
    SmallerLen = Len(SmallerNumber)
    If Not OnlyCarry Then
        ShunkA = Mid(LargerNumber, LargerLen - (DecimalRadixExponent * OrdinalWord) + 1, DecimalRadixExponent)
        ShunkB = Mid(SmallerNumber, SmallerLen - (DecimalRadixExponent * OrdinalWord) + 1, DecimalRadixExponent)
        ChunkSum = CStr(CDec(ShunkA) + CDec(ShunkB) + CDec(decCarry))
        If Len(ChunkSum) > DecimalRadixExponent And LargerLen > (DecimalRadixExponent * OrdinalWord) Then
            decCarry = Mid(ChunkSum, 1, 1)
            ChunkSum = Mid(ChunkSum, 2)
        Else
            decCarry = strZero
        End If
        If Len(ChunkSum) < DecimalRadixExponent Then
            ChunkSum = PrefixZeros(ChunkSum, (DecimalRadixExponent - Len(ChunkSum)))
        End If
    Else
        ShunkA = Mid(LargerNumber, LargerLen - (DecimalRadixExponent * OrdinalWord) + 1, DecimalRadixExponent)
        ChunkSum = CStr(CDec(ShunkA) + CDec(decCarry))
        If Len(ChunkSum) > DecimalRadixExponent And LargerLen > (DecimalRadixExponent * OrdinalWord) Then
            decCarry = Mid(ChunkSum, 1, 1)
            ChunkSum = Mid(ChunkSum, 2)
        Else
            decCarry = strZero
        End If
        If Len(ChunkSum) < DecimalRadixExponent Then
            ChunkSum = PrefixZeros(ChunkSum, (DecimalRadixExponent - Len(ChunkSum)))
        End If
    End If
End Function

''' <summary>
''' Calculate the product between the given LargerNumber and SmallerNumber.
''' The possible carry is returned in the decCarry parameter.
''' </summary>
Private Function ChunkProduct(ByRef SmallerNumber As String, ByRef LargerNumber As String, ByRef decCarry As String, _
                    ByRef MaxWordAmount As Long, ByRef Radix As Integer, ByRef LargerOrdinalWord As Long, _
                    ByRef SmallerOrdinalWord As Long, ByRef ChunkA As String, ByRef ChunkB As String, ByVal SkipLNchunk As Boolean) As String
    If Not SkipLNchunk Then
        ChunkA = Mid(SmallerNumber, Len(SmallerNumber) - Radix * SmallerOrdinalWord + 1, Radix)
    End If
    ChunkB = Mid(LargerNumber, Len(LargerNumber) - Radix * LargerOrdinalWord + 1, Radix)
    ChunkProduct = CStr(CDec(ChunkA) * CDec(ChunkB) + CDec(decCarry))
    If Len(ChunkProduct) > Radix And LargerOrdinalWord < MaxWordAmount Then
        decCarry = Mid(ChunkProduct, 1, Len(ChunkProduct) - Radix)
        ChunkProduct = Mid(ChunkProduct, Len(ChunkProduct) - Radix + 1)
    Else
        decCarry = strZero
    End If
    If Len(ChunkProduct) < Radix Then
        ChunkProduct = PrefixZeros(ChunkProduct, Radix - Len(ChunkProduct))
    End If
End Function

''' <summary>
''' Obtains a representation, of the number in this instance,
''' using the given exponnet.
''' </summary>
Public Function Cohort(ByVal exponent As Long) As Float
    Dim Dif As Long
    Dim n As Long
    Dim tmpStr As String
    Dim IsDecPartZero As Boolean
    
    If c_Represented Then
        If c_Sign Then
            StrIntPart = Mid(c_SignificandINT, 2)
        Else
            StrIntPart = c_SignificandINT
        End If
        StrDecPart = c_SignificandDEC
        strExponent = c_Magnitude
        Dif = exponent - CLng(strExponent)
        If Dif > 0 Then
            StrDecNotation = strZero + strDot + String(Dif - Len(StrIntPart), 48) + StrIntPart + StrDecPart
        ElseIf Dif <= 0 Then
            If Abs(Dif) >= Len(StrDecPart) Then
                tmpStr = StrIntPart + StrDecPart + String(Abs(Dif) - Len(StrDecPart) + 1, 48)
            Else
                tmpStr = StrIntPart + StrDecPart
            End If
            StrDecNotation = ZeroRemoval(Mid(tmpStr, 1, Abs(Dif) + 1) + strDot + Mid(tmpStr, Abs(Dif) + 2))
        End If
        n = InStr(1, StrDecNotation, strDot, vbBinaryCompare)
        If n > 0 Then
            StrIntPart = Mid(StrDecNotation, 1, n - 1)
        Else
            StrIntPart = strZero
        End If
        StrDecPart = Mid(StrDecNotation, n + 1)
        If c_Sign Then
            c_SignificandINT = Minus & StrIntPart
        Else
            c_SignificandINT = StrIntPart
        End If
        If Zero(StrDecPart) Then StrDecPart = strZero: IsDecPartZero = True
        c_SignificandDEC = StrDecPart
        c_Magnitude = exponent
        If Not IsDecPartZero Then
            c_Significand = DecZeroRemoval(c_SignificandINT + strDot + c_SignificandDEC)
        Else
            c_Significand = c_SignificandINT + strDot + c_SignificandDEC
        End If
        If c_Significand = strZero Then c_Significand = strZero + strDot + strZero
        c_Representation = c_Significand + SciNotSymbol + c_Magnitude
    End If
    Set Cohort = Me
End Function

''' <summary>
''' Returns a String indicating the result of a Float comparison:
'''       "-1" -> float1 is less than float2
'''       "0" -> float1 is equal to float2
'''       "1" -> float1 is greater than float2
'''       vbNullString -> float1 or float2 isn't represented
''' </summary>
Public Function Compare(float1 As Float, float2 As Float) As String
    If float1.Represented And float2.Represented Then
        Dim Number1 As String
        Dim Number2 As String
        Dim altNumberRep() As String
        Dim tmpResult As String
        
        Number1 = GetCohortZeroSignificant(float1)
        Number2 = GetCohortZeroSignificant(float2)
        altNumberRep() = SplitDecimalNumbers(Number1, Number2)
        'Compare integers
        tmpResult = CompareHelper(altNumberRep(0), altNumberRep(1))
        'Compare decimal
        If tmpResult = strZero Then
            tmpResult = CompareHelper(altNumberRep(2), altNumberRep(3))
        End If
        Compare = tmpResult
    End If
End Function

''' <summary>
''' Returns a String indicating the result of a number comparison:
'''       "-1" -> number1 is less than number2
'''       "0" -> number1 is equal to number2
'''       "1" -> number1 is greater than number2
'''       vbNullString -> number1 or number2 is null string
''' </summary>
Private Function Compare2(ByVal Number1 As String, ByVal Number2 As String) As String
    If Number1 <> vbNullString And Number2 <> vbNullString Then
        Dim altNumberRep() As String
        Dim tmpResult As String
        
        altNumberRep() = SplitDecimalNumbers(Number1, Number2)
        'Compare integers
        tmpResult = CompareHelper(altNumberRep(0), altNumberRep(1))
        'Compare decimal
        If tmpResult = strZero Then
            tmpResult = CompareHelper(altNumberRep(2), altNumberRep(3))
        End If
        Compare2 = tmpResult
    End If
End Function

Private Function CompareHelper(ByVal Number1 As String, ByVal Number2 As String) As String
    If GT(Number1, Number2) Then
        CompareHelper = strOne
    Else
        If Number1 = Number2 Then
            CompareHelper = strZero
        Else
            CompareHelper = strNegativeOne
        End If
    End If
End Function

Private Sub ComputeFillCsiNotComp(ByVal StrNum As String)
    Dim DecimalPointPosition As Long
    Dim tmpLng As Long
    Dim i As Long
    Dim FSDP As Long
    
    DecimalPointPosition = InStr(1, StrNum, strDot, vbBinaryCompare)
    tmpLng = DecimalPointPosition
    If DecimalPointPosition = 0 Then 'There is no decimal point in the number
        StrIntPart = Mid(StrNum, 1, 1)
        If Len(StrNum) >= 2 Then
            StrDecPart = Mid(StrNum, 2)
        Else
            StrDecPart = strZero
        End If
        If StrDecPart <> strZero Then
            strExponent = CStr(Len(StrDecPart))
        Else
            strExponent = strZero
        End If
    Else 'There is a decimal point that can be shifted
        StrIntPart = ZeroRemoval(Mid(StrNum, 1, tmpLng - 1))
        If StrIntPart = "" Then StrIntPart = strZero
        StrDecPart = Mid(StrNum, tmpLng + 1)
        If StrDecPart = "" Then StrDecPart = strZero
        i = tmpLng - 2
        If StrIntPart = strZero Then
            FSDP = FirstSignificantDigitPosition(StrNum)
            StrIntPart = Mid(StrNum, FSDP, 1)
            StrDecPart = Mid(StrNum, FSDP + 1)
            If StrDecPart = "" Then StrDecPart = strZero
            If StrIntPart <> strZero Or StrDecPart <> strZero Then
                strExponent = CStr(2 - FSDP)
            Else
                strExponent = strZero
            End If
        Else
            If Len(StrNum) >= 2 Then
                If Len(StrIntPart) > 1 Then StrIntPart = Mid(StrIntPart, 1, 1)
                StrDecPart = Mid(StrNum, 2, i) + Mid(StrNum, tmpLng + 1)
            Else
                StrDecPart = strZero
            End If
            If StrIntPart <> strZero And StrDecPart <> strZero Then
                strExponent = CStr(i)
            Else
                strExponent = strZero
            End If
        End If
    End If
End Sub

''' <summary>
''' Represents and stores the given StrNum in SCI notation and
''' fills the current instance.
''' </summary>
Public Function Create(StrNum As String) As Float
    Dim BolSciNotation As Boolean
    Dim SignAndNumber() As Variant
    Dim IsDecPartZero As Boolean
    
    StrNum = UCase(StrNum)
    BolSciNotation = IsPresentSciNotSymbol(StrNum)
    If Validate(StrNum, BolSciNotation) Then
        c_ProperlyRepresented = True
        SignAndNumber = SignCut(StrNum)
        StrNum = SignAndNumber(1)
        GetSignValue (SignAndNumber(0))
        If Zero(StrNum) Then c_IsZero = True
        If BolSciNotation Then
            GetFillCsiNotComp (StrNum)
        Else
            If Not c_IsZero Then StrNum = DecZeroRemoval(ZeroRemoval(StrNum))
            ComputeFillCsiNotComp (StrNum)
        End If
        If c_Sign Then
            c_Value = Minus & StrNum
            c_SignificandINT = Minus & StrIntPart
        Else
            c_Value = StrNum
            c_SignificandINT = StrIntPart
        End If
        If Zero(StrDecPart) Then StrDecPart = strZero: IsDecPartZero = True
        c_SignificandDEC = StrDecPart
        c_Magnitude = strExponent
        If Not IsDecPartZero Then
            c_Significand = DecZeroRemoval(c_SignificandINT + strDot + c_SignificandDEC)
        Else
            c_Significand = c_SignificandINT + strDot + c_SignificandDEC
        End If
        If c_Significand = strZero Or c_Significand = Minus & strZero _
                Or c_Significand = Plus & strZero Then
                    c_Significand = strZero + strDot + strZero
        End If
        c_Representation = c_Significand + SciNotSymbol + c_Magnitude
        c_Represented = True
        c_ErrorDescription = GetErrorDescription(c_ErrorDescriptionNumber)
    Else
        c_Value = "#NUM!"
        c_SignificandINT = "#NUM!"
        c_SignificandDEC = "#NUM!"
        c_Represented = False
        c_ProperlyRepresented = False
        c_Significand = "#NUM!"
        c_Magnitude = "#NUM!"
        c_Representation = "#NUM!"
        c_ErrorDescription = GetErrorDescription(c_ErrorDescriptionNumber)
    End If
    Set Create = Me
End Function

Public Function Create2(StrNum As String) As Float
    Dim BolSciNotation As Boolean
    Dim SignAndNumber() As Variant
    Dim IsDecPartZero As Boolean
    
    On Error GoTo Represent_err_handler
    c_Value = StrNum
    StrNum = UCase(StrNum)
    BolSciNotation = IsPresentSciNotSymbol(StrNum)
    c_ProperlyRepresented = True
    SignAndNumber = SignCut(StrNum)
    StrNum = SignAndNumber(1)
    GetSignValue (SignAndNumber(0))
    If Zero(StrNum) Then c_IsZero = True
    If BolSciNotation Then
        GetFillCsiNotComp (StrNum)
    Else
        If Not c_IsZero Then StrNum = DecZeroRemoval(ZeroRemoval(StrNum))
        ComputeFillCsiNotComp (StrNum)
    End If
    If c_Sign Then
        c_Value = Minus & StrNum
        c_SignificandINT = Minus & StrIntPart
    Else
        c_Value = StrNum
        c_SignificandINT = StrIntPart
    End If
    If Zero(StrDecPart) Then StrDecPart = strZero: IsDecPartZero = True
    c_SignificandDEC = StrDecPart
    c_Magnitude = strExponent
    If Not IsDecPartZero Then
        c_Significand = DecZeroRemoval(c_SignificandINT + strDot + c_SignificandDEC)
    Else
        c_Significand = c_SignificandINT + strDot + c_SignificandDEC
    End If
    If c_Significand = strZero Or c_Significand = Minus & strZero _
                Or c_Significand = Plus & strZero Then
                    c_Significand = strZero + strDot + strZero
    End If
    c_Representation = c_Significand + SciNotSymbol + c_Magnitude
    c_Represented = True
    Set Create2 = Me
    Exit Function
Represent_err_handler:
    c_Value = "#NUM!"
    c_SignificandINT = "#NUM!"
    c_SignificandDEC = "#NUM!"
    c_Represented = False
    c_ProperlyRepresented = False
    c_Significand = "#NUM!"
    c_Magnitude = "#NUM!"
    c_Representation = "#NUM!"
    Set Create2 = Me
End Function

Private Function DecZeroRemoval(ByVal StrNum As String) As String
    Dim DecimalPointPosition As Long
    Dim LSP As Long
    Dim tmpStr As String
    
    LSP = LastSignificantPosition(StrNum)
    DecimalPointPosition = InStr(1, StrNum, strDot, vbBinaryCompare)
    If DecimalPointPosition > 0 Then
        tmpStr = Mid(StrNum, 1, LSP)
    Else
        tmpStr = StrNum
    End If
    DecZeroRemoval = tmpStr
End Function

''' <summary>
''' This method calculates the quotiend of A (stored in the current instance)
''' divided by B (passed as an string argument), using the given radix.
''' </summary>
Public Function DividedBy(Divisor As Variant, Optional Radix As Integer = 14) As Float
    Dim Bfloat As Float
    Dim A As String, B As String
    Dim ka As Long, kb As Long
    Dim mA As String, mB As String
    Dim k As String, intDiv As String
    Dim ldFactor As Long, lFactor As Long
    
    Set Bfloat = New Float
    If TypeName(Divisor) <> TypeName(Me) Then
        Bfloat.Create CStr(Divisor)
    Else
        Bfloat.Create Divisor.Representation
    End If
    'Integer significants
    A = Me.SignificandIntegerPart + Me.SignificandDecimalPart
    B = Bfloat.SignificandIntegerPart + Bfloat.SignificandDecimalPart
    If Me.Sign Then
        A = Minus + A
    End If
    If Bfloat.Sign Then
        B = Minus + B
    End If
    ldFactor = LenB(A) - LenB(B)
    If ldFactor < 0 Then
        lFactor = ldFactor - Len(Bfloat.SignificandIntegerPart) + 1
        B = Mid(B, 1, Len(B) + lFactor)
    End If
    'Expansion factors for A and B
    ka = Len(Me.SignificandDecimalPart)
    kb = Len(Bfloat.SignificandDecimalPart) + lFactor
    'Current magnitudes
    mA = Me.Magnitude
    mB = Bfloat.Magnitude
    'Scaling factor
    k = IntStrSum(IntStrSum(mA, Minus + mB), Minus + IntStrSum(CStr(ka), Minus + CStr(kb))) '(mA - mB) - (ka - kb)
    'Partial division
    intDiv = BaseCaseDivRem(A, B, c_LastRemainder, Radix)
    'Return
    Me.Create intDiv + SciNotSymbol + k
    Set DividedBy = Me
End Function

Private Function DividendValueProjection(qj As String, qj_old As String, Dividend_old As String, cFactor As String) As String
    DividendValueProjection = IntStrSum(IntKBMult(Minus + cFactor, CStr(CDec(qj) - CDec(qj_old))), Dividend_old)
End Function

Private Function FirstSignificantDigitPosition(ByRef StrNum As String) As Double
    Dim j As Long
    Dim n As Long
    Dim k As Long
    Dim DecimalPointPosition As Long
    
    DecimalPointPosition = InStr(1, StrNum, strDot, vbBinaryCompare)
    j = 0
    k = 0
    n = 1
    Do
        j = InStr(1, StrNum, CStr(n), vbBinaryCompare)
        If k = 0 Then
            k = j
        Else
            If j < k And j > 0 Then k = j
        End If
        n = n + 1
    Loop While n <= 9 And k <> 1
    j = k
    FirstSignificantDigitPosition = j
End Function

Private Function FirstSignificantPosition(ByRef StrNum As String) As Double
    Dim j As Long
    Dim n As Long
    Dim k As Long
    Dim DecimalPointPosition As Long
    
    DecimalPointPosition = InStr(1, StrNum, strDot, vbBinaryCompare)
    j = 0
    k = 0
    n = 1
    Do
        j = InStr(1, StrNum, CStr(n), vbBinaryCompare)
        If k = 0 Then
            k = j
        Else
            If j < k And j > 0 Then k = j
        End If
        n = n + 1
    Loop While n <= 9 And k <> 1
    j = k
    If j > DecimalPointPosition And DecimalPointPosition > 1 Then
        FirstSignificantPosition = DecimalPointPosition - 1
    ElseIf j >= 0 Then
        FirstSignificantPosition = j
    End If
End Function

Private Function GetCohortZeroSignificant(aFloat As Float) As String
    Dim tmpFloat As Float
    
    Set tmpFloat = New Float
    With tmpFloat
        GetCohortZeroSignificant = .Create(aFloat.Representation).Cohort(0).Significand
    End With
End Function

Private Function GetDivisionSign(ByRef FirstSign As Boolean, ByRef SecondSign As Boolean) As Boolean
        GetDivisionSign = GetProductSign(FirstSign, SecondSign)
End Function

Private Function GetErrorDescription(ErrorValue As enumErrorDescription) As String
    If ErrorValue = ICI Then
        GetErrorDescription = "Invalid Character Input."
    ElseIf ErrorValue = IPP Then
        GetErrorDescription = "Invalid Point Position."
    ElseIf ErrorValue = SISI Then
        GetErrorDescription = "Significand Invalid Sign Input."
    ElseIf ErrorValue = EMISI Then
        GetErrorDescription = "Exponent Magnitude Invalid Sign Input."
    ElseIf ErrorValue = EMINSI Then
        GetErrorDescription = "Exponent Magnitude Invalid Notation Symbol Input."
    ElseIf ErrorValue = SSICI Then
        GetErrorDescription = "Significand Sign Invalid Character Input."
    Else
        GetErrorDescription = "Nothing"
    End If
End Function

Private Sub GetFillCsiNotComp(ByVal StrNum As String)
    Dim DecimalPointPosition As Long
    Dim tmpLng As Long
    Dim tmpStr As String
    
    DecimalPointPosition = InStr(1, StrNum, strDot, vbBinaryCompare)
    tmpLng = InStr(1, StrNum, SciNotSymbol, vbBinaryCompare)
    If DecimalPointPosition > 1 Then
        StrIntPart = Mid(StrNum, 1, DecimalPointPosition - 1)
        StrDecPart = Mid(StrNum, DecimalPointPosition + 1, tmpLng - (DecimalPointPosition + 1))
    ElseIf DecimalPointPosition = 0 Then
        StrIntPart = Mid(StrNum, 1, tmpLng - 1)
        StrDecPart = strZero
    End If
    If Not Zero(StrIntPart & StrDecPart) Then
        strExponent = Mid(StrNum, tmpLng + 1)
        tmpLng = Len(StrIntPart)
        If tmpLng > 1 Then
            tmpStr = Mid(StrIntPart, 2)
            StrIntPart = Mid(StrIntPart, 1, 1)
            StrDecPart = tmpStr + StrDecPart
            If StrIntPart <> strZero And StrDecPart <> strZero Then
                strExponent = CStr(CLng(strExponent) + tmpLng - 1)
            Else
                strExponent = strZero
            End If
        ElseIf StrIntPart = strZero Then
            tmpLng = FirstSignificantDigitPosition(StrDecPart)
            StrIntPart = Mid(StrDecPart, tmpLng, 1)
            tmpStr = Mid(StrDecPart, tmpLng + 1)
            If tmpStr <> "" Then
                StrDecPart = tmpStr
            Else
                StrDecPart = strZero
            End If
            strExponent = CStr(CLng(strExponent) - tmpLng)
        End If
    Else
        strExponent = strZero
    End If
End Sub

Private Function GetNumberSign(ByRef Number As String) As Boolean
    Dim SignChar As String
    
    SignChar = Mid(ApplyLawOfSigns(Number), 1, 1)
    If SignChar = Plus Then
        Number = Mid(Number, 2)
        GetNumberSign = False
    Else
        If SignChar = Minus Then
            Number = Mid(Number, 2)
            GetNumberSign = True
        Else
            GetNumberSign = False
        End If
    End If
End Function

Private Function GetProductSign(ByRef FirstSign As Boolean, ByRef SecondSign As Boolean) As Boolean
        If FirstSign And SecondSign Then
            GetProductSign = False
        ElseIf FirstSign Or SecondSign Then
            GetProductSign = True
        End If
End Function

Private Sub GetSignValue(ByVal tmpStrSign As String)
    Dim tmpSign As String
    
    tmpSign = ApplyLawOfSigns(tmpStrSign)
    If tmpSign = Minus Then
        c_Sign = True
    ElseIf StrSign = Plus Then
        c_Sign = False
    End If
End Sub

Private Function GetSumSign(ByRef FirstNumber As String, ByRef LargerNumber As String, _
                            ByRef FirstNumberSign As Boolean, ByRef SecondNumberSign As Boolean) As Boolean
    If FirstNumberSign And SecondNumberSign Then
        GetSumSign = True
    ElseIf FirstNumberSign Or SecondNumberSign Then
        If FirstNumberSign And LargerNumber = FirstNumber Then
            GetSumSign = True
        ElseIf SecondNumberSign And LargerNumber <> FirstNumber Then
            GetSumSign = True
        Else
            GetSumSign = False
        End If
    End If
End Function

''' <summary>
''' Returns "True" if the first number is greater than the second.
''' </summary>
Private Function GT(ByVal StrNum1 As String, ByVal StrNum2 As String) As Boolean
    Dim Mayornumber As String, MinorNumber As String, iCounter As Long
    Dim Number1Sign As Boolean, Number2Sign As Boolean
    Dim n As Long
    
    StrNum1 = StrNum1: StrNum2 = StrNum2
    Number1Sign = (Mid(StrNum1, 1, 1) = Minus)
    Number2Sign = (Mid(StrNum2, 1, 1) = Minus)
    If Number1Sign Or Number2Sign Then
        If Number1Sign And Not Number2Sign Then
            Mayornumber = StrNum2: MinorNumber = StrNum1
        ElseIf Not Number1Sign And Number2Sign Then
            Mayornumber = StrNum1: MinorNumber = StrNum2
        ElseIf Number1Sign And Number2Sign Then
            StrNum1 = Mid(StrNum1, 2)
            StrNum2 = Mid(StrNum2, 2)
            StrNum1 = ZeroRemoval(StrNum1)
            StrNum2 = ZeroRemoval(StrNum2)
            If Len(StrNum1) > Len(StrNum2) Then
                MinorNumber = StrNum1: Mayornumber = StrNum2
            ElseIf Len(StrNum1) < Len(StrNum2) Then
                MinorNumber = StrNum2: Mayornumber = StrNum1
            Else
                n = Fix(Len(StrNum1) / 28)
                If 28 * n - Len(StrNum1) <> 0 Then
                    n = n + 1
                End If
                Do While iCounter <= n - 1 And _
                        Mid(StrNum1, 28 * iCounter + 1, 28) = Mid(StrNum2, 28 * iCounter + 1, 28)
                    iCounter = iCounter + 1
                Loop
                If iCounter < n Then
                    If CDec(Mid(StrNum1, 28 * iCounter + 1, 28)) > CDec(Mid(StrNum2, 28 * iCounter + 1, 28)) Then
                        MinorNumber = StrNum1: Mayornumber = StrNum2
                    Else
                        MinorNumber = StrNum2: Mayornumber = StrNum1
                    End If
                End If
            End If
        End If
        If Mayornumber = StrNum1 And Mayornumber <> MinorNumber Then
            GT = True
        Else
            GT = False
        End If
    Else
        GT = AbsGT(StrNum1, StrNum2)
    End If
End Function

''' <summary>
''' This source method computes the product (A x B) of two big integers A and B
''' by Karatsuba multiplication until words have size MSwitch.
''' When words have less digits than MSwitch the standar multiplication is used.
''' </summary>
Private Function IntKBMult(ByVal FirstNumber As String, ByVal SecondNumber As String, Optional MSwitch As Integer = 512) As String
    Dim StrNumberA As String, StrNumberB As String, k As Long
    Dim n As Long, StrC0 As String, StrC1 As String, StrC2 As String
    Dim FirstNumberSign As Boolean, SecondNumberSign As Boolean
    Dim ProductSignFlag As Boolean, tmpStr As String, A(0 To 1) As String
    Dim B(0 To 1) As String, FirsTMPDigitsCount As Long
    Dim SecondTMPDigitsCount As Long, BalanceLenPwr As Long
    Dim BalanceLen As Long, StrIntermed As String
    Dim TmpStrNumberA As String, TmpStrNumberB As String, c As Long
    Dim WordsCount As Long, LenNumberA As Long, LenNumberB As Long

    If Not IsNumeric(MSwitch) Or MSwitch <= 0 Or MSwitch > 512 Then
        MSwitch = 512
    Else 'Requires a power of 2 less than 512
        Dim tmpWordLen As Double
        Dim fWlen As Long
        
        tmpWordLen = Log(MSwitch) / Log(2)
        fWlen = Fix(tmpWordLen)
        If fWlen - tmpWordLen <> 0 Then
            tmpWordLen = fWlen + 1
        End If
        MSwitch = CInt(2 ^ tmpWordLen)
    End If
    FirstNumberSign = GetNumberSign(FirstNumber)
    SecondNumberSign = GetNumberSign(SecondNumber)
    If AbsGT(FirstNumber, SecondNumber) Then
        WordsCount = WordCount(FirstNumber, MSwitch)
        StrNumberA = FirstNumber: StrNumberB = SecondNumber
    Else
        WordsCount = WordCount(SecondNumber, MSwitch)
        StrNumberA = SecondNumber: StrNumberB = FirstNumber
    End If
    LenNumberA = Len(StrNumberA)
    LenNumberB = Len(StrNumberB)
    StrNumberA = PrefixZeros(StrNumberA, WordsCount * MSwitch - LenNumberA)
    StrNumberB = PrefixZeros(StrNumberB, WordsCount * MSwitch - LenNumberB)
    ProductSignFlag = GetProductSign(FirstNumberSign, SecondNumberSign)
    BalanceLenPwr = NearBase2Exp(StrNumberA, MSwitch)
    BalanceLen = 2 ^ BalanceLenPwr
    n = BalanceLen
    If n = 1 Then
        If Not ProductSignFlag Then
            IntKBMult = IntStrMult(StrNumberA, StrNumberB)
        Else
            IntKBMult = IntStrMult(Minus & StrNumberA, StrNumberB)
        End If
    ElseIf Zero(StrNumberA) Or Zero(StrNumberB) Then
        IntKBMult = strZero
    Else
        k = n * MSwitch / 2
        StrNumberA = PrefixZeros(StrNumberA, n * MSwitch - Len(StrNumberA))
        StrNumberB = PrefixZeros(StrNumberB, n * MSwitch - Len(StrNumberB))
        LenNumberA = Len(StrNumberA)
        LenNumberB = Len(StrNumberB)
        For c = 1 To 2
            If c < 2 Then
                TmpStrNumberA = Mid(StrNumberA, LenNumberA - k + 1)
                TmpStrNumberB = Mid(StrNumberB, LenNumberB - k + 1)
            Else
                TmpStrNumberA = StrNumberA
                TmpStrNumberB = StrNumberB
            End If
            A(c - 1) = TmpStrNumberA
            B(c - 1) = TmpStrNumberB
            FirsTMPDigitsCount = Len(TmpStrNumberA)
            SecondTMPDigitsCount = Len(TmpStrNumberB)
            StrNumberA = Mid(StrNumberA, 1, LenNumberA - FirsTMPDigitsCount)
            StrNumberB = Mid(StrNumberB, 1, LenNumberB - SecondTMPDigitsCount)
        Next c
        StrC0 = IntKBMult(A(0), B(0), MSwitch)
        StrC1 = IntKBMult(A(1), B(1), MSwitch)
        StrC2 = IntKBMult(IntStrSum(A(1), Minus & A(0)), IntStrSum(B(1), Minus & B(0)), MSwitch)
        StrIntermed = IntStrSum(IntStrSum(StrC0, StrC1), Minus & StrC2) + String(k, 48)
        StrC1 = StrC1 + String(2 * k, 48)
        tmpStr = IntStrSum(IntStrSum(StrC0, StrIntermed), StrC1)
        If ProductSignFlag Then
            IntKBMult = Minus + tmpStr
        Else
            IntKBMult = tmpStr
        End If
    End If
End Function

''' <summary>
''' This source method computes the product (A x B) of two big integers A and B
''' by multiplying chunks of size RadixBaseExponent selected so that the radix(base)
''' is B=10^RadixBaseExponent. The length of the numbers entered must be a multiple
''' of the radix exponent. The default value is RadixBaseExponent=14
''' </summary>
Private Function IntStrMult(ByVal FirstNumber As String, ByVal SecondNumber As String, Optional RadixBaseExponent As Integer = 14) As String
    Dim ChunkA As String, ChunkB As String
    Dim LessWordCount As Long, MaxWordCount As Long
    Dim LargerNumber As String, SmallerNumber As String, SortedNumbers() As String
    Dim iCounter As Long, jCounter As Long, decCarry As String
    Dim Product As String, LargerDigitsCount As Long, SmallerDigitsCount As Long
    Dim tmpStr As String, FirstNumberSign As Boolean
    Dim SecondNumberSign As Boolean, ProductSignFlag As Boolean
    Dim ChunkOpIt As Long
    
    If Not IsNumeric(RadixBaseExponent) Or RadixBaseExponent > 14 Or RadixBaseExponent < 0 Then RadixBaseExponent = 14
    FirstNumberSign = GetNumberSign(FirstNumber)
    SecondNumberSign = GetNumberSign(SecondNumber)
    If Zero(FirstNumber) Or Zero(SecondNumber) Then
        IntStrMult = strZero
    Else
        FirstNumber = ZeroRemoval(FirstNumber)
        SecondNumber = ZeroRemoval(SecondNumber)
        ProductSignFlag = GetProductSign(FirstNumberSign, SecondNumberSign)
        SortedNumbers = SortNumbersABSVal(FirstNumber, SecondNumber)
        LargerNumber = SortedNumbers(0)
        SmallerNumber = SortedNumbers(1)
        MaxWordCount = WordCount(LargerNumber, RadixBaseExponent)
        LessWordCount = WordCount(SmallerNumber, RadixBaseExponent)
        LargerNumber = PrefixZeros(LargerNumber, RadixBaseExponent * MaxWordCount - Len(LargerNumber))
        SmallerNumber = PrefixZeros(SmallerNumber, RadixBaseExponent * LessWordCount - Len(SmallerNumber))
        LargerDigitsCount = Len(LargerNumber)
        SmallerDigitsCount = Len(SmallerNumber)
        Product = strZero
        decCarry = strZero
        For iCounter = 1 To LessWordCount
            ChunkOpIt = iCounter
            For jCounter = 1 To MaxWordCount
                tmpStr = ChunkProduct(SmallerNumber, LargerNumber, _
                                        decCarry, MaxWordCount, _
                                        RadixBaseExponent, jCounter, iCounter, _
                                        ChunkA, ChunkB, ChunkOpIt <> iCounter) + tmpStr
                ChunkOpIt = ChunkOpIt + 1
            Next jCounter
            tmpStr = SufixZeros(tmpStr, RadixBaseExponent * (iCounter - 1))
            Product = IntStrSum(Product, tmpStr)
            tmpStr = vbNullString
        Next iCounter
        If ProductSignFlag Then
            IntStrMult = Minus + ZeroRemoval(Product)
        Else
            IntStrMult = ZeroRemoval(Product)
        End If
    End If
End Function

''' <summary>
''' This source method computes the sum (A + B) of two big integers A and B
''' by adding chunks of size RadixBaseExponent selected so that the radix(base)
''' is B=10^RadixBaseExponent. The length of the numbers entered must be a
''' multiple of the radix exponent. The default value is RadixBaseExponent=28
''' </summary>
Private Function IntStrSum(ByVal FirstNumber As String, ByVal SecondNumber As String, Optional RadixBaseExponent As Integer = 28) As String
    Dim LargerNumber As String, SmallerNumber As String
    Dim SortedNumbers() As String, LessWordCount As Long
    Dim SmallerDigitsCount As Long, MaxWordCount As Long
    Dim LargerDigitsCount As Long
    Dim decCarry As String, iCounter As Long
    Dim tmpStr As String, SignFlag As Boolean
    Dim boolDifference As Boolean, tmpLng As Long
    Dim FirstNumberSign As Boolean, SecondNumberSign As Boolean
    
    FirstNumberSign = GetNumberSign(FirstNumber)
    SecondNumberSign = GetNumberSign(SecondNumber)
    If SumExecutionNeeded(FirstNumber, SecondNumber) Then
        If Not IsNumeric(RadixBaseExponent) Or RadixBaseExponent > 28 Or RadixBaseExponent < 0 Then RadixBaseExponent = 28
        FirstNumber = ZeroRemoval(FirstNumber)
        SecondNumber = ZeroRemoval(SecondNumber)
        SortedNumbers = SortNumbersABSVal(FirstNumber, SecondNumber)
        LargerNumber = SortedNumbers(0)
        SmallerNumber = SortedNumbers(1)
        boolDifference = IsDifference(FirstNumberSign, SecondNumberSign)
        SignFlag = GetSumSign(FirstNumber, LargerNumber, FirstNumberSign, SecondNumberSign)
        LargerDigitsCount = Len(LargerNumber)
        SmallerDigitsCount = Len(SmallerNumber)
        LessWordCount = WordCount(SmallerNumber, RadixBaseExponent)
        MaxWordCount = WordCount(LargerNumber, RadixBaseExponent)
        LargerNumber = PrefixZeros(LargerNumber, RadixBaseExponent * MaxWordCount - LargerDigitsCount)
        SmallerNumber = PrefixZeros(SmallerNumber, RadixBaseExponent * LessWordCount - SmallerDigitsCount)
        decCarry = strZero
        If boolDifference Then
            For iCounter = 1 To LessWordCount
                tmpStr = ChunkDif(LargerNumber, SmallerNumber, decCarry, RadixBaseExponent, iCounter) + tmpStr
            Next iCounter
            Do Until decCarry = strZero
                tmpStr = ChunkDif(LargerNumber, SmallerNumber, decCarry, RadixBaseExponent, iCounter, True) + tmpStr
                iCounter = iCounter + 1
            Loop
        Else
            For iCounter = 1 To LessWordCount
                tmpStr = ChunkSum(LargerNumber, SmallerNumber, decCarry, RadixBaseExponent, iCounter) + tmpStr
            Next iCounter
            Do Until decCarry = strZero
                tmpStr = ChunkSum(LargerNumber, SmallerNumber, decCarry, RadixBaseExponent, iCounter, True) + tmpStr
                iCounter = iCounter + 1
            Loop
        End If
        tmpLng = RadixBaseExponent * MaxWordCount
        If tmpLng > (RadixBaseExponent * (iCounter - 1)) Then
            tmpStr = Mid(LargerNumber, 1, tmpLng - (RadixBaseExponent * (iCounter - 1))) + tmpStr
        End If
        tmpStr = ZeroRemoval(tmpStr)
        If SignFlag Then
            If tmpStr = strZero Then
                IntStrSum = tmpStr
            Else
                IntStrSum = Minus + tmpStr
            End If
        Else
            IntStrSum = tmpStr
        End If
    Else
        IntStrSum = SumOutPutWOExecution(FirstNumber, FirstNumberSign, SecondNumber, SecondNumberSign)
    End If
End Function

Private Function IsDifference(ByRef FirstNumberSign As Boolean, ByRef SecondNumberSign As Boolean) As Boolean
    If FirstNumberSign And SecondNumberSign Then
        IsDifference = False
    ElseIf FirstNumberSign Or SecondNumberSign Then
        IsDifference = True
    End If
End Function

Private Function IsEven(ByRef strNumber As String) As Boolean
    IsEven = Not CBool(CLng(Mid$(strNumber, Len(strNumber), 1)) And 1)
End Function

Private Function IsPresentSciNotSymbol(ByVal StrNum As String) As Boolean
    If InStr(1, StrNum, SciNotSymbol, vbBinaryCompare) > 0 Then IsPresentSciNotSymbol = True
End Function

Private Function LastSignificantPosition(ByRef StrNum As String) As Double
    Dim NumLen As Long
    Dim j As Long
    Dim n As Long
    Dim k As Long
    Dim DecimalPointPosition As Long
    
    NumLen = Len(StrNum)
    DecimalPointPosition = InStr(1, StrNum, ".", vbBinaryCompare)
    j = 0
    k = 0
    n = 1
    Do
        j = InStrRev(StrNum, CStr(n), -1, vbBinaryCompare)
        If k = 0 Then
            k = j
        Else
            If j > k And j > 0 Then k = j
        End If
        n = n + 1
    Loop While n <= 9 And k <> NumLen
    j = k
    If j < DecimalPointPosition And DecimalPointPosition > 1 Then
        LastSignificantPosition = DecimalPointPosition - 1
    ElseIf j >= 0 Then
        LastSignificantPosition = j
    End If
End Function

Private Function NearBase2Exp(ByRef Number As String, ByRef Radix As Integer) As Long
    Dim tmpResult As Double
    
    tmpResult = Log(Len(Number) / Radix) / Log(2)
    If tmpResult - Fix(tmpResult) <> 0 Then
        tmpResult = Fix(tmpResult) + 1
    End If
    NearBase2Exp = tmpResult
End Function

''' <summary>
''' This method calculates the power of A (stored in the current instance)
''' raised to the integer exponent B (passed as an string argument).
''' </summary>
Public Function Pow(exponent As String) As Float
    Dim A As String, B As String
    Dim ka As Long
    Dim mA As String
    Dim k As String
    Dim IntPow As String
    
    If InStrB(1, exponent, strDot) Then Exit Function  'Fractional exponents are not allowed
    'Integer significants
    A = Me.SignificandIntegerPart + Me.SignificandDecimalPart
    B = exponent
    'Expansion factors for A and B
    ka = Len(Me.SignificandDecimalPart)
    'Current magnitude
    mA = Me.Magnitude
    'Scaling factor
    k = IntStrSum(IntKBMult(mA, B), Minus + IntKBMult(CStr(ka), B))
    'Partial division
    IntPow = Power(A, B)
    If Me.Sign Then
        If Not IsEven(B) Then
            IntPow = Minus + IntPow
        End If
    End If
    'Return
    Me.Create IntPow + SciNotSymbol + k
    Set Pow = Me
End Function

Private Function Power(ByVal strNumber As String, ByVal strExponent As String) As String
    Dim aRem As String
    
    If strExponent = strZero Then
        Power = strOne
    Else
        If strExponent = strOne Then
            Power = strNumber
        Else
            If IsEven(strExponent) Then
                Power = Power(IntKBMult(strNumber, strNumber), BaseCaseDivRem(strExponent, strTwo, aRem)) 'Power (x*x, n/2)
            Else
                Power = IntKBMult(Power(IntKBMult(strNumber, strNumber), BaseCaseDivRem(strExponent, strTwo, aRem)), strNumber) 'Power (x*x, n/2)*x
            End If
        End If
    End If
End Function

Private Function PrefixZeros(ByVal Number As String, ByVal Amount As Long) As String
    PrefixZeros = String(Amount, 48) + Number
End Function

''' <summary>
''' Finds a qj satisfying A>0 for integer division using a modified
''' bisection method in the interval La <= qj <= Lb
''' </summary>
''' <param name="La">Leftmost interval value.</param>
''' <param name="Lb">Rightmost interval value.</param>
''' <param name="qj">Current value for qj.</param>
''' <param name="Dividend">Current value for the dividend (A).</param>
''' <param name="cFactor">Current value for B*10^(radix*j).</param>
Private Function QuotientProjection(ByVal Lba As String, ByVal Lbb As String, qj As String, _
                        Dividend As String, cFactor As String) As String
    Dim La As Variant, Lb As Variant
    Dim Lc As Variant                      'Intermediate point |La---:Lc:---Lb|
    Dim evalCounter As Long
    Dim fa As String, fb As String
    Dim fc As String
    Dim k As Long
    Dim segmentLen As Variant
    Dim toleranceFlag As Boolean
    Dim epsilon As Variant
    Dim tolerance As Variant
    Dim tCheck As Variant
    
    On Error GoTo QuotientProjection_errHandler
    La = CDec(Lba): Lb = CDec(Lbb)
    If La >= Lb Then Exit Function
    k = 0
    epsilon = CDec(1) 'Project a quotient such that 1 <= La - Lb <= 2
    tolerance = CDec(2 * epsilon)
    'Evaluate the function at La, Lb
    fa = DividendValueProjection(CStr(La), qj, Dividend, cFactor): evalCounter = evalCounter + 1
    fb = DividendValueProjection(CStr(Lb), qj, Dividend, cFactor): evalCounter = evalCounter + 1
    Do
        'Divide the interval into three segments, assumes La < Lb
        segmentLen = CDec(Fix((Lb - La) / 2))
        Lc = CDec(La + segmentLen)
        'Evaluate the function at Lc
        fc = DividendValueProjection(CStr(Lc), qj, Dividend, cFactor): evalCounter = evalCounter + 1
        If fc = strZero Then
            Exit Do
        End If
        If Compare2(IntKBMult(fa, fc), strZero) = strNegativeOne Then 'Root in interval La-Lc
            Lb = Lc
            fb = fc
        ElseIf Compare2(IntKBMult(fc, fb), strZero) = strNegativeOne Then
            La = Lc
            fa = fc
        Else        'The interval does not satisfy the condition fa*fb<0.
            GoTo QuotientProjection_terminate
        End If
        tCheck = Lb - La
        toleranceFlag = (tCheck > tolerance)
        k = k + 1
    Loop While toleranceFlag And k < 155
    Select Case tCheck
        Case 1  'There is no integers betwen La < qj < Lb
            QuotientProjection = CStr(Lb)
        Case Else 'Return the integer in the interval La < qj < Lb
            QuotientProjection = CStr(La + 1)
    End Select
    'QuotientProjection = Ld
QuotientProjection_terminate:
    Exit Function
QuotientProjection_errHandler:
End Function

Private Function SignCut(ByRef StrNum As String) As Variant
    Dim leadinZeros As Boolean
    Dim cIdx As Long
    Dim tmpLng As Long
    Dim FSP As Long
    Dim tmpStr As String
    
    ReDim tmpArray(1)
    FSP = FirstSignificantPosition(StrNum)
    tmpLng = InStr(1, StrNum, strZero, vbBinaryCompare)
    leadinZeros = (FSP > tmpLng And tmpLng > 0)
    If Not leadinZeros Then
        If FSP >= 1 Then
            tmpArray(0) = Mid(StrNum, 1, FSP - 1)
        Else
            tmpArray(0) = Mid(StrNum, 1, tmpLng - 1)
        End If
    Else
        tmpStr = Mid(StrNum, 1, FSP - 1)
        Do While Mid(tmpStr, cIdx + 1, 1) Like "[+-]"
            cIdx = cIdx + 1
        Loop
        tmpStr = Mid(tmpStr, 1, cIdx)  'Skip zeros
        If tmpStr Like "*[+-]*" Then
            tmpArray(0) = tmpStr
        End If
    End If
    If FSP > 0 Then
        tmpArray(1) = Mid(StrNum, FSP)
    ElseIf FSP = 0 Then
        tmpArray(1) = Mid(StrNum, tmpLng)
    End If
    SignCut = tmpArray
    Erase tmpArray
End Function

Private Sub SignSubstitution(ByRef expression As String, FindStr As String, Replacement As String)
    If InStrB(1, expression, FindStr) Then
        expression = Replace(expression, FindStr, Replacement)
    End If
End Sub

Private Function SortNumbersABSVal(ByRef FirstNumber As String, SecondNumber As String) As Variant
    Dim tmpArray(0 To 1) As String
    If AbsGT(FirstNumber, SecondNumber) Then
        tmpArray(0) = FirstNumber
        tmpArray(1) = SecondNumber
    Else
        tmpArray(1) = FirstNumber
        tmpArray(0) = SecondNumber
    End If
    SortNumbersABSVal = tmpArray
End Function

''' <summary>
''' Returns an array with 4 elements as follows:
'''       Index 0 -> Number1 integer part
'''       Index 1 -> Number2 integer part
'''       Index 2 -> Number1 decimal part prefixed with "1"
'''       Index 3 -> Number2 decimal part prefixed with "1"
''' This function is used in numerical comparisons (Compare & Compare2).
''' </summary>
Private Function SplitDecimalNumbers(ByVal Number1 As String, ByVal Number2 As String) As String()
    Dim sArray1() As String
    Dim sArray2() As String
    Dim LenDec1 As Long
    Dim LenDec2 As Long
    Dim tmpResult(0 To 3) As String
    
    sArray1() = Split(Number1, strDot)
    sArray2() = Split(Number2, strDot)
    tmpResult(0) = sArray1(LBound(sArray1))
    tmpResult(1) = sArray2(LBound(sArray2))
    If UBound(sArray1) - LBound(sArray1) > 0 Then
        tmpResult(2) = strOne + sArray1(LBound(sArray1) + 1)
        LenDec1 = Len(tmpResult(2))
    Else
        tmpResult(2) = strOne
    End If
    If UBound(sArray2) - LBound(sArray2) > 0 Then
        tmpResult(3) = strOne + sArray2(LBound(sArray2) + 1)
        LenDec2 = Len(tmpResult(3))
    Else
        tmpResult(3) = strOne
    End If
    If LenDec1 <> LenDec2 Then  'Equalize decimal length
        If LenDec1 > LenDec2 Then
            tmpResult(3) = SufixZeros(tmpResult(3), LenDec1 - LenDec2)
        Else
            tmpResult(2) = SufixZeros(tmpResult(2), LenDec2 - LenDec1)
        End If
    End If
    '@-----------------------------------------------------------------
    ' Set the sign to decimal
    If InStrB(1, tmpResult(0), Minus) Then
        tmpResult(2) = Minus + tmpResult(2)
    End If
    If InStrB(1, tmpResult(1), Minus) Then
        tmpResult(3) = Minus + tmpResult(3)
    End If
    SplitDecimalNumbers = tmpResult
End Function

Private Function SufixZeros(ByVal Number As String, ByVal Amount As Long) As String
    SufixZeros = Number + String(Amount, 48)
End Function

''' <summary>
''' This method calculates the sum of the number stored in the current instance (A)
''' and the number passed as an argument (B), using the given radix.
''' </summary>
Public Function Sum(summand As Variant, Optional Radix As Integer = 28) As Float
    Dim Bfloat As Float
    Dim Aint As String, Adec As String
    Dim Bint As String, Bdec As String
    Dim decCarry As String
    Dim ka As Long, kb As Long
    Dim k As Long, kDecSum As Long
    Dim intSum As String, decSum As String
    
    Set Bfloat = New Float
    If TypeName(summand) <> TypeName(Me) Then
        Bfloat.Create CStr(summand)
    Else
        Bfloat.Create summand.Representation
    End If
    Bfloat.Cohort CLng(Me.Magnitude) 'Equalize A and B exponents by A cohort
    'Calculate the expansion factor for A and B
    ka = Len(Me.SignificandDecimalPart)
    kb = Len(Bfloat.SignificandDecimalPart)
    If ka > kb Then
        k = ka
    Else
        k = kb
    End If
    'Significants
    Aint = Me.SignificandIntegerPart
    If Not Me.Sign Then
        Adec = SufixZeros(Me.SignificandDecimalPart, k - ka)
    Else
        Adec = Minus + SufixZeros(Me.SignificandDecimalPart, k - ka)
    End If
    Bint = Bfloat.SignificandIntegerPart
    If Not Bfloat.Sign Then
        Bdec = SufixZeros(Bfloat.SignificandDecimalPart, k - kb)
    Else
        Bdec = Minus + SufixZeros(Bfloat.SignificandDecimalPart, k - kb)
    End If
    'Partial sums
    intSum = IntStrSum(Aint, Bint, Radix)
    decSum = IntStrSum(Adec, Bdec, Radix)
    'Correction
    kDecSum = Len(decSum)
    If kDecSum > k Then
        decCarry = Mid(decSum, 1, kDecSum - k)
        decSum = Mid(decSum, kDecSum - k + 1)
        intSum = IntStrSum(intSum, decCarry, Radix)
    End If
    'Return
    Me.Create intSum + strDot + decSum + SciNotSymbol + Me.Magnitude
    Set Sum = Me
End Function

Private Function SumExecutionNeeded(ByVal FirstNumber As String, ByVal SecondNumber As String) As Boolean
    SumFstOperandZero = Zero(FirstNumber)
    SumSndOperandZero = Zero(SecondNumber)
    SumExecutionNeeded = Not (SumFstOperandZero Or SumSndOperandZero)
End Function

Private Function SumOutPutWOExecution(ByRef FirstNumber As String, ByRef FirstNumberSign As Boolean _
                                        , ByRef SecondNumber As String, SecondtNumberSign As Boolean) As String
    If SumFstOperandZero And SumSndOperandZero Then
        SumOutPutWOExecution = strZero
    ElseIf SumFstOperandZero Then
        If SecondtNumberSign Then
            SumOutPutWOExecution = Minus + ZeroRemoval(SecondNumber)
        Else
            SumOutPutWOExecution = ZeroRemoval(SecondNumber)
        End If
    ElseIf SumSndOperandZero Then
         If FirstNumberSign Then
            SumOutPutWOExecution = Minus + ZeroRemoval(FirstNumber)
        Else
            SumOutPutWOExecution = ZeroRemoval(FirstNumber)
        End If
    End If
End Function

''' <summary>
''' This method calculates the product of the number stored in the current instance (A)
''' and the number passed as an argument (B), using the given WordSwitch.
''' </summary>
Public Function Times(Operand As Variant, Optional WordSwitch As Integer = 512) As Float
    Dim Bfloat As Float
    Dim A As String, B As String
    Dim ka As Long, kb As Long
    Dim mA As String, mB As String
    Dim k As String, intProd As String
    
    Set Bfloat = New Float
    If TypeName(Operand) <> TypeName(Me) Then
        Bfloat.Create CStr(Operand)
    Else
        Bfloat.Create Operand.Representation
    End If
    'Integer significants
    A = Me.SignificandIntegerPart + Me.SignificandDecimalPart
    B = Bfloat.SignificandIntegerPart + Bfloat.SignificandDecimalPart
    If Me.Sign Then
        A = Minus + A
    End If
    If Bfloat.Sign Then
        B = Minus + B
    End If
    'Expansion factors for A and B
    ka = Len(Me.SignificandDecimalPart)
    kb = Len(Bfloat.SignificandDecimalPart)
    'Current magnitudes
    mA = Me.Magnitude
    mB = Bfloat.Magnitude
    'Scaling factor
    k = IntStrSum(IntStrSum(mA, mB), Minus + IntStrSum(CStr(ka), CStr(kb))) '(mA + mB) - (ka + kb)
    'Partial product
    intProd = IntKBMult(A, B, WordSwitch)
    'Return
    Me.Create intProd + SciNotSymbol + k
    Set Times = Me
End Function

Private Function Validate(ByVal StrNum As String, SciNotation As Boolean) As Boolean
    Dim CPoint1 As Long
    Dim i As Long
    Dim n As Long
    Dim SignChar As String
    Dim tmpLng As Long
    Dim tmpStr As String
    Dim tmpVariant As Variant
    
    Validate = True
    i = 1
    If Mid(StrNum, 1, 1) = strDot Then Validate = False: c_ErrorDescriptionNumber = IPP
    Do While Validate = True And i <= 255
        tmpStr = Chr(i)
        n = InStr(1, StrNum, tmpStr, vbBinaryCompare)
        If n > 0 And (i < 48 Or i > 57) Then
            If tmpStr <> Minus And tmpStr <> Plus And tmpStr <> strDot And tmpStr <> SciNotSymbol Then
                Validate = False
                c_ErrorDescriptionNumber = ICI
            ElseIf tmpStr = strDot Then
                CPoint1 = n
                n = InStr(n + 1, StrNum, strDot, vbBinaryCompare)
                If CPoint1 < n Then
                    Validate = False
                    c_ErrorDescriptionNumber = IPP
                End If
            ElseIf tmpStr = SciNotSymbol Then
                tmpLng = n
                If n < InStr(n + 1, StrNum, SciNotSymbol, vbBinaryCompare) Then
                    Validate = False
                    c_ErrorDescriptionNumber = EMINSI
                End If
            End If
        End If
        i = i + 1
    Loop
    If Validate Then
        n = FirstSignificantPosition(StrNum)
        If SciNotation Then
            If Abs(CPoint1 - tmpLng) < 2 Or CPoint1 > tmpLng Then
                Validate = False
                c_ErrorDescriptionNumber = IPP
            End If
            If Validate Then
                If n > 0 Then
                    tmpStr = Mid(StrNum, n, tmpLng - n)
                Else
                    tmpStr = Mid(StrNum, 1, tmpLng)
                End If
                If InStr(1, tmpStr, Plus, vbBinaryCompare) <> 0 Or InStr(1, tmpStr, Minus, vbBinaryCompare) <> 0 Then
                    Validate = False
                    c_ErrorDescriptionNumber = SISI
                End If
            End If
            If Validate Then
                tmpStr = Mid(StrNum, tmpLng + 1)
                If InStr(2, tmpStr, Plus, vbBinaryCompare) <> 0 Or InStr(2, tmpStr, Minus, vbBinaryCompare) <> 0 Then
                    Validate = False
                    c_ErrorDescriptionNumber = EMISI
                ElseIf tmpStr = Plus Or tmpStr = Minus Then
                    Validate = False
                    c_ErrorDescriptionNumber = EMISI
                ElseIf tmpStr = "" Then
                    Validate = False
                    c_ErrorDescriptionNumber = EMISI
                End If
            End If
        Else
            tmpVariant = SignCut(StrNum)
            tmpStr = tmpVariant(0)
            i = 1
            tmpLng = Len(tmpStr)
            Do While Validate = True And i <= tmpLng
                SignChar = Mid(tmpStr, i, 1)
                If SignChar <> Minus And SignChar <> Plus Then
                    Validate = False
                    c_ErrorDescriptionNumber = SSICI
                End If
                i = i + 1
            Loop
            If Validate Then
                tmpStr = tmpVariant(1)
                If InStr(1, tmpStr, Minus, vbBinaryCompare) <> 0 Or InStr(1, tmpStr, Plus, vbBinaryCompare) <> 0 Then
                    Validate = False
                    c_ErrorDescriptionNumber = SISI
                End If
            End If
        End If
    End If
End Function

Private Function WordCount(ByRef Number As String, ByRef Radix As Integer) As Long
    Dim tmpResult As Double
    Dim tmpRem As Double
    
    tmpResult = Len(Number) / Radix
    tmpRem = Fix(tmpResult)
    If tmpResult - tmpRem <> 0 Then
        tmpResult = tmpRem + 1
    End If
    WordCount = CLng(tmpResult)
End Function

Private Function Zero(ByVal StrNum As String) As Boolean
    Dim j As Long
    Dim n As Long
    Dim k As Long
    Dim BoolNotZero As Boolean
    
    n = 1
    Do
        j = InStr(1, StrNum, CStr(n), vbBinaryCompare)
        BoolNotZero = (j <> 0)
        n = n + 1
    Loop While n <= 9 And Not BoolNotZero
    k = InStr(1, StrNum, SciNotSymbol, vbBinaryCompare)
    If k > 0 Then
        BoolNotZero = (k > j)
    End If
    Zero = Not BoolNotZero
End Function

Private Function ZeroRemoval(ByVal StrNum As String) As String
    Dim FSP As Long
    Dim tmpStr As String
    
    FSP = FirstSignificantPosition(StrNum)
    If FSP = 0 Then
        tmpStr = strZero
    Else
        tmpStr = Mid(StrNum, FSP)
    End If
    ZeroRemoval = tmpStr
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
