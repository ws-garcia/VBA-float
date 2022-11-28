# VBA-float

## Intro
This class module is a wrapper that allows to treat numbers as large text strings expressed in a variant of scientific notation. 

Despite the above, a representation similar to that of programming languages is adopted (using the `E` symbol instead of the `x10^` characters). A peculiarity that differentiates this implementation from the `float` data type is that it is allowed to obtain a representation (cohort) given a magnitude (exponent) at the user's request, so that the decimal represented can contain a variable number of characters representing its integer part instead of a single character in the `0-9` range used in computational data types.

As an added value, methods have been integrated to perform calculations on large integers (addition, subtraction, multiplication, division, exponentiation) and to make inferences (comparisons) between values stored in two instances of the same class. 

Although `VBAfloat` is not currently focused on performance, the class is capable of delivering computations over a thousand digits in the blink of an eye, thanks to the implementation of high performance routines for three of the four basic operations, with division being the least optimized process in the package.

There is currently an [open issue](https://github.com/ws-garcia/VBA-float/issues/1) regarding the division. Although a routine has been implemented in reference literature with high credibility, there are cases in which the performance of the division algorithm does not depend on the length of the dividend, nor the divisor, nor the quotient, and rather it seems to be a function of the base selected to perform the computations. For example, if you use the division method to compute `987659876598765987654321098765432109876543210 / 9876598765987659876`, using a base `B=10^6` you can notice that it takes 2000 times longer than computing the same quotient with a base `B=10^5` (a very unexpected result). As more of you help, maybe it is all a misinterpretation of the algorithm, the probabilities of giving answers to this problem increase! 

## Using the code
```
Sub Test()
    Dim Number As Float
    Dim summand As Float
	 
    'Initialize
    Set Number = New Float
    Set summand = New Float
    summand.Create "-11.11" 'Get a like float representation
    With Number
        .Create "-9999999"
        Debug.Print "Value: "; .value
        Debug.Print "Representation: "; .Representation
        .Sum summand, 3 'A+B using a base equal to 10^3
        Debug.Print "Value after sum: "; .value
        Debug.Print "Representation after sum: "; .Representation
        Debug.Print "Base cohort significand: "; .Cohort(0).Significand    'Output a decimal
        Debug.Print "--------------------------------------------------"
    End With
End Sub
```