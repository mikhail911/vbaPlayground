# VBA Cheatsheet

## Table of contents

1. [Data types](#1-Data-types)
2. [Variables](#2-Variables)
3. [Operators](#3-Operators)
4. [Constant](#4-Constant)
5. [Functions](#5-Functions)  
5.1. [Function](#51-Function)  
5.2. [Sub](#52-Sub)  
5.3. [Conversion functions](#53-Conversion-functions)  
5.4. [Math functions](#54-Math-functions)
6. [Comments](#6-Comments)
7. [Loops](#7-Loops)  
7.1. [For Loop](#71-For-Loop)  
7.2. [For Each Loop](#72-For-Each-Loop)  
7.3. [While Loop](#73-While-Loop)  
7.4. [Do While Loop](#74-Do-While-Loop)  
7.5. [Do Until Loop](#75-Do-Until-Loop)

### 1. Data types

| Data type           | Storage size             | Range                                                                                                                                           |
|---------------------|--------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------|
| Boolean             | 2 bytes                  | True or false                                                                                                                                   |
| Byte                | 1 byte                   | 0 to 255                                                                                                                                        |
| Currency            | 8 bytes                  | -922,337,203,685,477.5808 to 922,337,203,685,477.5807                                                                                           |
| Date                | 8 bytes                  | January 1, 100, to December 31, 9999                                                                                                            |
| Decimal             | 14 bytes                 | ±79,228,162,514,264,337,593,543,950,335 with no decimal point <br /> ±7.9228162514264337593543950335 with 28 places to the right of the decimal |
| Double              | 8 bytes                  | -1.79769313486231E308 to -4.94065645841247E-324 for negative values <br /> 4.94065645841247E-324 to 1.79769313486232E308 for positive values    |
| Integer             | 2 bytes                  | -32,768 to 32,767                                                                                                                               |
| Long (Long integer) | 4 bytes                  | -2,147,483,648 to 2,147,483,647                                                                                                                 |
| LongLong            | 8 bytes                  | -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 <br /> Valid on 64-bit platforms only.                                                  |
| Object              | 4 bytes                  | Any Object reference                                                                                                                            |
| Single              | 4 bytes                  | -3.402823E38 to -1.401298E-45 for negative values <br /> 1.401298E-45 to 3.402823E38 for positive values                                        |
| String              | 10 bytes + string length | 0 to approximately 2 billion                                                                                                                    |
| Variant             | 16 bytes                 | Any numeric value up to the range of a Double                                                                                                   |

### 2. Variables

```vbnet
Dim x As Integer
x = 2
```
### 3. Operators

| Operator | Description                                                  |
|:--------:|--------------------------------------------------------------|
| +        | Adds the two operands                                        |
| -        | Subtracts the second operand from the first                  |
| *        | Multiplies both the operands                                 |
| /        | Divides the numerator by the denominator                     |
| %        | Modulus operator and the remainder after an integer division |
| ^        | Exponentiation operator                                      |

### 4. Constant

Constant hold a value that cannot be change during script execution.

Example:

```vbnet
Private Sub Test_Constant()
   Const Pi As Double = 3.14159265359
   MsgBox "Pi: " & Pi
End Sub
```

### 5. Functions

### 5.1. Function

Function can return a value, if variable name used in function is the same as function name, it will be returned by default.

Example:

```vbnet
Function Square(x As Double) As Double
    Square = x * x
End Function 
```

### 5.2. Sub

Sub can only perform an action.

Example:

```vbnet
Sub Square(x As Double)
    MsgBox x * x
End Sub
```

### 5.3. Conversion functions

| Function | Argument    | Description                                                                                                              |
|----------|-------------|--------------------------------------------------------------------------------------------------------------------------|
| Asc      | string      | Returns an Integer representing the character code corresponding to the first letter in a string                         |
| Chr      | charcode    | Returns a String containing the character associated with the specified character code                                   |
| CVErr    | errornumber | Returns a Variant of subtype Error containing an error number specified by the user                                      |
| Format   | expression  | Returns a Variant (String) containing an expression formatted according to instructions contained in a format expression |
| Hex      | number      | Returns a String representing the hexadecimal value of a number                                                          |
| Oct      | number      | Returns a Variant (String) representing the octal value of a number                                                      |
| Str      | number      | Returns a Variant (String) representation of a number                                                                    |
| Val      | string      | Returns the numbers contained in a string as a numeric value of appropriate type                                         |


### 5.4. Math functions

| Function | Argument | Description                                                                                     |
|----------|----------|-------------------------------------------------------------------------------------------------|
| Abs      | number   | Returns a value of the same type that is passed to it specifying the absolute value of a number |
| Atn      | number   | Returns a Double specifying the arctangent of a number                                          |
| Cos      | number   | Returns a Double specifying the cosine of an angle                                              |
| Exp      | number   | Returns a Double specifying e (the base of natural logarithms) raised to a power                |
| Int, Fix | number   | Returns the integer portion of a number                                                         |
| Log      | number   | Returns a Double specifying the natural logarithm of a number                                   |
| Rnd      | number   | Returns a Single containing a pseudo-random number                                              |
| Sgn      | number   | Returns a Variant (Integer) indicating the sign of a number                                     |
| Sin      | number   | Returns a Double specifying the sine of an angle                                                |
| Sqr      | number   | Returns a Double specifying the square root of a number                                         |
| Tan      | number   | Returns a Double specifying the tangent of an angle                                             |

### 6. Comments

Unlike most programming languages, VBA doesn't provide multiline comments.

Example:

```vbnet
' Comment

Rem this is a comment too
```
### 7. Loops

### 7.1. For Loop

Example:
```vbnet
Private Sub Loop_Test()
   ' Count to ten with step by 1
   Dim x As Integer
   x = 10
   
   For i = 0 To x Step 1
      Debug.Print i
   Next
End Sub
```

### 7.2. For Each Loop

### 7.3. While Loop

### 7.4. Do While Loop

### 7.5. Do Until Loop