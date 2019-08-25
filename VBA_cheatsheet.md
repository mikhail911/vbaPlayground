# VBA Cheatsheet

## Table of contents

1. [Data types](#1.-Data-types)
2. [Variables](#2.-Variables)
3. [Functions](#3.-Functions)  
3.1. [Function](#3.1.-Function)  
3.2. [Sub](#3.2.-Sub)

### 1. Data types

| Data type | Storage size | Range |
|-----------|--------------|-------|
| Boolean   | 2 bytes      | True or false|
| Byte | 1 byte | 0 to 255 |
| Currency | 8 bytes | -922,337,203,685,477.5808 to 922,337,203,685,477.5807|
| Date | 8 bytes | January 1, 100, to December 31, 9999|
| Decimal | 14 bytes | ±79,228,162,514,264,337,593,543,950,335 with no decimal point <br /> ±7.9228162514264337593543950335 with 28 places to the right of the decimal |
| Double | 8 bytes | -1.79769313486231E308 to -4.94065645841247E-324 for negative values <br /> 4.94065645841247E-324 to 1.79769313486232E308 for positive values |
| Integer | 2 bytes | -32,768 to 32,767 |
| Long (Long integer) | 4 bytes | -2,147,483,648 to 2,147,483,647 |
| LongLong | 8 bytes | -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 <br />Valid on 64-bit platforms only. |
| Single | 4 bytes | -3.402823E38 to -1.401298E-45 for negative values <br />1.401298E-45 to 3.402823E38 for positive values|
| String | 10 bytes + string length | 0 to approximately 2 billion |

### 2. Variables

```vbnet
Dim x As Integer
x = 2
```

### 3. Functions

### 3.1. Function

Function can return a value, if variable name used in function is the same as function name, it will be returned by default.

```vbnet
Function Square(x As Double) As Double
    Square = x * x
End Function 
```

### 3.2. Sub

Sub can only perform an action.

```vbnet
Sub Square(x As Double)
    MsgBox x * x
End Sub
```