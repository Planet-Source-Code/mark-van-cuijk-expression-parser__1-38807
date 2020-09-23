<div align="center">

## Expression parser


</div>

### Description

This module can be used to parse a mathematical expression given as a string. The result is just a double value, which can be used for anything you'd like to.
 
### More Info
 
A string containing an expression. Spaces between operators and operands are ignored. The following operators are supported:

- exponentiations (^)

- multiplications (*)

- divisions (/)

- modulus (%)

- additions (+)

- substractions (-)

Besides it's possible to use some built-in functions, putting parameters between braces:

- sin, asin, cos, acos, tan, atan

- int (returns integer part)

- frac (returns decimal part)

- log (logarithm with base 10)

- ln (natural logarithm with base e)

- abs (absolute value)

- sign (returns -1, 0 or 1)

- rnd (random number between 0 and param)

- sqrt (square root)

You can define your own functions. These functions are implemented in BasicFunctions.

I haven't implemented variable support, because the best way to do so depends on how you're going to use this code. The ReadVariable-function is called when the value of a variable is needed. It's your own job to implement it.

Last but not least sub-expressions are supported by encapsulating them with braces.

Example (showing some features):

12 * (14-36 ^ (-sin(-4*sqrt(18)))) / (15 * rnd(4) +1)

Just call ParseExpression("1+1") to see what happens.

Variable support isn't implemented.

A double value containing the result.

Operator precedence:

1. Functions and variables

2. Sub-expressions

3. Exponentiations (^)

4. Multiplications (*) and divisions (/)

5. Modulus (%)

6. Additions (+) and substractions (-)


<span>             |<span>
---                |---
**Submitted On**   |2002-09-09 13:19:04
**By**             |[Mark van Cuijk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-van-cuijk.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Expression128807992002\.zip](https://github.com/Planet-Source-Code/mark-van-cuijk-expression-parser__1-38807/archive/master.zip)








