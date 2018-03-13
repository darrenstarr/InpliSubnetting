# Inpli Subnetting macros for Excel

## Intro

Included in this repository are some useful macros for use with Excel to make IP subnetting a little easier

## How to use

### Windows

1) Press Alt+F11 to open the Microsoft Visual Basic for Applications window and click "Insert" and then "Module".

2) Paste the content of the script file (for example ipsubnettingmacros.txt) into the file.

Now those functions should be available within the spreadsheet.

In order to use these macros within a spreadsheet, the spreadsheet must be saved as "Excel Macro-Enabled Workbook".

## Example of changing bits in an IP address

### Before
|    |          A |    B |      C |                                                           D |
|----|------------|------|--------|-------------------------------------------------------------|
|  1 | 172.16.0.0 |      |        |                                                             |
|  2 |      24    | Zone | Subnet | Modified                                                    |
|  3 |            |    4 |      5 | =IPModBits(IPModBits($A$1, 16, 19, B3),20,23,C3) &"/"& $A$2 |

### After
|    |          A |    B |      C |              D |
|----|------------|------|--------|----------------|
|  1 | 172.16.0.0 |      |        |                |
|  2 |      24    | Zone | Subnet | Modified       |
|  3 |            |    4 |      5 | 172.16.69.0/24 |

### Function documentation
The IPModBits function has 4 parameters

* ip - which is a string formated as an IPv4 address. There is no error handling on this and therefore must be formatted correctly or it will cause an error
* fromBit - the starting bit number counting from zero to change
* toBit - the ending bit number counting from zero to change
* v - the value to insert in the change bits

## Example of changing bits in an IPv6 Address

Operationally, this works the same as the IPv4 address example above, using the following information :

```excel
Cell J10 = 2001:db8:1:0::
```

Then applying the formula :

```excel
=IPv6ModBits(J10,32,39,23)
```

Bits 32 through 39 are replaced with the value of 23 and the output is as follows

```excel
2001:db8:1701:db8:1:0::
```
### Explaination

To see this clearly consider that :

```excel
1 = 0000 0000 0000 0001
```

and

```excel
23 = 0001 0111
```

therefore replacing bits 32 through 39 would produce

```excel
0001 0111 0000 0001 = 1701
```

### Function Documentation

The IPv6ModBits function has 4 parameters

* ipv6 - which is a string formated as an IPv6 address. There is no error handling on this and therefore must be formatted correctly or it will cause an error
* fromBit - the starting bit number counting from zero to change
* toBit - the ending bit number counting from zero to change
* v - the value to insert in the change bits. Any value greater than 31 bits will overflow Excel

## Function IPv6ToBinary

As the name suggests, given an IPv6 address it produces a binary string.

### Example

```excel
E9 = Fe80::001:022:3333:4444
I9 = =IPv6ToBinary(E9)
```

The output of I9 is as follows

```excel
11111110100000000000000000000000000000000000000000000000000000000000000000000001000000000010001000110011001100110100010001000100
```

### Function Documentation

The IPv6ToBinary function has a single parameter

* ipv6 - and IPv6 address string to convert to binary

## Function CompressIPv6Address

This function takes a suboptimal IPv6 address string as input and recompresses it to "optimal" format.

### Example

```excel
E9 = Fe80::001:022:3333:4444
H9 = =CompressIPv6Address(E9)
```

The output of H9 is as follows

```excel
Fe80::1:22:3333:4444
```

### Function Documentation

CompressIPv6Address has a single parameter

* ipv6 - the IPv6 address to compress

## Function FullExpandIPv6

This takes a compressed IPv6 address and decompresses it to it's full form

### Example

```excel
E9 = Fe80::001:022:3333:4444
E10 = 2001:db8:1:0::
E11 = ::

G9 = =FullExpandIPv6(E9)
G10 = =FullExpandIPv6(E10)
G11 = =FullExpandIPv6(E11)
```

The output of the formula cells are as follows

```excel
G9 = fe80:0000:0000:0000:0001:0022:3333:4444
G10 = 2001:0db8:0001:0000:0000:0000:0000:0000
G11 = 0000:0000:0000:0000:0000:0000:0000:0000
```

### Function documentation

FullExpandIPv6 has a single input parameter

* ipv6 - the IPv6 address to decompress to full form

## Function ExpandIPv6

ExpandIPv6 performs a partial decompression of an IPv6 address to ensure that all the hextets are present and :: substitution for runs of zeros aren't present. The function does not alter the form of the individual hextets.

### Example

```excel
E9 = Fe80::001:022:3333:4444
E10 = 2001:db8:1:0::
E11 = ::

F9 = =ExpandIPv6(E9)
F10 = =ExpandIPv6(E10)
F11 = =ExpandIPv6(E11)
```

The resulting output is as follows
```excel
F9 = Fe80:0:0:0:001:022:3333:4444
F10 = 2001:db8:1:0:0:0:0:0
F11 = 0:0:0:0:0:0:0:0
```

### Function Documentation

ExpandIPv6 has a single input parameter

* ipv6 - the IPv6 address string to expand to contain all 8 hextets

## Change bits within a VLAN Identifier

ModVLANBits can be used to change the value of a VLAN identifier by injecting an integer value within a certain range of bits.

This code does not have extensive error handling and therefore values passed to the function should be within valid ranges for VLAN IDs, bit values and replacement values. For example, if you are trying to replace 4 bits with a value like 127 which would use 7 bits, it will not work. However if you're replacing those same four bits with the number 3, it will work just fine.

### Example

```excel 
A19 = 256
B19 = 15
C19 = 15
D19 = =ModVLANBits(ModVLANBits(A19, 4,7,B19), 8,11,C19)
```

The resulting value for D19 would be as follows :

```excel
D19 = 511
```

### Explaination

Let's first convert the original VLAN ID

```excel
256 = 0001 0000 0000

15 = 1111
```

When evaluating

```Excel
ModVLANBits(A19, 4,7,B19)
```

The following is the process

```excel
                    11
Bits    0123 4567 8901
256 =   0001 0000 0000
Range        |  |
15 =         1111

Result  0001 1111 0000 = 496
````

Then the compounded statement

```excel
=ModVLANBits(ModVLANBits(A19, 4,7,B19), 8,11,C19)
```

would operate as follows

```excel
                    11
Bits    0123 4567 8901
256 =   0001 0000 0000
Range        |  | |  |
15 =         1111 |  |
Range             |  |
15 =              1111
Result  0001 1111 1111 = 511
````

### Function Documentation

ModVLANBits has four input parameters

* vlan - The base VLAN ID to alter (must be between 0 and 4095 inclusive)
* fromBit - the zero based starting bit
* toBit - the zero based ending bit
* v - the value to inject (should not be larger than the specified bits will allow)
