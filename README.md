# Inpli Subnetting macros for Excel

## Intro

Included in this repository are some useful macros for use with Excel to make IP subnetting a little easier

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

### Explaination
The IPModBits function has 4 parameters

* ip - which is a string formated as an IPv4 address. There is no error handling on this and therefore must be formatted correctly or it will cause an error
* fromBit - the starting bit number counting from zero to change
* toBit - the ending bit number counting from zero to change
* v - the value to insert in the change bits