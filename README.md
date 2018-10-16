# IBAN-validator
VB.NET IBAN validator

### How To Use
- Copy the class iban.vb into your App_Code folder in your project
- Call the namespace using ```iban.cIban.ValidateIBAN("<YOUR_IBAN_NUMBER>")```

### Example
```
Dim check_iban as integer = iban.cIban.ValidateIBAN("BE68539007547034")
if check_iban = 0 then
    ' IBAN is valid
else
    ' IBAN is not valid
end if
```

### Error code's
```
0 = ok, IBAN is valid
1 = Country code unknow, mod97 check is ok
2 = Length invalid with country code
3 = Mod97 check failed
4 = Format not recognized
```