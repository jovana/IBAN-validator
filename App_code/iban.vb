''*******************************************************************************
''   European IBAN validator
''
'' DESCRIPTION:
''   European IBAN validator.
''   Use this functions to validate the IBAN bank account numbers 
''
'' The validation has built using this document:
'' http://nl.wikipedia.org/wiki/International_Bank_Account_Number

Namespace IBAN
    Public Class cIban

        
        ''' <summary>
        ''' This function checkes if the given IBAN number correct is.
        ''' </summary>
        ''' <param name="IBAN">string IBAn number</param>
        ''' <returns></returns>
        Public Shared Function ValidateIBAN(ByVal p_IBAN As String) As Integer
            Dim objRegExp As Object
            Dim IBANformat As String
            Dim IBANNR As String
            Dim ReplaceChr As String
            Dim ReplaceBy As String
            Dim LeftOver As Integer
            Dim Nr, i, digit As Integer

            '' Clear the input (remove spaces and convert to upper case) 
            p_IBAN = Replace(UCase(p_IBAN), " ", "")

            '' Chekc if the length is correct (length 15-31 characters)
            ' first countrycode, then two digits, then number
            '-------------------------------------------------------------------------
            objRegExp = CreateObject("vbscript.regexp")
            objRegExp.IgnoreCase = True
            objRegExp.Global = True
            objRegExp.Pattern = "[a-zA-Z]{2}[0-9]{2}[a-zA-Z0-9]{11}([a-zA-Z0-9]?){0,16}"
            IBANformat = objRegExp.Test(p_IBAN)

            '' Validate the format
            If IBANformat = False Then
                'VALIDATEIBAN = "FORMAT NOT RECOGNIZED"
                VALIDATEIBAN = 4
            Else
                '' Move the first 4 characters to the rear
                IBANNR = Right(p_IBAN, Len(p_IBAN) - 4) & Left(p_IBAN, 4)

                '' Replace all letters by numbers
                For Nr = 10 To 35
                    ReplaceChr = Chr(Nr + 55)
                    ReplaceBy = Trim(CStr(Nr))
                    IBANNR = Replace(IBANNR, ReplaceChr, ReplaceBy)
                Next

                '' Check the number for Mod 97
                LeftOver = 0
                For i = 1 To Len(IBANNR)
                    digit = CInt(Mid(IBANNR, i, 1))
                    LeftOver = (10 * LeftOver + digit) Mod 97
                Next

                If LeftOver = 1 Then
                    If Len(p_IBAN) = IBANLEN(Left(p_IBAN, 2)) Then
                        VALIDATEIBAN = 0 'IBAN OK

                    ElseIf IBANLEN(Left(p_IBAN, 2)) = 0 Then
                        VALIDATEIBAN = 1 ' COUNTRYCODE UNKNOWN, 97 CHECK OK
                    Else
                        VALIDATEIBAN = 2 ' LENGTH INVALID WITH COUNTRYCODE
                    End If
                Else
                    VALIDATEIBAN = 3 ' 97 CHECK FAILED
                End If
            End If
        End Function

        ''' <summary>
        ''' This function return the length of the IBAN by country
        ''' </summary>
        ''' <param name="CountryCode"></param>
        ''' <returns></returns>
        Private Shared Function IBANLEN(ByVal CountryCode As String) As Integer

            '' If the country code is more the two, error
            If Len(CountryCode) <> 2 Then
                IBANLEN = 0
            End If

            '' List for the country codes: http://nl.wikipedia.org/wiki/ISO_3166-1
            Select Case CountryCode
                Case "AL"
                    IBANLEN = 28
                Case "AD"
                    IBANLEN = 24
                Case "AE"
                    IBANLEN = 23
                Case "AO"
                    IBANLEN = 25
                Case "AT"
                    IBANLEN = 20
                Case "AZ"
                    IBANLEN = 28
                Case "BH"
                    IBANLEN = 22
                Case "BE"
                    IBANLEN = 16
                Case "BA"
                    IBANLEN = 20
                Case "BF"
                    IBANLEN = 27
                Case "BI"
                    IBANLEN = 16
                Case "BJ"
                    IBANLEN = 28
                Case "BR"
                    IBANLEN = 29
                Case "BG"
                    IBANLEN = 22
                Case "CH"
                    IBANLEN = 21
                Case "CI"
                    IBANLEN = 28
                Case "CM"
                    IBANLEN = 27
                Case "CR"
                    IBANLEN = 21
                Case "CV"
                    IBANLEN = 25
                Case "CY"
                    IBANLEN = 28
                Case "CZ"
                    IBANLEN = 24
                Case "DE"
                    IBANLEN = 22
                Case "DK"
                    IBANLEN = 18
                Case "DO"
                    IBANLEN = 28
                Case "EE"
                    IBANLEN = 20
                Case "ES"
                    IBANLEN = 24
                Case "FO"
                    IBANLEN = 18
                Case "FI"
                    IBANLEN = 18
                Case "FR"
                    IBANLEN = 27
                Case "GB"
                    IBANLEN = 22
                Case "GE"
                    IBANLEN = 22
                Case "GI"
                    IBANLEN = 23
                Case "GR"
                    IBANLEN = 27
                Case "GL"
                    IBANLEN = 18
                Case "GT"
                    IBANLEN = 28
                Case "HR"
                    IBANLEN = 21
                Case "HU"
                    IBANLEN = 28
                Case "IE"
                    IBANLEN = 22
                Case "IL"
                    IBANLEN = 23
                Case "IR"
                    IBANLEN = 26
                Case "IS"
                    IBANLEN = 26
                Case "IT"
                    IBANLEN = 27
                Case "KZ"
                    IBANLEN = 20
                Case "KW"
                    IBANLEN = 30
                Case "LB"
                    IBANLEN = 28
                Case "LI"
                    IBANLEN = 21
                Case "LT"
                    IBANLEN = 20
                Case "LU"
                    IBANLEN = 20
                Case "LV"
                    IBANLEN = 21
                Case "MC"
                    IBANLEN = 27
                Case "MD"
                    IBANLEN = 24
                Case "ME"
                    IBANLEN = 22
                Case "MG"
                    IBANLEN = 27
                Case "MK"
                    IBANLEN = 19
                Case "ML"
                    IBANLEN = 28
                Case "MT"
                    IBANLEN = 31
                Case "MR"
                    IBANLEN = 27
                Case "MU"
                    IBANLEN = 30
                Case "MZ"
                    IBANLEN = 25
                Case "NL"
                    IBANLEN = 18
                Case "NO"
                    IBANLEN = 15
                Case "PK"
                    IBANLEN = 24
                Case "PS"
                    IBANLEN = 29
                Case "PL"
                    IBANLEN = 28
                Case "PT"
                    IBANLEN = 25
                Case "RO"
                    IBANLEN = 24
                Case "RS"
                    IBANLEN = 22
                Case "SA"
                    IBANLEN = 24
                Case "SE"
                    IBANLEN = 24
                Case "SI"
                    IBANLEN = 19
                Case "SK"
                    IBANLEN = 24
                Case "SM"
                    IBANLEN = 27
                Case "SN"
                    IBANLEN = 28
                Case "TN"
                    IBANLEN = 24
                Case "TR"
                    IBANLEN = 26
                Case "UA"
                    IBANLEN = 29
                Case "VG"
                    IBANLEN = 24
                Case Else
                    IBANLEN = 0
            End Select
        End Function


    End Class
End Namespace