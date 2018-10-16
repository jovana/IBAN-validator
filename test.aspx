<%@ Page Language="VB" debug="true" %>

<%
' Test the IBAN Number
response.write("BE68539007547034: " & iban.cIban.ValidateIBAN("BE68539007547034") & "<br>")
response.write("AD1200012030200359100100: " & iban.cIban.ValidateIBAN("AD1200012030200359100100") & "<br>")
response.write("NL91ABNA0417164300: " & iban.cIban.ValidateIBAN("NL91ABNA0417164300") & "<br>")
response.write("BE68512907547034: " & iban.cIban.ValidateIBAN("BE68512907547034") & "<br>")
%>