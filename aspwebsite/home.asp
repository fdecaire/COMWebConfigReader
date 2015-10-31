<%

' www.webconfigreader.com

Dim MyComObject
Dim MyText

Set MyComObject = Server.CreateObject("COMWebConfigReader.WebConfigReader")

MyText = MyComObject.ReadAppSetting("appsetting2")

%>
<html>
<head></head>
<body>
    <%=MyText %>
</body>
</html>
