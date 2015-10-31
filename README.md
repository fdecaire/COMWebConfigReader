# COM Web Configuration Reader

This project is used to read AppSettings or ConnectionStrings from a web.config file using a COM application.  

To use this COM application, create a COM directory in your web directory, then execute from Classic ASP with 
the following syntax:

```VB
Dim ConfigurationManager
Dim MyAppSetting

Set ConfigurationManager = Server.CreateObject("COMWebConfigReader.WebConfigReader")

MyAppSetting = ConfigurationManager.ReadAppSetting("appsetting2")
```
