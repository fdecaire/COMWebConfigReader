using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

// directory:	 cd "C:\Users\Frank\Documents\Visual Studio 2015\Projects\COMWebConfigReader\SampleWebApplication\COM"
// sn tool:      "C:\Program Files (x86)\Microsoft SDKs\Windows\v8.1A\bin\NETFX 4.5.1 Tools\sn" -k WebConfigReader.snk
// regasm tool:  "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm" /codebase COMWebConfigReader.dll
// regasm tool:  "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm" /u COMWebConfigReader.dll

namespace COMWebConfigReader
{
	[Guid("BB6DE975-4483-47CD-80F0-7F75173E64FC"),
	InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IWebConfigReader
	{
		string ReadAppSetting(string name);
		string ReadConnectionString(string name);
	}
}
