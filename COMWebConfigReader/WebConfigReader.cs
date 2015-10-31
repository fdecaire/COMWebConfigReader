using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;

namespace COMWebConfigReader
{
	[Guid("E9447028-ECEA-4D48-9B71-E798D2D6E72A"),
	ClassInterface(ClassInterfaceType.None)]
	public class WebConfigReader : IWebConfigReader
	{
		public string ReadAppSetting(string name)
		{
			return ReadXMLStringFromWebConfig(name, "appsettings", "add", "key", "value");
		}

		public string ReadConnectionString(string name)
		{
			return ReadXMLStringFromWebConfig(name, "connectionstrings", "add", "name", "connectionString");
        }

		private string ReadXMLStringFromWebConfig(string name, string section, string itemName, string attributeName, string attributeValue)
		{
			try
			{
				var siteRoot = this.GetType().Assembly.Location;
				siteRoot = siteRoot.Replace("COM\\" + this.GetType().Assembly.GetName().Name + ".dll", "");

				using (StreamReader reader = new StreamReader(siteRoot + "web.config"))
				{
					XmlDocument document = new XmlDocument();
					document.LoadXml(reader.ReadToEnd());

					foreach (XmlNode element in document.ChildNodes)
					{
						if (element.Name.ToLower() == "configuration")
						{
							foreach (XmlNode configurationSubNodes in element.ChildNodes)
							{
								if (configurationSubNodes.Name.ToLower() == section)
								{
									foreach (XmlNode connectionStringNodes in configurationSubNodes.ChildNodes)
									{
										if (connectionStringNodes.Name.ToLower() == itemName)
										{
											string connectionStringName = connectionStringNodes.Attributes[attributeName].Value;

											if (connectionStringName.ToLower() == name)
											{
												return connectionStringNodes.Attributes[attributeValue].Value;
											}
										}
									}
								}
							}
						}
					}
				}
			}
			catch
			{

			}

			return "";
		}
	}
}
