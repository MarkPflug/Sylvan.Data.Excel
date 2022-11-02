using System.IO;
using System.Xml;
using Xunit;

namespace Sylvan.Data.Excel;

public class XmlWriterTests
{
	[Fact]
	public void NamespaceTest()
	{
		var sw = new StringWriter();
		var xw = XmlWriter.Create(sw);

		xw.WriteStartElement("a");
		xw.WriteAttributeString("xmlns", "r", null, "https://example.com");
		xw.WriteStartElement("b");		
		xw.WriteAttributeString("id", "https://example.com", "1");
		xw.WriteEndElement();
		xw.WriteEndElement();
		xw.Dispose();
		var str = sw.ToString();
	}
}
