using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using System.Linq;

namespace Word_To_Markdown
{
	class DocXParser
	{
		private readonly string filename;
		public enum Style
		{
			DOC_HEADER,
			SECTION_HEADER,
			LIST_ELEMENT,
			BASIC_LINE,
			EMPTY_LINE
		}

		public DocXParser(string filename)
		{
			this.filename = filename;
		}

		public void process()
		{
			// Load docx into memory
			Package wordPackage = Package.Open(filename, FileMode.Open, FileAccess.Read);
			using WordprocessingDocument document = WordprocessingDocument.Open(wordPackage); 
			Body body = document.MainDocumentPart.Document.Body;

			List<string> lines = new List<String>();

			HashSet<string> uniqueStyles = new HashSet<string>();
			foreach (var child in body.ChildElements)
			{
				// Some method here to determine the "style" of the current block.
				// Will need to parse XML out to find a style value (header1, header2, list, etc)
				// Probably useful to just use some sort of enum here.
				Style style = parseStyle(child.InnerXml);

				if (style == Style.EMPTY_LINE && child.InnerText != "")
				{
					style = Style.SECTION_HEADER;
				}

				int indentLevel = 0;
				if (style == Style.LIST_ELEMENT)
				{
					int indexOfIndent = child.InnerXml.IndexOf("<w:numPr><w:ilvl w:val=\"");
					int indexOfOpeningQuote = indexOfIndent + "<w:numPr><w:ilvl w:val=\"".Length;
					int indexOfClosingQuote = child.InnerXml.Substring(indexOfOpeningQuote).IndexOf("\"");

					if (indexOfIndent != -1)
					{
						indentLevel = Int32.Parse(child.InnerXml.Substring(indexOfOpeningQuote, indexOfClosingQuote));
					}
				}

				// Once style is determined, conditionally add to the MarkdownDoc out element.
				switch (style)
				{
					case Style.DOC_HEADER:
						lines.Add("# " + child.InnerText);
						break;
					case Style.SECTION_HEADER:
						lines.Add("## " + child.InnerText);
						break;
					case Style.LIST_ELEMENT:
						if (child.InnerText != "")
						{
							string indents = "";
							for (int i = 0; i < indentLevel; i++)
							{
								indents += "  ";
							}
							lines.Add(indents + "- " + child.InnerText);
						}
						else
						{
							lines.Add("\n");
						}
						break;
					case Style.BASIC_LINE:
						lines.Add(child.InnerText);
						break;
					case Style.EMPTY_LINE:
						lines.Add("\n");
						break;
				}
			}

			// Output is just writing the markdown to a .md file.
			// Trim output filename
			string outname = filename.Substring(0, filename.LastIndexOf(".docx"));
			outname = outname.Substring(outname.LastIndexOf("\\")+1);
			outname = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + outname + ".md";
			File.WriteAllLines(outname, lines);
		}


		public Style parseStyle(string xml_string)
		{
			// trying to convert this to an XNode was proving difficult because the linter was refusing to see that XNode.FirstNode could be nested
			int indexOfStyle = xml_string.IndexOf("w:pStyle w:val=\"");

			if (indexOfStyle == -1)
			{
				return Style.BASIC_LINE;
			}

			int indexOfOpeningQuote = indexOfStyle + "w:pStyle w:val=\"".Length;
			int indexOfClosingQuote = xml_string.Substring(indexOfOpeningQuote).IndexOf("\"");

			string style = xml_string.Substring(indexOfOpeningQuote, indexOfClosingQuote);

			switch (style)
			{
				case "Heading1":
					return Style.DOC_HEADER;
				case "Subtitle":
					return Style.DOC_HEADER;
				case "Title":
					return Style.DOC_HEADER;
				case "Heading2":
					return Style.SECTION_HEADER;
				case "ListParagraph":
					return Style.LIST_ELEMENT;
				case "":
					return Style.EMPTY_LINE;
				default:
					return Style.BASIC_LINE;
			}
		}
	}
}
