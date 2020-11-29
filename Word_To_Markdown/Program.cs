using System;
using System.Collections.Generic;

namespace Word_To_Markdown
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length >= 1)
			{
				Console.WriteLine("Processing. This could take some time, depending on the size of the document(s) provided.");
				List<string> filenames = new List<string>();
				int i = 0;
				string curr = "";
				while (i < args.Length)
				{
					while (!args[i].EndsWith(".docx"))
					{
						if (curr != "")
						{
							curr = curr + " " + args[i];
						}
						else
						{
							curr += args[i];
						}
						i++;
					}

					curr = curr + " " + args[i];
					filenames.Add(curr);
					curr = "";
					i++;
				}

				if (filenames.Count == 0)
				{
					Console.Error.WriteLine("Input does not contain valid filename(s).");
					return;
				}
				else
				{
					foreach (string s in filenames)
					{
						DocXParser docXParser = new DocXParser(s);
						docXParser.process();
					}
				}
			}

			else
			{
				Console.Error.WriteLine("Please drop a .docx file or .docx files onto the .exe.");
				Console.ReadKey();
				return;
			}
		}
	}
}
