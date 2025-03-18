using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Xml.Linq;
using NGS.Templater;

namespace CollapseRegion
{
	public class Program
	{
		public static void Main(string[] args)
		{
			File.Copy("template/Collapse.docx", "Collapse.docx", true);
			var application1 =
					new Application()
							.setPaybackYears(20)
							.setUcCheck(true).setUcCheckResponse("Ok")
							.setApplicant(new Applicant("first applicant").setFrom("Google", 2012, 11).addChild("Mary"));
			application1.getLoans().Add(new Loan("Big Bank", 10000, Color.Blue));
			application1.getLoans().Add(new Loan("Small Bank", 2000, Color.Lime));
			var application2 =
					new Application().hideLoans()
							.setPaybackYears(15)
							.setUcCheck(false)
							.setUcCheckResponse("Not good enough")
							.setApplicant(new Applicant("second applicant").setFrom("Apple", 2015, 12))
							.setCoApplicant(new Applicant("second co-applicant").setFromUntil("IBM", 2014, 11, 2015, 12));
			var application3 =
					new Application()
							.setPaybackYears(10)
							.setUcCheck(true).setUcCheckResponse("Ok")
							.setApplicant(new Applicant("third applicant").setFrom("Microsoft", 2010, 1).addChild("Jack").addChild("Jane"));
			var yes = XElement.Parse("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">YES</w:p>");
			var no = XElement.Parse("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">NO</w:p>");
			var factory = Configuration.Builder.Include((value, metadata, path, position, templater) =>
			{
				var str = value as string;
				if (str != null && metadata.StartsWith("collapseIf("))
				{
					//Extract the matching expression
					var expression = metadata.Substring("collapseIf(".Length, metadata.Length - "collapseIf(".Length - 1);
					if (str == expression)
					{
						//remove the context around the specific property
						if (position == -1)
						{
							//when position is -1 it means non sharing tag is being used, in which case we can resize that region via "standard" API
							templater.Resize(new[] { path }, 0);
						}
						else
						{
							//otherwise we need to use "advanced" resize API to specify which exact tag to replace
							templater.Resize(new[] { new TagPosition(path, position) }, 0);
						}
						return Handled.NestedTags;
					}
				}
				return Handled.Nothing;
			}).Build();

			using (var doc = factory.Open("Collapse.docx"))
			{
				doc.Process(new[] { application1, application2, application3 });
			}
			Process.Start(new ProcessStartInfo("Collapse.docx") { UseShellExecute = true });
		}
	}
}
