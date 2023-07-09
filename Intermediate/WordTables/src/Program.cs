﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NGS.Templater;

namespace WordDataTable
{
	public class Program
	{
		static object Top10Rows(object argument, string metadata)
		{
			//if we find exact metadata and type invoke the plugin
			if (metadata == "top10" && argument is DataTable)
			{
				var dt = argument as DataTable;
				var newDt = dt.Clone();
				var max = Math.Min(10, dt.Rows.Count);
				for (int i = 0; i < max; i++)
					newDt.ImportRow(dt.Rows[i]);
				return newDt;
			}
			return argument;
		}

		static bool Limit10Table(string prefix, ITemplater templater, DataTable table)
		{
			if (table.Rows.Count > 10)
			{
				//simplified way to match columns against tags
				var tags = table.Columns.Cast<DataColumn>().Select(it => prefix + it.ColumnName).ToList();
				//if any of the found tags matches limit10 condition
				if (tags.Any(t => templater.GetMetadata(t, true).Contains("limit10")))
				{
					templater.Resize(tags, 10);
					for (int i = 0; i < 10; i++)
					{
						DataRow r = table.Rows[i];
						foreach (DataColumn c in table.Columns)
							templater.Replace(prefix + c.ColumnName, r[c]);
					}
					return true;
				}
			}
			return false;
		}

		//for this example position is ignored as it is always -1
		static Handled CollapseNonEmpty(object value, string metadata, string tag, int position, ITemplater templater)
		{
			var dt = value as DataTable;
			if (dt != null && (metadata == "collapseNonEmpty" || metadata == "collapseEmpty"))
			{
				var isEmpty = dt.Rows.Count == 0;
				//loop until all tags with the same name are processed
				do
				{
					var md = templater.GetMetadata(tag, false);
					var collapseOnEmpty = md.Contains("collapseEmpty");
					var collapseNonEmpty = md.Contains("collapseNonEmpty");
					if (isEmpty)
					{
						if (collapseOnEmpty)
						{
							//when position is -1 it means non sharing tag is being used, in which case we can resize that region via "standard" API
							//otherwise we need to use "advanced" resize API to specify which exact tag to replace
							if (position == -1)
								templater.Resize(new[] { tag }, 0);
							else
								templater.Resize(new[] { new TagPosition(tag, position) }, 0);
						}
						else
						{
							//when position is -1 it means non sharing tag is being used, in which case we can just replace the first tag
							//otherwise we can replace that exact tag via position API
							//replacing the first tag is the same as calling replace(tag, 0, value)
							if (position == -1)
								templater.Replace(tag, "");
							else
								templater.Replace(tag, position, "");
						}
					}
					else
					{
						if (collapseNonEmpty)
						{
							if (position == -1)
								templater.Resize(new[] { tag }, 0);
							else
								templater.Resize(new[] { new TagPosition(tag, position) }, 0);
						}
						else
						{
							if (position == -1)
								templater.Replace(tag, "");
							else
								templater.Replace(tag, position, "");
						}
					}
				} while (templater.Tags.Contains(tag));
				return Handled.NestedTags;
			}
			return Handled.Nothing;
		}

		static object LimitDataTable(object parent, object value, string member, string metadata)
		{
			var rs = value as DataTable;
			//check if plugin is applicable
			if (rs == null || !metadata.StartsWith("limit(")) return value;
			int limit = int.Parse(metadata.Substring(6, metadata.Length - 7));
			var dt = rs.Clone();
			for (int i = 0; i < limit; i++)
				dt.ImportRow(rs.Rows[i]);
			//return different object which will be used further in the processing
			return dt;
		}

		static object SumExpression(object parent, object value, string member, string metadata)
		{
			var arr = value as object[];
			//check if plugin is applicable
			if (arr == null || !metadata.StartsWith("Sum(")) return value;
			if (arr.Length == 0 || arr.Contains(null)) return 0;
			var signature = arr[0].GetType();
			var propertyName = metadata.Substring(4, metadata.Length - 5);
			var property = signature.GetField(propertyName);
			var result = (decimal)0;
			foreach (var el in arr)
				result += (decimal)property.GetValue(el);
			return result;
		}

		public static void Main(string[] args)
		{
			File.Copy("template/Tables.docx", "WordTables.docx", true);
			var dtData = new DataTable();
			dtData.Columns.Add("Col1");
			dtData.Columns.Add("Col2");
			dtData.Columns.Add("Col3");
			for (int i = 0; i < 100; i++)
				dtData.Rows.Add("a" + i, "b" + i, "c" + i);
			var dtEmpty = new DataTable();
			dtEmpty.Columns.Add("Name");
			dtEmpty.Columns.Add("Description");
			//for (int i = 0; i < 10; i++)
			//dt4.Rows.Add("Name" + i, "Description" + i);
			var factory =
				Configuration.Builder
				.Include(Top10Rows)
				.Include<DataTable>(Limit10Table)
				.NavigateSeparator(':', null)
				.Include(LimitDataTable)
				.Include(CollapseNonEmpty)
				.Include(SumExpression)
				.Build();
			var dynamicResize1 = new object[7, 3]{
				{"a", "b", "c"},
				{"a", null, "c"},
				{"a", "b", null},
				{null, "b", "c"},
				{"a", null, null},
				{null, null, null},
				{"a", "b", "c"},
			};
			var dynamicResize2 = new object[7, 3]{
				{"a", "b", "c"},
				{null, null, "c"},
				{null, null, null},
				{null, "b", "c"},
				{"a", null, null},
				{null, "b", null},
				{"a", "b", null},
			};
			var map = new Dictionary<string, object>[] {
				new Dictionary<string, object>{{"1", "a"}, {"2","b"},{"3","c"}},
				new Dictionary<string, object>{{"1", "a"}, {"2",null},{"3","c"}},
				new Dictionary<string, object>{{"1", "a"}, {"2","b"},{"3",null}},
				new Dictionary<string, object>{{"1", null}, {"2","b"},{"3","c"}},
				new Dictionary<string, object>{{"1", "a"}, {"2",null},{"3",null}},
				new Dictionary<string, object>{{"1", null}, {"2",null},{"3",null}},
				new Dictionary<string, object>{{"1", "a"}, {"2","b"},{"3","c"}},
			};
			var combined = new Combined
			{
				Beers = new[] 
				{ 
					new Beer { Name = "Heineken", Description = "Green and cold", Columns = new [,] { {"Light", "International"} }},
					new Beer { Name = "Leila", Description = "Blueish", Columns = new [,] { {"Blue", "Domestic"} }}
				},
				Headers = new[,] { { "Bottle", "Where" } }
			};
			var fixedItems = new Fixed[] {
				new Fixed{ Name = "A", Quantity = 1, Price = 42 },
				new Fixed{ Name = "B", Quantity = 2, Price = 23 },
				new Fixed{ Name = "C", Quantity = 3, Price = 505 },
				new Fixed{ Name = "D", Quantity = 4, Price = 99 },
				new Fixed{ Name = "E", Quantity = 5, Price = 199 },
				new Fixed{ Name = "F", Quantity = 6, Price = 0 },
				new Fixed{ Name = "G", Quantity = 7, Price = 7 }
			};
			using (var doc = factory.Open("WordTables.docx"))
			{
				doc.Process(
					new
					{
						Table1 = dtData,
						Table2 = dtData,
						DynamicResize = dynamicResize1,
						DynamicResizeAndMerge = dynamicResize2,
						Nulls = map,
						Table4 = dtEmpty,
						Table5 = dtEmpty,
						Combined = combined,
						Fixed = fixedItems
					});
			}
			Process.Start(new ProcessStartInfo("WordTables.docx") { UseShellExecute = true });
		}

		class Combined
		{
			public Beer[] Beers;
			public string[,] Headers;
		}
		class Beer
		{
			public string Name;
			public string Description;
			public string[,] Columns;
		}
		class Fixed
		{
			public string Name;
			public int Quantity;
			public decimal Price;
		}
	}
}
