using System;
using System.IO;

namespace Diff.Tool
{
	public class Pandoc
	{
		public static string GetExePath()
		{
			return GetFirstWorking("pandoc.exe", GetEnviromentPaths());
		}

		private static string[] GetEnviromentPaths()
		{
			return Environment.GetEnvironmentVariable("PATH")?.Split(';') ?? new string[0];
		}

		private static string GetFirstWorking(string path, string[] paths)
		{
			var path1 = path;
			foreach(var path2 in paths)
			{
				path1 = Path.Combine(path2, path);
				if(File.Exists(path1))
					break;
			}
			return path1;
		}
	}
}