using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Diff.Tool
{
	public class TortoiseGitMerge
	{
		public static void Compare(string beforeFilepath, string afterFilepath)
		{
			Process.Start(GetExePath(), $"\"{beforeFilepath}\" \"{afterFilepath}\"");
		}

		private static string GetExePath()
		{
			return GetPathInProgramFiles(@"TortoiseGit\bin\TortoiseGitMerge.exe");
		}

		private static string GetPathInProgramFiles(string path)
		{
			var programFilesPaths = GetProgramFilesPaths();
			return GetFirstWorking(path, programFilesPaths);
		}

		private static string[] GetProgramFilesPaths()
		{
			return new HashSet<string>
			{
				Environment.GetEnvironmentVariable("ProgramFiles(x86)"),
				Environment.GetEnvironmentVariable("ProgramFiles"),
				Environment.GetEnvironmentVariable("ProgramW6432")
			}.Where(p => p != null).ToArray();
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