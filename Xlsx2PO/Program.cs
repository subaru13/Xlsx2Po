using MathNet.Numerics;
using MathNet.Numerics.Distributions;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.Fonts;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Xlsx2PO
{
	using MsgTable = Dictionary<string, POConverter.MsgTableRow>;
	using TargetTable = Dictionary<string, FileInfo[]>;
	using TranslationTable = Dictionary<string, Dictionary<string, POConverter.MsgTableRow>>;

	internal class POConverter
	{
		public struct ConvertInfo
		{
			public ConvertInfo()
			{
				Cultures = [];
				Column = string.Empty;
			}
			// 地域コード
			public List<string> Cultures { get; set; }
			// 読み取り列名
			public string Column { get; set; }
			[JsonIgnore]
			public int ColumnIndex { get; set; }
		}

		public struct ProjectInfo
		{
			public ProjectInfo()
			{
				WorkingDirectory = string.Empty;
				SrcDirectory = string.Empty;
				OutDirectory = string.Empty;
				KeysColumn = string.Empty;
				RowHead = 0;
				ConvertTable = [];
				PostActions = [];
			}

			public string WorkingDirectory { get; set; }
			public string SrcDirectory { get; set; }
			public string OutDirectory { get; set; }
			public string KeysColumn { get; set; }
			[JsonIgnore]
			public int KeysColumnIndex { get; set; }
			public int RowHead { get; set; }
			// 変換テーブル
			public Dictionary<string, ConvertInfo> ConvertTable { get; set; }
			// 終了後処理
			public string[]? PostActions { get; set; }
		}

		public struct MsgTableRow
		{
			public string Namespace { get; set; }
			public string Key { get; set; }
			public string Text { get; set; }
		}


		static readonly string DefaultProjectFilePath = "../../../Project.x2p.json";
		static readonly string PoFormat = "msgctxt \"{0}\"\nmsgid \"{1}\"\nmsgstr \"{2}\"\n\n";
		static readonly string IniFormat = "NSLOCTEXT(\"{0}\",\"{1}\",\"{1}\");";
		static readonly string CsvFormat = "\"{0}\",\"{0}\"";

		private static bool CheckColumnName(string s)
		{
			foreach (var c in s)
			{
				if (!((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')))
				{
					return false;
				}
			}
			return s.Length > 0;
		}
		private static int ColumnNameToIndex(string columnName)
		{
			int index = 0;

			foreach (char c in columnName.ToUpper())
			{
				index *= 26;
				index += (c - 'A' + 1);
			}

			return index - 1; // 0始まり
		}

		private static bool ValidationProjectInfo(ref ProjectInfo projectInfo)
		{
			bool isSuccessfully = true;
			if ((!String.IsNullOrEmpty(projectInfo.WorkingDirectory))
				&& (!Directory.Exists(projectInfo.WorkingDirectory)))
			{
				Console.WriteLine($"作業ディレクトリ<{projectInfo.WorkingDirectory}>が存在しません。");
				isSuccessfully = false;
			}

			// 列名に不正な数値が設定されていないか確認
			if (CheckColumnName(projectInfo.KeysColumn))
			{
				// ついでにインデックスに変換しておく
				projectInfo.KeysColumnIndex = ColumnNameToIndex(projectInfo.KeysColumn);
			}
			else
			{
				Console.WriteLine($"キーの列名<{projectInfo.KeysColumn}>に不正な値が設定されています。");
				isSuccessfully = false;
			}

			foreach (var key in projectInfo.ConvertTable.Keys)
			{
				ref var convertTableRow = ref CollectionsMarshal.GetValueRefOrNullRef(projectInfo.ConvertTable, key);
				if (CheckColumnName(convertTableRow.Column))
				{
					// ついでにインデックスに変換しておく
					convertTableRow.ColumnIndex = ColumnNameToIndex(convertTableRow.Column);
				}
				else
				{
					convertTableRow.ColumnIndex = -1;
					Console.WriteLine($"変換デーブル<{key}>の列名<{convertTableRow.Column}>に不正な値が設定されています。");
				}
				if (convertTableRow.Cultures.Count == 0)
				{
					Console.WriteLine($"変換デーブル<{key}>の出力カルチャーが設定されていません");
				}
			}

			return isSuccessfully;
		}

		private static bool ReadProjectInfo(string projectFilePath, ref ProjectInfo projectInfo)
		{
			if (!File.Exists(projectFilePath))
			{
				Console.WriteLine($"プロジェクトファイル<\"{projectFilePath}\">が存在しません");
				return false;
			}

			string jsonString = File.ReadAllText(projectFilePath);
			if (jsonString.Length == 0)
			{
				return false;
			}

			projectInfo = JsonSerializer.Deserialize<ProjectInfo>(jsonString);

			if (!String.IsNullOrEmpty(projectInfo.WorkingDirectory))
			{
				projectInfo.WorkingDirectory = Environment.ExpandEnvironmentVariables(projectInfo.WorkingDirectory);
			}

			if (String.IsNullOrEmpty(projectInfo.SrcDirectory))
			{
				projectInfo.SrcDirectory = "./Src";
			}

			if (String.IsNullOrEmpty(projectInfo.OutDirectory))
			{
				projectInfo.OutDirectory = "./Out";
			}

			projectInfo.SrcDirectory = Environment.ExpandEnvironmentVariables(projectInfo.SrcDirectory);
			projectInfo.OutDirectory = Environment.ExpandEnvironmentVariables(projectInfo.OutDirectory);

			if (String.IsNullOrEmpty(projectInfo.KeysColumn))
			{
				projectInfo.KeysColumn = "A";
			}

			// エクセル上での行番号なので0スタートの数値に変換しておく
			projectInfo.RowHead -= 1;

			return ValidationProjectInfo(ref projectInfo);
		}


		private static TargetTable CreateTargetTable(ProjectInfo projectInfo)
		{
			var targetTable = new TargetTable();

			var srcDirectory = new DirectoryInfo(projectInfo.SrcDirectory);
			var subDirs = srcDirectory.GetDirectories();
			foreach (var subDir in subDirs)
			{
				// ディレクトリ名がターゲット名として機能する
				targetTable.Add(subDir.Name, subDir.GetFiles("*.xlsx"));
			}

			return targetTable;
		}


		private static void ReadMsgTable(
			ISheet sheet,
			int keysColumnIndex,
			int textColumnIndex,
			int rowHead,
			string msgNamespace,
			ref MsgTable msgTable)
		{
			for (int i = rowHead; i < sheet.LastRowNum; i++)
			{
				IRow row = sheet.GetRow(i);

				if (row == null)
				{
					continue;
				}

				ICell keyCell = row.GetCell(keysColumnIndex);
				if (keyCell == null)
				{
					continue;
				}
				var key = keyCell.ToString();
				if (String.IsNullOrEmpty(key))
				{
					continue;
				}

				ICell textCell = row.GetCell(textColumnIndex);
				string? text = textCell?.ToString();

				var msgTableRow = new MsgTableRow
				{
					Namespace = msgNamespace,
					Key = key,
					Text = text ?? string.Empty
				};

				if (!msgTable.TryAdd($"{msgNamespace},{key}", msgTableRow))
				{
					Console.WriteLine($"キーの重複 namespace = {msgNamespace} key = {key}");
				}
			}
		}

		private static void ReadTranslationTable(
			ProjectInfo projectInfo,
			FileInfo fileInfo,
			ref TranslationTable translationTable)
		{
			string msgNamespace = Path.ChangeExtension(fileInfo.Name, null);


			try
			{
				Console.WriteLine($"\"{fileInfo.Name}\" : 読み込み開始");

				// xlsx ファイルを読み込む
				var workbook = WorkbookFactory.Create(fileInfo.FullName);
				var sheet = workbook.GetSheetAt(0);

				foreach (var convertTableRow in projectInfo.ConvertTable)
				{
					if (convertTableRow.Value.ColumnIndex < 0)
					{
						return;
					}

					// 要素を取得、なければ作成して参照を取得
					ref var msgTable = ref CollectionsMarshal.GetValueRefOrAddDefault(
						translationTable,
						convertTableRow.Key,
						out _);

					msgTable ??= [];

					ReadMsgTable(
						sheet,
						projectInfo.KeysColumnIndex,
						convertTableRow.Value.ColumnIndex,
						projectInfo.RowHead,
						msgNamespace,
						ref msgTable);
				}

				Console.WriteLine($"\"{fileInfo.Name}\" : 読み込み終了");
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}

		private static void ExportMsgTable(string poFilePath, MsgTable msgTable)
		{
			using StreamWriter poWriter = new(poFilePath, false, Encoding.UTF8);
			foreach (var msgTableRow in msgTable)
			{
				string msgText = msgTableRow.Value.Text.Replace("\n", "\\n");

				poWriter.Write(string.Format(PoFormat, msgTableRow.Key, msgTableRow.Value.Key, msgText));
			}
		}

		private static void ExportGatherFile(
		   string iniFilePath,
		   MsgTable msgTable)
		{

			using StreamWriter iniWriter = new(iniFilePath, false, Encoding.UTF8);
			foreach (var msgTableRow in msgTable.Values)
			{
				string msgNamespace = msgTableRow.Namespace;
				string msgKey = msgTableRow.Key;

				iniWriter.WriteLine(String.Format(IniFormat, msgNamespace, msgKey));
			}
		}

		private static Dictionary<string, List<string>> CreateStringTables(
		   MsgTable msgTable)
		{
			Dictionary<string, List<string>> stringTables = [];

			foreach (var msgTableRow in msgTable.Values)
			{
				// 要素を取得、なければ作成して参照を取得
				ref var stringTable = ref CollectionsMarshal.GetValueRefOrAddDefault(
					stringTables,
					msgTableRow.Namespace,
					out _);

				stringTable ??= [];

				stringTable.Add(msgTableRow.Key);
			}

			return stringTables;
		}

		private static void ExportStringTable(
		   string csvFilePath,
		   List<string> keys)
		{
			using StreamWriter csvWriter = new(csvFilePath, false, Encoding.UTF8);
			csvWriter.WriteLine("Keys,SourceString");
			foreach (var key in keys)
			{
				csvWriter.WriteLine(String.Format(CsvFormat, key));
			}
		}

		private static void ExportNativeMsgTables(
			string outDir,
			string targetName,
			TranslationTable translationTable)
		{
			if (!Directory.Exists(outDir))
			{
				Directory.CreateDirectory(outDir);
			}
			var msgTable = translationTable.First().Value;
			string iniFilePath = Path.Combine(outDir, $"{targetName}.ini");
			ExportGatherFile(iniFilePath, msgTable);

			var stringTables = CreateStringTables(msgTable);

			foreach (var stringTable in stringTables)
			{
				string csvFilePath = Path.Combine(outDir, $"{stringTable.Key}.csv");
				ExportStringTable(csvFilePath, stringTable.Value);
			}
		}

		private static void ExportTranslationTable(
			ProjectInfo projectInfo,
			string targetName,
			TranslationTable translationTable)
		{
			string targetOutDir = Path.Combine(projectInfo.OutDirectory, targetName);
			string poFileName = string.Format("{0}.po", targetName);

			foreach (var convertTableRow in projectInfo.ConvertTable)
			{
				if (!translationTable.TryGetValue(convertTableRow.Key, out var msgTable))
				{
					Console.WriteLine($"出力用テーブルが存在しません target = {targetName} key = {convertTableRow.Key}");
					continue;
				}
				foreach (var culture in convertTableRow.Value.Cultures)
				{
					string cultureDir = Path.Combine(targetOutDir, culture);
					if (!Directory.Exists(cultureDir))
					{
						Directory.CreateDirectory(cultureDir);
					}
					string poFilePath = Path.Combine(cultureDir, poFileName);

					ExportMsgTable(poFilePath, msgTable);
				}
			}

			ExportNativeMsgTables(targetOutDir, targetName, translationTable);
		}

		private static void RunPostAction(string postAction)
		{
			var match = Regex.Match(postAction, @" (?=(?:(?:[^""]*""){2})*[^""]*$)");
			string Args = string.Empty;
			string FileName;
			if (match.Success)
			{
				FileName = postAction[..match.Index];
				Args = postAction[(match.Index + match.Length)..];
			}
			else
			{
				FileName = postAction;
			}

			var psi = new ProcessStartInfo()
			{
				FileName = FileName,
				Arguments = Args,
				UseShellExecute = false,
				RedirectStandardOutput = true,
				CreateNoWindow = true,
			};

			try
			{
				using var proc = Process.Start(psi);

				if (proc == null)
				{
					return;
				}

				// 出力結果をすべて読み込む
				string output = proc.StandardOutput.ReadToEnd();
				proc.WaitForExit(); // プロセス終了を待機
				Console.WriteLine(output);
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Error : PostAction<{postAction}>");
				Console.WriteLine(ex);
			}
		}

		public static void Execute(string projectFilePath)
		{
			// プロジェクトファイルの読み込み
			var projectInfo = new ProjectInfo();
			if (!ReadProjectInfo(projectFilePath, ref projectInfo))
			{
				return;
			}

			using var scopedWorkingDirectories = new ScopedWorkingDirectories(projectInfo.WorkingDirectory);

			// 出力先をクリアする
			if (Directory.Exists(projectInfo.OutDirectory))
			{
				Directory.Delete(projectInfo.OutDirectory, true);
			}

			// ターゲットテーブルの作成
			var targetTable = CreateTargetTable(projectInfo);

			if (targetTable.Count == 0)
			{
				Console.WriteLine("ターゲットがありません");
				return;
			}

			// xlsx -> po の変換
			foreach (var target in targetTable)
			{
				TranslationTable translationTable = [];
				foreach (var srcFile in target.Value)
				{
					ReadTranslationTable(
						projectInfo,
						srcFile,
						ref translationTable);
				}
				ExportTranslationTable(projectInfo, target.Key, translationTable);
			}

			if (projectInfo.PostActions != null)
			{
				foreach (var PostAction in projectInfo.PostActions)
				{
					RunPostAction(PostAction);
				}
			}
		}

		public static void Main(string[] Args)
		{
			if (Args.Length > 0)
			{
				foreach (var arg in Args)
				{
					// プロジェクトファイルなら実行
					if (arg.EndsWith(".x2p.json", StringComparison.OrdinalIgnoreCase))
					{
						Execute(arg);
					}
				}
			}
			else
			{
				Execute(DefaultProjectFilePath);
			}
		}

	}


}
