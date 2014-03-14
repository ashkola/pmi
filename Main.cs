﻿using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MailForALM;
using Microsoft.Office.Interop.Word;
using NotesFor.HtmlToOpenXml;
using PMI.Properties;
using TDAPIOLELib;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using List = TDAPIOLELib.List;

namespace PMI
{
	class Main
	{
		private const bool G_VISIBLE_WORD = true;
		private const bool G_IS_FULL_REPORT = true;
		public string Value;
		private string _docPath;

		public string DoWork()
		{
			var xDic = new List<List<string>>();
			var tdConnection = new TDAPIOLELib.TDConnection( );
//			Config.Get();
			try
			{
				tdConnection.InitConnectionEx(Config._server);
				tdConnection.ConnectProjectEx(Config._domain, Config._project, Config._login, Config._password);
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message, "Ошибка подключения", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
//			tdConnection.IgnoreHtmlFormat = true;
			var testF = tdConnection.TestFactory;
			var testFilter = testF.Filter;
			testFilter.Filter["TS_TYPE"] = "MANUAL";
			testFilter.Filter["TS_SUBJECT"] = Config._root;
			testFilter.Order["TS_SUBJECT"] = 1;
			testFilter.OrderDirection["TS_SUBJECT"] = tagTDAPI_FILTERORDER.TDOLE_ASCENDING;
			testFilter.Order["TS_NAME"] = 2;
			testFilter.OrderDirection["TS_SUBJECT"] = tagTDAPI_FILTERORDER.TDOLE_ASCENDING;
			List testList = testFilter.NewList;
			this.Value = testList[1].Name;
			foreach (Test testObj in testList)
			{
				setEvaluatedSteps(testObj, testObj, ref xDic);
			}
			createWord(xDic);
			return _docPath;
		}

		private void setEvaluatedSteps(Test rootTest, Test testObj, ref List<List<string>> resultList)
		{
			var stepF = testObj.DesignStepFactory;
			var stepFilter = stepF.Filter;
			stepFilter.Order["DS_STEP_ORDER"] = 1;
			stepFilter.OrderDirection["DS_STEP_ORDER"] = tagTDAPI_FILTERORDER.TDOLE_ASCENDING;
			string testAttachPath = "";
			string stepAttachPath = "";
			if (Config._attachment)
			{
				List testAttachList = testObj.Attachments.NewList("");
				foreach (Attachment testAttach in testAttachList)
				{
					testAttach.Load(false, "");
					testAttachPath = testAttach.FileName + ";" + testAttachPath;
				}
			}
			List stepList = stepFilter.NewList;			
			foreach (DesignStep stepObj in stepList)
			{
				if (stepObj.LinkTestID != 0)
				{
					setEvaluatedSteps(rootTest, stepObj.LinkTest, ref resultList);
				} else
				{
					var line = new List<string>
					           	{
					           		stepObj.StepName,
					           		stepObj.EvaluatedStepDescription,
					           		stepObj.EvaluatedStepExpectedResult,
					           		"",
					           		rootTest.ID.ToString(),
					           		rootTest.Name,
					           		rootTest["TS_SUBJECT"].Path,
									rootTest["TS_DESCRIPTION"] ?? "",
					           		testAttachPath
					           	};
					if (Config._attachment)
					{
						List stepAttachList = stepObj.Attachments.NewList("");
						foreach (Attachment stepAttach in stepAttachList)
						{
							stepAttach.Load(false, "");
							stepAttachPath = stepAttach.FileName + ";" + stepAttachPath;
						}
						line.Add(stepAttachPath);
					}
					resultList.Add(line);
				}
			}
			return;
		}

		private void createWord(List<List<string>> iDic)
		{
			var word = new Microsoft.Office.Interop.Word.Application();
			word.Visible = G_VISIBLE_WORD;
			var doc = word.Documents.Add();
			doc.Save();
			_docPath = doc.FullName;
			var xStartRow = 0;
			var xEndRow = 0;
			var xCurrentTestId = iDic[0][4] ?? "";
			var xPreviousTreePosition = "";
			int tblCount = 1;

			iDic.Add(new List<string>
			                   	{
			                   		"","","","","","","","","",""
			                   	});
			for(int i=0;i<iDic.Count;i++)
			{
				if(iDic[i][4] == xCurrentTestId)
				{
					xEndRow = i;
				} else
				{
					if (i > 0)
					{
//						Word Jobbing
						int rowsCnt = xEndRow - xStartRow + 1;
						word.Selection.TypeParagraph();
//						Пишем главу, построение дерева
						string xTreePosition = iDic[i - 1][6];
						if (!xTreePosition.Equals(xPreviousTreePosition))
						{
							var xNowArr = new List<string>(xTreePosition.Split('\\'));
							var xBeforeArr = new List<string>(xPreviousTreePosition.Split('\\'));
							for(int level = 1; level < xNowArr.Count(); level++)
							{
								if(xBeforeArr.Count < xNowArr.Count)
								{
									xBeforeArr.Add("");	
								}
								if (!xNowArr[level].Equals(xBeforeArr[level]))
								{	
									word.Selection.set_Style(getStyle(level));
									word.Selection.TypeText(xNowArr[level]);
									word.Selection.TypeParagraph();
								}
							}
							xPreviousTreePosition = xTreePosition;
						}
						if (G_IS_FULL_REPORT)
						{
//							'Название теста-кейса
							word.Selection.set_Style(getStyle(5));
							word.Selection.TypeText(iDic[i - 1][5]);
							word.Selection.TypeParagraph();
//							'Предусловия
							word.Selection.set_Style(getStyle(0));
							word.Selection.TypeText(StripHTML(iDic[i - 1][7]));
//							word.Selection.Text = htmlToText(iDic[i - 1][7]);

							word.Selection.TypeParagraph();
							if (Config._attachment)
							{
								if (iDic[i - 1][8] != "")
								{
									word.Selection.TypeText("Вложения:");
									word.Selection.TypeParagraph();
									foreach (string xFile in iDic[i - 1][8].Split(';'))
									{
										if (!xFile.Equals(""))
										{
											var xArr = xFile.Split('\\');
											word.Selection.InlineShapes.AddOLEObject(null, xFile, false, true, null, 1, xArr[xArr.Count() - 1]);
										}
									}
								}
							}
//							'Рисуем таблицу
							word.Selection.Tables.Add(word.Selection.Range, rowsCnt + 1, 3, 1);
							word.Selection.Tables[1].Select();
							word.Selection.set_Style(getStyle(0));
							word.Selection.MoveDown();

							var table = word.ActiveDocument.Tables[tblCount];
							table.Columns[1].Width = 50;
							table.Columns[2].Width = 200;
							table.Columns[3].Width = 200;

							table.Rows[1].Select();
							word.Selection.Font.Bold = 9999998;

							table.Cell(1, 1).Select();

//							Заголовок таблицы
							word.Selection.TypeText("Номер шага");
							table.Cell(1, 2).Select();
							word.Selection.TypeText("Действие в системе");
							table.Cell(1, 3).Select();
							word.Selection.TypeText("Ожидаемая реакция системы");

//							Внесение данных в таблицы
							for (int row = 2; row <= rowsCnt + 1; row++)
							{
								table.Cell(row, 1).Select();
								word.Selection.TypeText((row - 1).ToString() + ".");

								if (Config._attachment)
								{
									if (iDic[row - 2][9] != "")
									{
										word.Selection.TypeParagraph();
										foreach (string xFile in iDic[row - 2][9].Split(';'))
										{
											if (!xFile.Equals(""))
											{
												var xArr = xFile.Split('\\');
												word.Selection.InlineShapes.AddOLEObject(null, xFile, false, true, null, 1, xArr[xArr.Count() - 1]);
											}
										}
									}
								}
								for (int colInRow = 2; colInRow <= 3; colInRow++)
								{
									table.Cell(row, colInRow).Select();
//									word.Selection.Text = htmlToText(iDic[i - rowsCnt + row - 2][colInRow - 1]);
									word.Selection.Text = StripHTML(iDic[i - rowsCnt + row - 2][colInRow - 1]);
								}
							}
							word.Selection.MoveDown(4, 3);
							tblCount = tblCount + 1;
						} else
						{
//							Название теста-кейса
							word.Selection.set_Style(getStyle(0));
							word.Selection.TypeText("Тест-кейс: " + iDic[i - 1][5]);
						}
//						Периодически перезаркрываю док, для очистки памяти
						if(tblCount % 20 == 0)
						{
							doc.Save();
							doc.Close();
							doc = word.Documents.Open(_docPath);
							word.Selection.EndKey(6);
						}
					}
//					Ставим курсоры на новый блок
					xStartRow = i;
					xEndRow = i;
					xCurrentTestId = iDic[i][4];
				}
			}
			doc.Save();
			return;
		}

		private WdBuiltinStyle getStyle(int level)
		{
			switch (level)
			{
				case 1:
					return WdBuiltinStyle.wdStyleHeading1;
				case 2:
					return WdBuiltinStyle.wdStyleHeading2;
				case 3:
					return WdBuiltinStyle.wdStyleHeading3;
				case 4:
					return WdBuiltinStyle.wdStyleHeading4;
				case 5:
					return WdBuiltinStyle.wdStyleHeading5;
				case 6:
					return WdBuiltinStyle.wdStyleHeading6;
				default:
					return WdBuiltinStyle.wdStyleDefaultParagraphFont;
			}
		}

		public string StripHTML(string source)
		{
			try
			{
				string result;

				// Remove HTML Development formatting
				// Replace line breaks with space
				// because browsers inserts space
				result = source.Replace("\r", " ");
				// Replace line breaks with space
				// because browsers inserts space
				result = result.Replace("\n", " ");
				// Remove step-formatting
				result = result.Replace("\t", string.Empty);
				// Remove repeating spaces because browsers ignore them
				result = System.Text.RegularExpressions.Regex.Replace(result,
																	  @"( )+", " ");

				// Remove the header (prepare first by clearing attributes)
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*head([^>])*>", "<head>",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"(<( )*(/)( )*head( )*>)", "</head>",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(<head>).*(</head>)", string.Empty,
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// remove all scripts (prepare first by clearing attributes)
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*script([^>])*>", "<script>",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"(<( )*(/)( )*script( )*>)", "</script>",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				//result = System.Text.RegularExpressions.Regex.Replace(result,
				//         @"(<script>)([^(<script>\.</script>)])*(</script>)",
				//         string.Empty,
				//         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"(<script>).*(</script>)", string.Empty,
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// remove all styles (prepare first by clearing attributes)
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*style([^>])*>", "<style>",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"(<( )*(/)( )*style( )*>)", "</style>",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(<style>).*(</style>)", string.Empty,
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// insert tabs in spaces of <td> tags
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*td([^>])*>", "\t",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// insert line breaks in places of <BR> and <LI> tags
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*br( )*>", "\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*li( )*>", "\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// insert line paragraphs (double line breaks) in place
				// if <P>, <DIV> and <TR> tags
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*div([^>])*>", "\r\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*tr([^>])*>", "\r\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<( )*p([^>])*>", "\r\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// Remove remaining tags like <a>, links, images,
				// comments etc - anything that's enclosed inside < >
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"<[^>]*>", string.Empty,
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// replace special characters:
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @" ", " ",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&bull;", " * ",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&lsaquo;", "<",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&rsaquo;", ">",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&trade;", "(tm)",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&frasl;", "/",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&lt;", "<",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&gt;", ">",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&copy;", "(c)",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&reg;", "(r)",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				// Remove all others. More can be added, see
				// http://hotwired.lycos.com/webmonkey/reference/special_characters/
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 @"&(.{2,6});", string.Empty,
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// for testing
				//System.Text.RegularExpressions.Regex.Replace(result,
				//       this.txtRegex.Text,string.Empty,
				//       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

				// make line breaking consistent
				result = result.Replace("\n", "\r");

				// Remove extra line breaks and tabs:
				// replace over 2 breaks with 2 and over 4 tabs with 4.
				// Prepare first to remove any whitespaces in between
				// the escaped characters and remove redundant tabs in between line breaks
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(\r)( )+(\r)", "\r\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(\t)( )+(\t)", "\t\t",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(\t)( )+(\r)", "\t\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(\r)( )+(\t)", "\r\t",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				// Remove redundant tabs
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(\r)(\t)+(\r)", "\r\r",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				// Remove multiple tabs following a line break with just one tab
				result = System.Text.RegularExpressions.Regex.Replace(result,
						 "(\r)(\t)+", "\r\t",
						 System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				// Initial replacement target string for line breaks
				string breaks = "\r\r\r";
				// Initial replacement target string for tabs
				string tabs = "\t\t\t\t\t";
				for (int index = 0; index < result.Length; index++)
				{
					result = result.Replace(breaks, "\r\r");
					result = result.Replace(tabs, "\t\t\t\t");
					breaks = breaks + "\r";
					tabs = tabs + "\t";
				}

				// That's it.
				return result;
			}
			catch
			{
//				MessageBox.Show("Error");
				return source;
			}
		}

		public string htmlToText(string sHTML)
		{
			
			var browser = new WebBrowser();
			browser.DocumentText = sHTML;
			do
			{
				System.Windows.Forms.Application.DoEvents();
			} while (browser.ReadyState != WebBrowserReadyState.Complete);
			return browser.Document.Body.OuterText;
		}

//		private string htmlToText2(string sHTML)
//		{
//
//			var HtmlDocument = new H
//			browser.DocumentText = sHTML;
//			do
//			{
//				System.Windows.Forms.Application.DoEvents();
//			} while (browser.ReadyState != WebBrowserReadyState.Complete);
//			return browser.Document.Body.OuterText;
//		}


		public void DoOpenXml(string str)
		{
			const string filename = "test.docx";
//			string html = File.ReadAllText("c:\\test.dat");
			string html = Resources.DemoHtml;
			if (File.Exists(filename)) File.Delete(filename);

			using (var generatedDocument = new MemoryStream())
			{
				using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
				{
					MainDocumentPart mainPart = package.MainDocumentPart;
					if (mainPart == null)
					{
						mainPart = package.AddMainDocumentPart();
						new Document(new Body()).Save(mainPart);
					}

					var converter = new HtmlConverter(mainPart);
					var body = mainPart.Document.Body;
					var paragraphs = converter.Parse(html);
					foreach (OpenXmlCompositeElement t in paragraphs)
					{
						body.Append(t);
					}

					mainPart.Document.Save();
				}

				File.WriteAllBytes(filename, generatedDocument.ToArray());
			}

			System.Diagnostics.Process.Start(filename);
		}
	}
}