using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using BarTender;
using Excel;
using LabelManager2;

namespace PrintLabel
{
	// Token: 0x02000002 RID: 2
	public class Setup
	{
		// Token: 0x06000001 RID: 1 RVA: 0x0000214C File Offset: 0x0000034C
		public bool GetPrintData(string sType, ref System.Windows.Forms.ListBox ListParam, ref System.Windows.Forms.ListBox ListData)
		{
			DataSet dataSet = ClientUtils.ExecuteSQL(" select * from sajet.s_sys_print_data  where data_type ='" + sType + "' order by data_order ");
			for (int i = 0; i <= dataSet.Tables[0].Rows.Count - 1; i++)
			{
				string text = dataSet.Tables[0].Rows[i]["DATA_SQL"].ToString().ToUpper();
				string text2 = dataSet.Tables[0].Rows[i]["INPUT_PARAM"].ToString().ToUpper().TrimEnd(new char[]
				{
					';'
				});
				string text3 = dataSet.Tables[0].Rows[i]["INPUT_FIELD"].ToString().ToUpper().TrimEnd(new char[]
				{
					';'
				});
				string text4 = dataSet.Tables[0].Rows[i]["OUTPUT_PARAM"].ToString().ToUpper().TrimEnd(new char[]
				{
					';'
				});
				string[] array = text2.Split(new char[]
				{
					';'
				});
				string[] array2 = text3.Split(new char[]
				{
					';'
				});
				string[] array3 = text4.Split(new char[]
				{
					';'
				});
				string text5 = dataSet.Tables[0].Rows[i]["DATA_SQL2"].ToString().ToUpper();
				string text6 = dataSet.Tables[0].Rows[i]["INPUT_PARAM2"].ToString().ToUpper().TrimEnd(new char[]
				{
					';'
				});
				string text7 = dataSet.Tables[0].Rows[i]["INPUT_FIELD2"].ToString().ToUpper().TrimEnd(new char[]
				{
					';'
				});
				string text8 = dataSet.Tables[0].Rows[i]["OUTPUT_PARAM2"].ToString().ToUpper().TrimEnd(new char[]
				{
					';'
				});
				string[] array4 = text6.Split(new char[]
				{
					';'
				});
				string[] array5 = text7.Split(new char[]
				{
					';'
				});
				string[] array6 = text8.Split(new char[]
				{
					';'
				});
				for (int j = 0; j <= array.Length - 1; j++)
				{
					string text9 = array[j].ToString();
					string text10 = array2[j].ToString();
					string str;
					if (i == 0)
					{
						if (j > ListData.Items.Count - 1)
						{
							ListData.Items.Add(ListData.Items[0].ToString());
						}
						str = ListData.Items[j].ToString();
						ListParam.Items.Add(text10);
					}
					else
					{
						str = this.Get_ParamData(text10, ListParam, ListData);
					}
					if (text.IndexOf(text9) >= 0)
					{
						text = text.Replace(text9, " '" + str + "' ");
					}
				}
				DataSet dataSet2 = ClientUtils.ExecuteSQL(text);
				for (int k = 0; k <= dataSet2.Tables[0].Rows.Count - 1; k++)
				{
					for (int l = 0; l <= dataSet2.Tables[0].Columns.Count - 1; l++)
					{
						string text11 = dataSet2.Tables[0].Columns[l].ColumnName.ToString();
						if (text4 == "")
						{
							if (ListParam.Items.IndexOf(text11) == -1)
							{
								ListParam.Items.Add(text11);
								ListData.Items.Add(dataSet2.Tables[0].Rows[k][text11].ToString());
							}
						}
						else if (l < array3.Length)
						{
							string text12 = array3[l] + Convert.ToString(k + 1);
							if (ListParam.Items.IndexOf(text12) == -1)
							{
								ListParam.Items.Add(text12);
								ListData.Items.Add(dataSet2.Tables[0].Rows[k][text11].ToString());
							}
							if (k == dataSet2.Tables[0].Rows.Count - 1)
							{
								if (ListParam.Items.IndexOf(array3[l] + "END") == -1)
								{
									ListParam.Items.Add(array3[l] + "END");
									ListData.Items.Add(dataSet2.Tables[0].Rows[k][text11].ToString());
								}
								if (ListParam.Items.IndexOf(array3[l] + "COUNT") == -1)
								{
									ListParam.Items.Add(array3[l] + "COUNT");
									ListData.Items.Add(dataSet2.Tables[0].Rows.Count.ToString());
								}
							}
						}
					}
					if (text8 != "")
					{
						string text13 = text5;
						for (int m = 0; m <= array4.Length - 1; m++)
						{
							string text14 = array4[m].ToString();
							string columnName = array5[m].ToString();
							string str2 = dataSet2.Tables[0].Rows[k][columnName].ToString();
							if (text13.IndexOf(text14) >= 0)
							{
								text13 = text13.Replace(text14, " '" + str2 + "' ");
							}
						}
						DataSet dataSet3 = ClientUtils.ExecuteSQL(text13);
						for (int n = 0; n <= dataSet3.Tables[0].Rows.Count - 1; n++)
						{
							for (int num = 0; num <= dataSet3.Tables[0].Columns.Count - 1; num++)
							{
								if (num < array6.Length)
								{
									string columnName2 = dataSet3.Tables[0].Columns[num].ColumnName.ToString();
									string text15 = array6[num] + Convert.ToString(k + 1) + "_" + Convert.ToString(n + 1);
									if (ListParam.Items.IndexOf(text15) == -1)
									{
										ListParam.Items.Add(text15);
										ListData.Items.Add(dataSet3.Tables[0].Rows[n][columnName2].ToString());
									}
								}
							}
						}
					}
				}
			}
			return true;
		}

		// Token: 0x06000002 RID: 2 RVA: 0x00002888 File Offset: 0x00000A88
		private string Get_ParamData(string sParam, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData)
		{
			int num = ListParam.Items.IndexOf(sParam.ToUpper());
			string result;
			if (num < 0)
			{
				result = "";
			}
			else
			{
				result = ListData.Items[num].ToString();
			}
			return result;
		}

		// Token: 0x06000003 RID: 3 RVA: 0x000028C8 File Offset: 0x00000AC8
		private bool Get_Sample_File(string sExeName, string sFileTitle, string sFileName, string sPrintMethod, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sSampleFile)
		{
			bool s = true;
			DataSet dataSet = ClientUtils.ExecuteSQL("select lablel_part_id from sajet.g_part_lablel WHERE FIELD_TYPE = 'LABLE'");
			string[] part = dataSet.Tables[0].Rows[0]["LABLEL_PART_ID"].ToString().Split(',');
			string str = System.Windows.Forms.Application.StartupPath + "\\" + sExeName + "\\";
			if (Directory.Exists(str + "PrintLabel\\"))
			{
				str += "PrintLabel\\";
			}
			sSampleFile = "";
			string text;
			if (sPrintMethod.ToUpper() == "CODESOFT")
			{
				text = ".LAB";
			}
			else if (sPrintMethod.ToUpper() == "BARTENDER")
			{
				text = ".btw";
			}
			else if (sPrintPort.ToUpper() == "EXCEL")
			{
				text = ".xlt";
			}
			else
			{
				text = ".txt";
			}
			if (sFileName != "")
			{
				sSampleFile = str + sFileName + text;
				if (!File.Exists(sSampleFile))
				{
					return false;
				}
			}
			else
			{
				sSampleFile = str + sFileTitle + this.Get_ParamData("LABEL_FILE", ListParam, ListData) + text;
				if (!File.Exists(sSampleFile))
				{
					for (int i = 0; i < part.Length; i++)
					{

						if (part[i].ToString() == this.Get_ParamData("PART_NO", ListParam, ListData).Substring(0, part[i].Length) && part[i] != null)
						{
							int g = 0;
							sSampleFile = str + sFileTitle + part[i].ToString() + text;
							while (File.Exists(sSampleFile))
							{
								ts.Add(sSampleFile);
								g++;
								sSampleFile = str + sFileTitle + part[i].ToString() + "(" + g + ")" + text;
							}
						}
					}
				}
				if (!File.Exists(sSampleFile))
				{
					sSampleFile = str + sFileTitle + this.Get_ParamData("PART_NO", ListParam, ListData) + text;
					ts.Add(sSampleFile);
					for (int d = 0; d < 100; d++)
					{
						sSampleFile = str + sFileTitle + this.Get_ParamData("PART_NO", ListParam, ListData)+"("+d+")" + text;
						if(File.Exists(sSampleFile))
						{
							ts.Add(sSampleFile);
						}
					}
				}
				if (!File.Exists(sSampleFile))
				{
					sSampleFile = str + sFileTitle + this.Get_ParamData("WORK_ORDER", ListParam, ListData) + text;
					ts.Add(sSampleFile);
				}
				if (!File.Exists(sSampleFile))
				{
					sSampleFile = str + sFileTitle + this.Get_ParamData("CUSTOMER_CODE", ListParam, ListData) + text;
					ts.Add(sSampleFile);
				}
				if (!File.Exists(sSampleFile))
				{
					sSampleFile = str + sFileTitle + this.Get_ParamData("MODEL_NAME", ListParam, ListData) + text;
					ts.Add(sSampleFile);
				}
				if (!File.Exists(sSampleFile))
				{
					sSampleFile = str + sFileTitle + "DEFAULT" + text;
					
				}
				if (ts.Count < 1)
				{
					return false;
				}
			}
			return true;
		}

		// Token: 0x06000004 RID: 4 RVA: 0x00002A60 File Offset: 0x00000C60
		private string Get_Bartender_Split_Symbol(string sType)
		{
			DataSet dataSet = ClientUtils.ExecuteSQL(" select PARAM_VALUE from SAJET.SYS_BASE  where PROGRAM='ALL'    AND PARAM_NAME ='" + sType + "'    AND ROWNUM = 1 ");
			string result;
			if (dataSet.Tables[0].Rows.Count > 0)
			{
				result = dataSet.Tables[0].Rows[0]["PARAM_VALUE"].ToString();
			}
			else
			{
				result = ",";
			}
			return result;
		}
		// Token: 0x06000005 RID: 5 RVA: 0x00002ACC File Offset: 0x00000CCC
		public bool Print_Bartender_DataSource(string sExeName, string sType, string sFileTitle, string sFileName, int iPrintQty, string sPrintMethod, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			ts = new List<string>();
			sMessage = "OK";
			string text = "";
			System.Windows.Forms.ListBox listBox = new System.Windows.Forms.ListBox();
			listBox.Items.Add(ListData.Items[0]);
			ListParam.Items.Clear();
			this.GetPrintData(sType, ref ListParam, ref listBox);
			bool result = true;
			if (!this.Get_Sample_File(sExeName, sFileTitle, sFileName, sPrintMethod, sPrintPort, ListParam, listBox, out text))
			{
				sMessage = string.Concat(new string[]
				{
					"Sample File not Exist",
					Environment.NewLine,
					"(",
					text,
					")",
					Environment.NewLine,
					"Print Method ",
					sPrintMethod,
					Environment.NewLine,
					"Print Port ",
					sPrintPort
				});
				result = false;
			}
			else
			{
				int P = 0;
				for (int y = 0; y < ts.Count; y++)
				{
					if (File.Exists(ts[y]))
					{
						P++;
						string startupPath = System.Windows.Forms.Application.StartupPath;
						FileInfo fileInfo = new FileInfo(ts[y]);
						fileInfo.Extension.ToString().ToUpper();
						string directoryName = fileInfo.DirectoryName;
						sFileName = Path.GetFileNameWithoutExtension(ts[y]);
						string text2 = ts[y];
						string text3 = string.Concat(new string[]
						{
					startupPath,
					"\\",
					sExeName,
					"\\",
					sFileName,
					".lst"
						});
						string text4 = string.Concat(new string[]
						{
					startupPath,
					"\\",
					sExeName,
					"\\",
					sFileName,
					".dat"
						});
						string text5 = string.Concat(new string[]
						{
					startupPath,
					"\\",
					sExeName,
					"\\",
					sFileTitle,
					"DEFAULT.dat"
						});
						string text6 = startupPath + "\\" + sExeName + "\\PrintGo.bat";
						string sFile = startupPath + "\\" + sExeName + "\\PrintLabel.bat";
						if (!File.Exists(text2))
						{
							sMessage = "Label File Not Found (btw)" + Environment.NewLine + Environment.NewLine + text2;
							result = false;
						}
						else
						{
							if (!File.Exists(text4))
							{
								if (!File.Exists(text5))
								{
									sMessage = string.Concat(new string[]
									{
								"Label File Not Found (.dat)",
								Environment.NewLine,
								Environment.NewLine,
								text4,
								Environment.NewLine,
								Environment.NewLine
									});
									if (text4 != text5)
									{
										sMessage = string.Concat(new string[]
										{
									sMessage,
									"OR ",
									Environment.NewLine,
									Environment.NewLine,
									text5
										});
									}
									return false;
								}
								text4 = text5;
							}
							if (File.Exists(text3))
							{
								File.Delete(text3);
							}
							string text7 = this.Get_Bartender_Split_Symbol(sType);
							System.Windows.Forms.ListBox listBox2 = this.LoadFileHeader(text4, ref sMessage, text7);
							if (!string.IsNullOrEmpty(sMessage))
							{
								result = false;
							}
							else
							{
								string text8 = string.Empty;
								for (int i = 0; i <= listBox2.Items.Count - 1; i++)
								{
									if (text7 == "1")
									{
										text8 = text8 + listBox2.Items[i].ToString() + "\t";
									}
									else
									{
										text8 = text8 + listBox2.Items[i].ToString() + text7;
									}
								}
								if (!string.IsNullOrEmpty(text8))
								{
									if (text7 == "1")
									{
										text8 = text8.TrimEnd(new char[]
										{
									'\t'
										});
									}
									else
									{
										text8 = text8.Substring(0, text8.Length - 1);
									}
								}
								this.WriteToTxt(text3, text8);
								if (sType == "PAGE")
								{
									text8 = string.Empty;
									for (int j = 0; j <= listBox2.Items.Count - 1; j++)
									{
										string value = listBox2.Items[j].ToString();
										int num = ListParam.Items.IndexOf(value);
										if (num >= 0)
										{
											if (text7 == "1")
											{
												text8 = text8 + ListData.Items[num].ToString() + "\t";
											}
											else
											{
												text8 = text8 + ListData.Items[num].ToString() + text7;
											}
										}
										else if (text7 == "1")
										{
											text8 += "\t";
										}
										else
										{
											text8 += text7;
										}
									}
									if (!string.IsNullOrEmpty(text8))
									{
										if (text7 == "1")
										{
											text8 = text8.TrimEnd(new char[]
											{
										'\t'
											});
										}
										else
										{
											text8 = text8.Substring(0, text8.Length - 1);
										}
									}
									this.WriteToTxt(text3, text8);
								}
								else
								{
									for (int k = 0; k <= ListData.Items.Count - 1; k++)
									{
										if (k > 0)
										{
											ListParam.Items.Clear();
											listBox.Items.Clear();
											listBox.Items.Add(ListData.Items[k]);
											this.GetPrintData(sType, ref ListParam, ref listBox);
										}
										text8 = string.Empty;
										for (int l = 0; l <= listBox2.Items.Count - 1; l++)
										{
											string value2 = listBox2.Items[l].ToString();
											int num2 = ListParam.Items.IndexOf(value2);
											if (num2 >= 0)
											{
												if (text7 == "1")
												{
													text8 = text8 + listBox.Items[num2].ToString() + "\t";
												}
												else
												{
													text8 = text8 + listBox.Items[num2].ToString() + text7;
												}
											}
											else if (text7 == "1")
											{
												text8 += "\t";
											}
											else
											{
												text8 += text7;
											}
										}
										if (!string.IsNullOrEmpty(text8))
										{
											if (text7 == "1")
											{
												text8 = text8.TrimEnd(new char[]
												{
											'\t'
												});
											}
											else
											{
												text8 = text8.Substring(0, text8.Length - 1);
											}
										}
										this.WriteToTxt(text3, text8);
									}
								}
								string text9 = this.LoadBatFile(sFile, ref sMessage);
								if (!string.IsNullOrEmpty(sMessage))
								{
									result = false;
								}
								else
								{
									StringBuilder stringBuilder = new StringBuilder();
									foreach (string str in this.getGroupSampleFile(text2))
									{
										stringBuilder.AppendLine(text9.Replace("@PATH1", "\"" + str + "\"").Replace("@PATH2", "\"" + text3 + "\"").Replace("@QTY", iPrintQty.ToString()));
									}
									bool flag = true;
									bool flag2 = false;
									DateTime now = DateTime.Now;
									while (flag && !flag2)
									{
										try
										{
											flag = false;
											Process[] processes = Process.GetProcesses();
											for (int m = 0; m <= processes.Length - 1; m++)
											{
												if (processes[m].ProcessName.ToUpper() == "BARTEND")
												{
													flag = true;
												}
											}
											if (flag && (DateTime.Now - now).TotalSeconds > 60.0)
											{
												flag2 = true;
											}
										}
										catch (Exception ex)
										{
											sMessage = ex.Message;
											return false;
										}
									}
									if (flag2)
									{
										sMessage = "Print " + sType + "  Label TimeOut (60 Seconds)";
										result = false;
									}
									else
									{
										this.WriteToPrintGo(text6, stringBuilder.ToString());
										Setup.WinExec(text6, 0);
										sMessage = "OK";
										result = true;
										Thread.Sleep(2000);
									}
								}
							}
						}
					}
				}
				if (P == 0)
				{
					MessageBox.Show("没有模板哦☹☹", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			return result;
		}

        // Token: 0x06000006 RID: 6 RVA: 0x00003270 File Offset: 0x00001470
        public bool Print_Bartender_DataSource_Single(string sExeName, string sType, string sFileTitle, string sFileName, int iPrintQty, string sPrintMethod, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			sMessage = "OK";
			string text = "";
			bool result;
			if (!this.Get_Sample_File(sExeName, sFileTitle, sFileName, sPrintMethod, sPrintPort, ListParam, ListData, out text))
			{
				sMessage = string.Concat(new string[]
				{
					"Sample File not Exist",
					Environment.NewLine,
					"(",
					text,
					")",
					Environment.NewLine,
					"Print Method ",
					sPrintMethod,
					Environment.NewLine,
					"Print Port ",
					sPrintPort
				});
				result = false;
			}
			else
			{
				string startupPath = System.Windows.Forms.Application.StartupPath;
				FileInfo fileInfo = new FileInfo(text);
				fileInfo.Extension.ToString().ToUpper();
				string directoryName = fileInfo.DirectoryName;
				sFileName = Path.GetFileNameWithoutExtension(text);
				string text2 = text;
				string text3 = string.Concat(new string[]
				{
					startupPath,
					"\\",
					sExeName,
					"\\",
					sFileName,
					".lst"
				});
				string text4 = string.Concat(new string[]
				{
					startupPath,
					"\\",
					sExeName,
					"\\",
					sFileName,
					".dat"
				});
				string text5 = string.Concat(new string[]
				{
					startupPath,
					"\\",
					sExeName,
					"\\",
					sFileTitle,
					"DEFAULT.dat"
				});
				string text6 = startupPath + "\\" + sExeName + "\\PrintGo.bat";
				string sFile = startupPath + "\\" + sExeName + "\\PrintLabel.bat";
				if (!File.Exists(text2))
				{
					sMessage = "Label File Not Found (btw)" + Environment.NewLine + Environment.NewLine + text2;
					result = false;
				}
				else
				{
					if (!File.Exists(text4))
					{
						if (!File.Exists(text5))
						{
							sMessage = string.Concat(new string[]
							{
								"Label File Not Found (.dat)",
								Environment.NewLine,
								Environment.NewLine,
								text4,
								Environment.NewLine,
								Environment.NewLine,
								"OR ",
								Environment.NewLine,
								Environment.NewLine,
								text5
							});
							return false;
						}
						text4 = text5;
					}
					if (File.Exists(text3))
					{
						File.Delete(text3);
					}
					string text7 = "N/A";
					if (sFileTitle == "PKIN_")
					{
						text7 = "1";
					}
					System.Windows.Forms.ListBox listBox = this.LoadFileHeader(text4, ref sMessage, text7);
					if (!string.IsNullOrEmpty(sMessage))
					{
						result = false;
					}
					else
					{
						string text8 = string.Empty;
						for (int i = 0; i <= listBox.Items.Count - 1; i++)
						{
							if (text7 == "1")
							{
								text8 = text8 + listBox.Items[i].ToString() + "\t";
							}
							else
							{
								text8 = text8 + listBox.Items[i].ToString() + ",";
							}
						}
						if (!string.IsNullOrEmpty(text8))
						{
							if (text7 == "1")
							{
								text8 = text8.Trim(new char[]
								{
									'\t'
								});
							}
							else
							{
								text8 = text8.Substring(0, text8.Length - 1);
							}
						}
						this.WriteToTxt(text3, text8);
						new System.Windows.Forms.ListBox();
						if (sType == "PAGE")
						{
							text8 = string.Empty;
							for (int j = 0; j <= listBox.Items.Count - 1; j++)
							{
								string value = listBox.Items[j].ToString();
								int num = ListParam.Items.IndexOf(value);
								if (num >= 0)
								{
									if (text7 == "1")
									{
										text8 = text8 + ListData.Items[num].ToString() + "\t";
									}
									else
									{
										text8 = text8 + ListData.Items[num].ToString() + ",";
									}
								}
								else if (text7 == "1")
								{
									text8 += "\t";
								}
								else
								{
									text8 += ",";
								}
							}
							if (!string.IsNullOrEmpty(text8))
							{
								if (text7 == "1")
								{
									text8 = text8.Trim(new char[]
									{
										'\t'
									});
								}
								else
								{
									text8 = text8.Substring(0, text8.Length - 1);
								}
							}
							this.WriteToTxt(text3, text8);
						}
						else
						{
							text8 = string.Empty;
							for (int k = 0; k <= listBox.Items.Count - 1; k++)
							{
								string value2 = listBox.Items[k].ToString();
								int num2 = ListParam.Items.IndexOf(value2);
								if (num2 >= 0)
								{
									if (text7 == "1")
									{
										text8 = text8 + ListData.Items[num2].ToString() + "\t";
									}
									else
									{
										text8 = text8 + ListData.Items[num2].ToString() + ",";
									}
								}
								else if (text7 == "1")
								{
									text8 += "\t";
								}
								else
								{
									text8 += ",";
								}
							}
							if (!string.IsNullOrEmpty(text8))
							{
								if (text7 == "1")
								{
									text8 = text8.Trim(new char[]
									{
										'\t'
									});
								}
								else
								{
									text8 = text8.Substring(0, text8.Length - 1);
								}
							}
							this.WriteToTxt(text3, text8);
						}
						string text9 = this.LoadBatFile(sFile, ref sMessage);
						if (!string.IsNullOrEmpty(sMessage))
						{
							result = false;
						}
						else
						{
							text9 = text9.Replace("@PATH1", "\"" + text2 + "\"");
							text9 = text9.Replace("@PATH2", "\"" + text3 + "\"");
							text9 = text9.Replace("@QTY", iPrintQty.ToString());
							bool flag = true;
							bool flag2 = false;
							DateTime now = DateTime.Now;
							while (flag && !flag2)
							{
								try
								{
									flag = false;
									Process[] processes = Process.GetProcesses();
									for (int l = 0; l <= processes.Length - 1; l++)
									{
										if (processes[l].ProcessName.ToUpper() == "BARTEND")
										{
											flag = true;
										}
									}
									if (flag && (DateTime.Now - now).TotalSeconds > 60.0)
									{
										flag2 = true;
									}
								}
								catch (Exception ex)
								{
									sMessage = ex.Message;
									return false;
								}
							}
							if (flag2)
							{
								sMessage = "Print " + sType + "  Label TimeOut (60 Seconds)";
								result = false;
							}
							else
							{
								this.WriteToPrintGo(text6, text9);
								Setup.WinExec(text6, 0);
								sMessage = "OK";
								result = true;
							}
						}
					}
				}
			}
			return result;
		}

		// Token: 0x06000007 RID: 7 RVA: 0x00003934 File Offset: 0x00001B34
		public bool Print(string sExeName, string sType, string sFileTitle, string sFileName, int iPrintQty, string sPrintMethod, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			bool flag = true;
			sMessage = "OK";
			string text = "";
			bool result;
			if (!this.Get_Sample_File(sExeName, sFileTitle, sFileName, sPrintMethod, sPrintPort, ListParam, ListData, out text))
			{
				sMessage = string.Concat(new string[]
				{
					"Sample File not Exist",
					Environment.NewLine,
					"(",
					text,
					")",
					Environment.NewLine,
					"Print Method ",
					sPrintMethod,
					Environment.NewLine,
					"Print Port ",
					sPrintPort
				});
				result = false;
			}
			else
			{
				if (sPrintMethod.ToUpper() == "CODESOFT")
				{
					flag = this.Print_CodeSoft(text, iPrintQty, sPrintPort, ListParam, ListData, ref sMessage);
				}
				else if (sPrintMethod.ToUpper() == "BARTENDER")
				{
					if (sPrintPort.ToUpper() == "STANDARD")
					{
						flag = this.Print_BarTender_Standard(text, iPrintQty, sPrintPort, ListParam, ListData, ref sMessage);
					}
				}
				else if (sPrintMethod.ToUpper() == "DLL")
				{
					this.ListMultiParam.Items.Clear();
					DataSet dataSet = ClientUtils.ExecuteSQL("select * from sajet.s_sys_print_data where data_type = '" + sType + "' and (OUTPUT_PARAM is not null or OUTPUT_PARAM2 is not null) order by DATA_ORDER");
					for (int i = 0; i <= dataSet.Tables[0].Rows.Count - 1; i++)
					{
						string text2 = dataSet.Tables[0].Rows[i]["OUTPUT_PARAM"].ToString().ToUpper().TrimEnd(new char[]
						{
							';'
						});
						string text3 = dataSet.Tables[0].Rows[i]["OUTPUT_PARAM2"].ToString().ToUpper().TrimEnd(new char[]
						{
							';'
						});
						string[] array = text2.Split(new char[]
						{
							';'
						});
						string[] array2 = text3.Split(new char[]
						{
							';'
						});
						if (text2 != "")
						{
							for (int j = 0; j <= array.Length - 1; j++)
							{
								this.ListMultiParam.Items.Add(array[j].ToString());
							}
						}
						if (text3 != "")
						{
							for (int k = 0; k <= array2.Length - 1; k++)
							{
								this.ListMultiParam.Items.Add(array2[k].ToString());
							}
						}
					}
					if (sPrintPort.ToUpper() == "EXCEL")
					{
						flag = this.Print_Excel(text, iPrintQty, sPrintPort, ListParam, ListData, out sMessage);
					}
					else
					{
						flag = this.Print_DLL(text, iPrintQty, sPrintPort, ListParam, ListData, out sMessage);
					}
				}
				result = flag;
			}
			return result;
		}

		// Token: 0x06000008 RID: 8 RVA: 0x00003BFC File Offset: 0x00001DFC
		public bool Print_MultiLabel(string sExeName, string sType, string sFileTitle, string sFileName, int iPrintQty, string sPrintMethod, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			bool flag = true;
			sMessage = "OK";
			string text = "";
			bool result;
			if (!this.Get_Sample_File(sExeName, sFileTitle, sFileName, sPrintMethod, sPrintPort, ListParam, ListData, out text))
			{
				sMessage = string.Concat(new string[]
				{
					"Sample File not Exist",
					Environment.NewLine,
					"(",
					text,
					")",
					Environment.NewLine,
					"Print Method ",
					sPrintMethod,
					Environment.NewLine,
					"Print Port ",
					sPrintPort
				});
				result = false;
			}
			else
			{
				if (sPrintMethod.ToUpper() == "CODESOFT")
				{
					flag = this.Print_CodeSoft(text, iPrintQty, sPrintPort, ListParam, ListData, ref sMessage);
				}
				else if (sPrintMethod.ToUpper() == "BARTENDER")
				{
					if (sPrintPort.ToUpper() == "STANDARD")
					{
						flag = this.Print_BarTender_Standard(text, iPrintQty, sPrintPort, ListParam, ListData, ref sMessage);
					}
				}
				else if (sPrintMethod.ToUpper() == "DLL")
				{
					this.ListMultiParam.Items.Clear();
					this.ListMultiParam.Items.Add(sType.ToUpper() + "_");
					if (sPrintPort.ToUpper() == "EXCEL")
					{
						flag = this.Print_Excel(text, iPrintQty, sPrintPort, ListParam, ListData, out sMessage);
					}
					else
					{
						flag = this.Print_DLL(text, iPrintQty, sPrintPort, ListParam, ListData, out sMessage);
					}
				}
				result = flag;
			}
			return result;
		}

		// Token: 0x06000009 RID: 9 RVA: 0x00003D7C File Offset: 0x00001F7C
		public bool Print_TestPage(string sExeName, string sDefultText, string sFileTitle, string sFileName, int iPrintQty, string sPrintMethod, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			bool flag = true;
			sMessage = sDefultText;
			string text = "";
			bool result;
			if (!this.Get_Sample_File(sExeName, sFileTitle, sFileName, sPrintMethod, sPrintPort, ListParam, ListData, out text))
			{
				sMessage = string.Concat(new string[]
				{
					"Sample File not Exist",
					Environment.NewLine,
					"(",
					text,
					")",
					Environment.NewLine,
					"Print Method ",
					sPrintMethod,
					Environment.NewLine,
					"Print Port ",
					sPrintPort
				});
				result = false;
			}
			else
			{
				if (sPrintMethod.ToUpper() == "CODESOFT")
				{
					flag = this.Print_CodeSoft(text, iPrintQty, sPrintPort, ListParam, ListData, ref sMessage);
				}
				else if (sPrintMethod.ToUpper() == "BARTENDER")
				{
					if (sPrintPort.ToUpper() == "STANDARD")
					{
						flag = this.Print_BarTender_Standard(text, iPrintQty, sPrintPort, ListParam, ListData, ref sMessage);
					}
					else if (sPrintPort.ToUpper() == "DATASOURCE")
					{
						if (!string.IsNullOrEmpty(text))
						{
							text = Path.GetFileNameWithoutExtension(text);
						}
						flag = this.Print_Bartender_DataSource(sExeName, sDefultText, sFileTitle, text, iPrintQty, sPrintMethod, sPrintPort, ListParam, ListData, out sMessage);
					}
				}
				else if (sPrintMethod.ToUpper() == "DLL")
				{
					if (sPrintPort.ToUpper() == "EXCEL")
					{
						flag = this.Print_Excel(text, iPrintQty, sPrintPort, ListParam, ListData, out sMessage);
					}
					else
					{
						flag = this.Print_DLL(text, iPrintQty, sPrintPort, ListParam, ListData, out sMessage);
					}
				}
				result = flag;
			}
			return result;
		}

		// Token: 0x0600000A RID: 10 RVA: 0x00002050 File Offset: 0x00000250
		private void Open_CodeSoft()
		{
			if (this.lbl == null)
			{
				this.lbl = new LabelManager2.ApplicationClass();
			}
		}

		// Token: 0x0600000B RID: 11 RVA: 0x00002065 File Offset: 0x00000265
		private void Close_CodeSoft()
		{
			this.lbl.Quit();
		}

		// Token: 0x0600000C RID: 12 RVA: 0x00003F04 File Offset: 0x00002104
		private bool Print_CodeSoft(string sSampleFile, int iPrintQty, string sCodeSoftVer, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, ref string sMessage)
		{
			foreach (string text in this.getGroupSampleFile(sSampleFile))
			{
				try
				{
					this.lbl.Documents.Open(text, false);
					Document activeDocument = this.lbl.ActiveDocument;
					string[] array = new string[(int)activeDocument.Variables.FormVariables.Count];
					for (int i = 1; i <= (int)activeDocument.Variables.FormVariables.Count; i++)
					{
						array[i - 1] = activeDocument.Variables.FormVariables.Item(i).Name;
					}
					for (int j = 0; j <= array.Length - 1; j++)
					{
						string text2 = array[j].ToString();
						string text3 = this.Get_ParamData(text2, ListParam, ListData);
						if (sMessage != "OK" && text3 == "")
						{
							text3 = sMessage;
						}
						activeDocument.Variables.FormVariables.Item(text2).Value = text3;
					}
					activeDocument.PrintDocument(iPrintQty);
					sMessage = "OK";
				}
				catch (Exception ex)
				{
					sMessage = ex.Message;
					return false;
				}
			}
			return true;
		}

		// Token: 0x0600000D RID: 13
		[DllImport("kernel32.dll")]
		public static extern int WinExec(string exeName, int operType);

		// Token: 0x0600000E RID: 14 RVA: 0x00004080 File Offset: 0x00002280
		private bool Print_DLL(string sSampleFile, int iPrintQty, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			sMessage = "OK";
			foreach (string path in this.getGroupSampleFile(sSampleFile))
			{
				string text = File.ReadAllText(path);
				for (int i = 0; i <= ListParam.Items.Count - 1; i++)
				{
					string text2 = "%" + ListParam.Items[i].ToString() + "%";
					if (text.IndexOf(text2) > -1)
					{
						text = text.Replace(text2, ListData.Items[i].ToString());
					}
				}
				for (int j = 0; j <= this.ListMultiParam.Items.Count - 1; j++)
				{
					text = this.ReplaceNull(text, "%" + this.ListMultiParam.Items[j].ToString());
				}
				string text3 = System.Windows.Forms.Application.StartupPath + "\\PrintTemp";
				text3 = text3 + Convert.ToString(ListParam.Items.Count % 20) + ".txt";
				File.WriteAllText(text3, text);
				for (int k = 1; k <= iPrintQty; k++)
				{
					this.SendCommand(sPrintPort, text3);
				}
				File.Delete(text3);
			}
			return true;
		}

		// Token: 0x0600000F RID: 15 RVA: 0x000041F0 File Offset: 0x000023F0
		private void SendCommand(string sPort, string sFileName)
		{
			Process process = Process.Start(new ProcessStartInfo
			{
				FileName = "print",
				Arguments = string.Concat(new string[]
				{
					"/d:",
					sPort,
					" \"",
					sFileName,
					"\""
				}),
				UseShellExecute = false,
				CreateNoWindow = true
			});
			process.WaitForExit();
			process.Close();
		}

		// Token: 0x06000010 RID: 16 RVA: 0x00004260 File Offset: 0x00002460
		private string ReplaceNull(string sFileText, string sParam)
		{
			string text = sFileText;
			for (int i = text.IndexOf(sParam); i > -1; i = text.IndexOf(sParam))
			{
				int num = text.IndexOf("%", i + 1, text.Length - i - 1);
				string oldValue = text.Substring(i, num - i + 1);
				text = text.Replace(oldValue, "");
			}
			return text;
		}

		// Token: 0x06000011 RID: 17 RVA: 0x00002072 File Offset: 0x00000272
		private void Open_Excel()
		{
			this.ExcelApp = new Excel.ApplicationClass();
		}

		// Token: 0x06000012 RID: 18 RVA: 0x000042BC File Offset: 0x000024BC
		private void Close_Excel()
		{
			this.ExcelApp.Quit();
			this.NAR(this.ExcelWorksheet);
			this.NAR(this.ExcelWorkbook);
			this.NAR(this.ExcelWorkbooks);
			this.NAR(this.ExcelApp);
			GC.Collect();
		}

		// Token: 0x06000013 RID: 19 RVA: 0x0000430C File Offset: 0x0000250C
		private void NAR(object o)
		{
			try
			{
				if (o != null)
				{
					Marshal.ReleaseComObject(o);
				}
			}
			finally
			{
				o = null;
			}
		}

		// Token: 0x06000014 RID: 20 RVA: 0x0000433C File Offset: 0x0000253C
		private bool Print_Excel(string sSampleFile, int iPrintQty, string sPrintPort, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, out string sMessage)
		{
			sMessage = "OK";
			foreach (string template in this.getGroupSampleFile(sSampleFile))
			{
				try
				{
					this.ExcelWorkbooks = this.ExcelApp.Workbooks;
					this.ExcelWorkbook = this.ExcelWorkbooks.Add(template);
					this.ExcelWorksheet = (Worksheet)this.ExcelWorkbook.Worksheets[1];
					int num = this.ExcelWorksheet.UsedRange.Rows.Count + 1;
					int num2 = this.ExcelWorksheet.UsedRange.Columns.Count + 1;
					for (int i = 1; i <= num2; i++)
					{
						for (int j = 1; j <= num; j++)
						{
							string text = this.ExcelWorksheet.get_Range(this.ExcelWorksheet.Cells[j, i], this.ExcelWorksheet.Cells[j, i]).Text.ToString();
							if (text.IndexOf("%") > -1)
							{
								for (int k = 0; k <= ListParam.Items.Count - 1; k++)
								{
									string text2 = "%" + ListParam.Items[k].ToString() + "%";
									if (text.IndexOf(text2) > -1)
									{
										text = text.Replace(text2, ListData.Items[k].ToString());
										if (text.IndexOf("%") == -1)
										{
											break;
										}
									}
								}
								if (text.IndexOf("%") > -1)
								{
									for (int l = 0; l <= this.ListMultiParam.Items.Count - 1; l++)
									{
										text = this.ReplaceNull(text, "%" + this.ListMultiParam.Items[l].ToString());
									}
								}
								this.ExcelWorksheet.Cells[j, i] = text;
							}
						}
					}
					this.ExcelWorkbook.PrintOut(1, Type.Missing, iPrintQty, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				}
				catch (Exception ex)
				{
					sMessage = ex.Message;
					return false;
				}
				finally
				{
					this.ExcelWorkbook.Close(false, false, false);
					this.ExcelWorkbooks.Close();
				}
			}
			return true;
		}

		// Token: 0x06000015 RID: 21 RVA: 0x0000207F File Offset: 0x0000027F
		private void Open_BarTender()
		{
			if (this.bt == null)
			{
				this.bt = new BarTender.ApplicationClass();
			}
		}

		// Token: 0x06000016 RID: 22 RVA: 0x00002094 File Offset: 0x00000294
		private void Close_BarTender()
		{
			this.bt.Quit(BtSaveOptions.btDoNotSaveChanges);
		}

		// Token: 0x06000017 RID: 23 RVA: 0x00004644 File Offset: 0x00002844
		private bool Print_BarTender_Standard(string sSampleFile, int iPrintQty, string sCodeSoftVer, System.Windows.Forms.ListBox ListParam, System.Windows.Forms.ListBox ListData, ref string sMessage)
		{
			foreach (string text in this.getGroupSampleFile(sSampleFile))
			{
				try
				{
                    BarTender.Format format = this.bt.Formats.Open(text, false, "");
					foreach (object obj in format.NamedSubStrings)
					{
						string name = ((SubString)obj).Name;
						string text2 = this.Get_ParamData(name, ListParam, ListData);
						if (sMessage != "OK" && text2 == "")
						{
							text2 = sMessage;
						}
						format.SetNamedSubStringValue(name, text2);
					}
					format.PrintSetup.IdenticalCopiesOfLabel = iPrintQty;
					format.PrintOut(false, false);
					format.Close(BtSaveOptions.btDoNotSaveChanges);
					sMessage = "OK";
				}
				catch (Exception ex)
				{
					sMessage = ex.Message;
					return false;
				}
			}
			return true;
		}

		// Token: 0x06000018 RID: 24 RVA: 0x0000477C File Offset: 0x0000297C
		private System.Windows.Forms.ListBox LoadFileHeader(string sFile, ref string sMessage, string sSplitType)
		{
			sMessage = string.Empty;
			System.Windows.Forms.ListBox listBox = new System.Windows.Forms.ListBox();
			if (!File.Exists(sFile))
			{
				sMessage = "File not exist - " + sFile;
			}
			else
			{
				StreamReader streamReader = new StreamReader(sFile);
				try
				{
					char c = '\t';
					char[] array = sSplitType.ToCharArray();
					string[] array2;
					if (sSplitType == "1")
					{
						array2 = streamReader.ReadLine().Trim().Split(new char[]
						{
							c
						});
					}
					else
					{
						array2 = streamReader.ReadLine().Trim().Split(new char[]
						{
							array[0]
						});
					}
					for (int i = 0; i <= array2.Length - 1; i++)
					{
						listBox.Items.Add(array2[i].ToString());
					}
				}
				finally
				{
					streamReader.Close();
				}
			}
			return listBox;
		}

		// Token: 0x06000019 RID: 25 RVA: 0x00004850 File Offset: 0x00002A50
		private void WriteToTxt(string sFile, string sData)
		{
			StreamWriter streamWriter = null;
			try
			{
				streamWriter = new StreamWriter(sFile, true, Encoding.Default);
				streamWriter.WriteLine(sData);
			}
			finally
			{
				if (streamWriter != null)
				{
					streamWriter.Close();
				}
			}
		}

		// Token: 0x0600001A RID: 26 RVA: 0x000020A2 File Offset: 0x000002A2
		private void WriteToPrintGo(string sFile, string sData)
		{
			if (File.Exists(sFile))
			{
				File.Delete(sFile);
			}
			File.AppendAllText(sFile, sData, Encoding.Default);
		}

		// Token: 0x0600001B RID: 27 RVA: 0x00004890 File Offset: 0x00002A90
		private string LoadBatFile(string sFile, ref string sMessage)
		{
			sMessage = string.Empty;
			string result = string.Empty;
			if (!File.Exists(sFile))
			{
				DataSet dataSet = ClientUtils.ExecuteSQL("SELECT *    FROM SAJET.SYS_BASE  WHERE PROGRAM='Barcode Center'    AND PARAM_NAME = 'Bartender Print Command'    AND ROWNUM = 1 ");
				if (dataSet.Tables[0].Rows.Count > 0)
				{
					result = dataSet.Tables[0].Rows[0]["PARAM_VALUE"].ToString();
				}
				else
				{
					sMessage = "File not exist - " + sFile;
				}
			}
			else
			{
				StreamReader streamReader = new StreamReader(sFile);
				try
				{
					result = streamReader.ReadLine().Trim();
				}
				finally
				{
					streamReader.Close();
				}
			}
			return result;
		}

		// Token: 0x0600001C RID: 28 RVA: 0x000020BE File Offset: 0x000002BE
		public void Open(string sPrintMethod)
		{
			if (sPrintMethod == "CODESOFT")
			{
				this.Open_CodeSoft();
				return;
			}
			if (sPrintMethod == "EXCEL")
			{
				this.Open_Excel();
				return;
			}
			if (sPrintMethod == "BARTENDER")
			{
				this.Open_BarTender();
			}
		}

		// Token: 0x0600001D RID: 29 RVA: 0x000020FB File Offset: 0x000002FB
		public void Close(string sPrintMethod)
		{
			if (sPrintMethod == "CODESOFT")
			{
				this.Close_CodeSoft();
				return;
			}
			if (sPrintMethod == "EXCEL")
			{
				this.Close_Excel();
				return;
			}
			if (sPrintMethod == "BARTENDER")
			{
				this.Close_BarTender();
			}
		}

		// Token: 0x0600001E RID: 30 RVA: 0x0000493C File Offset: 0x00002B3C
		private List<string> getGroupSampleFile(string sSampleFile)
		{
			List<string> list = new List<string>();
			list.Add(sSampleFile);
			new FileInfo(sSampleFile);
			string extension = Path.GetExtension(sSampleFile);
			string directoryName = Path.GetDirectoryName(sSampleFile);
			string searchPattern = Path.GetFileNameWithoutExtension(sSampleFile) + "_*" + extension;
			foreach (string item in Directory.GetFiles(directoryName, searchPattern))
			{
				list.Add(item);
			}
			return list;
		}

		// Token: 0x04000001 RID: 1
		private System.Windows.Forms.ListBox ListMultiParam = new System.Windows.Forms.ListBox();

		// Token: 0x04000002 RID: 2
		private LabelManager2.ApplicationClass lbl;

		// Token: 0x04000003 RID: 3
		private Excel.Application ExcelApp;

		// Token: 0x04000004 RID: 4
		private Workbooks ExcelWorkbooks;

		// Token: 0x04000005 RID: 5
		private Workbook ExcelWorkbook;

		// Token: 0x04000006 RID: 6
		private Worksheet ExcelWorksheet;

		// Token: 0x04000007 RID: 7
		private BarTender.ApplicationClass bt;

		private List<string> ts ;
	}
}
