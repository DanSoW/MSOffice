using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace MSOffice
{
	public partial class Window : Form
	{
		private string _lastOpenExcelFile = "";

		public Window()
		{
			InitializeComponent();
		}

		private void button3_Click(object sender, EventArgs e)
		{
			openFileDialog1.Filter = "Файлы изображений|*.bmp;*.png;*.jpg";
			if (openFileDialog1.ShowDialog() != DialogResult.OK)
				return;

			try
			{
				pictureBox1.Image = System.Drawing.Image.FromFile(openFileDialog1.FileName);
				pictureBox1.ImageLocation = openFileDialog1.FileName;
			}
			catch (OutOfMemoryException)
			{
				MessageBox.Show("Ошибка чтения картинки");
				return;
			}

			pictureBox1.Invalidate();

			MessageBox.Show("Изображение загружено!");
		}

		private void button4_Click(object sender, EventArgs e)
		{
			openFileDialog1.Filter = "Файлы Word|*.docx";
			if (openFileDialog1.ShowDialog() != DialogResult.OK)
				return;

			string pathDocument = openFileDialog1.FileName;
			DocX document = DocX.Load(pathDocument);

			Xceed.Document.NET.Image image = document.AddImage(pictureBox1.ImageLocation);

			Paragraph paragraph = document.InsertParagraph();

			paragraph.AppendPicture(image.CreatePicture());

			paragraph.Alignment = Alignment.center;

			document.Save();

			MessageBox.Show("Картинка добавлена в документ!");
		}

		private void button6_Click(object sender, EventArgs e)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataRow row;

			openFileDialog1.Filter = "Файлы Excel|*.xlsx";
			if (openFileDialog1.ShowDialog() != DialogResult.OK)
			{
				return;
			}

			_lastOpenExcelFile = openFileDialog1.FileName;
			dataGridView1.Rows.Clear();
			dataGridView1.Columns.Clear();

			try
			{
				Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
				Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(_lastOpenExcelFile);
				Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
				Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

				int rowCount = excelRange.Rows.Count;
				int colCount = excelRange.Columns.Count;

				for (int i = 1; i <= rowCount; i++)
				{
					for (int j = 1; j <= colCount; j++)
					{
						dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
					}
					break;
				}

				int rowCounter = 0;
				for (int i = 2; i <= rowCount; i++)
				{
					row = dt.NewRow();
					rowCounter = 0;
					for (int j = 1; j <= colCount; j++)
					{
						if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
						{
							row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
						}
						else
						{
							row[rowCounter] = "";
						}
						rowCounter++;
					}
					dt.Rows.Add(row);
				}

				dataGridView1.DataSource = dt;

				// Закрытие и очистка Excel процесса
				GC.Collect();
				GC.WaitForPendingFinalizers();
				Marshal.ReleaseComObject(excelRange);
				Marshal.ReleaseComObject(excelWorksheet);
				// Выход из Excel
				excelWorkbook.Close();
				Marshal.ReleaseComObject(excelWorkbook);
				excelApp.Quit();
				Marshal.ReleaseComObject(excelApp);

				MessageBox.Show("Данные из документа Excel загружены!");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				_lastOpenExcelFile = "";
			}
		}

		private void button5_Click(object sender, EventArgs e)
		{
			if(_lastOpenExcelFile.Length <= 0)
			{
				MessageBox.Show("Необходимо открытие документа Excel");
				return;
			}

			try
			{
				Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
				Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(_lastOpenExcelFile);
				Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
				Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

				for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
				{
					excelWorksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
				}

				for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
				{
					for (int j = 0; j < dataGridView1.Columns.Count; j++)
					{
						excelWorksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
					}
				}


				excelWorkbook.SaveAs(_lastOpenExcelFile,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
					Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, 
					Type.Missing, Type.Missing, Type.Missing, Type.Missing);

				// Закрытие и очистка Excel процесса
				GC.Collect();
				GC.WaitForPendingFinalizers();
				Marshal.ReleaseComObject(excelRange);
				Marshal.ReleaseComObject(excelWorksheet);
				// Выход из Excel
				excelWorkbook.Close();
				Marshal.ReleaseComObject(excelWorkbook);
				excelApp.Quit();
				Marshal.ReleaseComObject(excelApp);

				MessageBox.Show("Данные были сохранены!");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				_lastOpenExcelFile = "";
			}
		}

		private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
		{
			System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)e.Control;
			tb.KeyPress += new KeyPressEventHandler(tb_KeyPress);
		}

		void tb_KeyPress(object sender, KeyPressEventArgs e)
		{

			if (!(Char.IsDigit(e.KeyChar)))
			{
				if (e.KeyChar != (char)Keys.Back)
				{
					e.Handled = true;
				}
			}
		}

		private void button7_Click(object sender, EventArgs e)
		{
			System.Data.DataTable dt = new System.Data.DataTable();

			saveFileDialog1.Filter = "Файлы Excel|*.xlsx";
			if (saveFileDialog1.ShowDialog() != DialogResult.OK)
			{
				return;
			}

			_lastOpenExcelFile = saveFileDialog1.FileName;

			try
			{
				Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
				Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Add(Type.Missing);
				Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

				for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
				{
					for (int j = 0; j < dataGridView1.Columns.Count; j++)
					{
						var value = dataGridView1.Rows[i].Cells[j].Value;
						excelWorksheet.Cells[i + 1, j + 1]
							= (value == null) ? "" : value.ToString();
					}
				}

				excelWorkbook.SaveAs(_lastOpenExcelFile,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing);

				// Закрытие и очистка Excel процесса
				GC.Collect();
				GC.WaitForPendingFinalizers();
				Marshal.ReleaseComObject(excelWorksheet);
				// Выход из Excel
				excelWorkbook.Close();
				Marshal.ReleaseComObject(excelWorkbook);
				excelApp.Quit();
				Marshal.ReleaseComObject(excelApp);

				MessageBox.Show("Данные были загружены!");
				_lastOpenExcelFile = "";
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				_lastOpenExcelFile = "";
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			saveFileDialog1.Filter = "Файлы Word|*.docx";
			if (saveFileDialog1.ShowDialog() != DialogResult.OK)
			{
				return;
			}

			string pathDocument = saveFileDialog1.FileName;
			DocX document = DocX.Create(pathDocument);

			Table table = document.AddTable(dataGridView2.Rows.Count, dataGridView2.Columns.Count);
			table.Alignment = Alignment.center;
			table.Design = TableDesign.TableGrid;

			for(int i = 0; i < dataGridView2.Rows.Count; i++)
			{
				for(int j = 0; j < dataGridView2.Rows[i].Cells.Count; j++)
				{
					var value = dataGridView2.Rows[i].Cells[j].Value;
					table.Rows[i].Cells[j].Paragraphs[0].Append(
						(value != null)? value.ToString() : ""
						);
				}
			}

			document.InsertParagraph().InsertTableAfterSelf(table);

			document.Save();

			MessageBox.Show("Таблица загружена в Word документ!");
		}

		private void button2_Click(object sender, EventArgs e)
		{
			openFileDialog1.Filter = "Файлы Word|*.docx";
			if (openFileDialog1.ShowDialog() != DialogResult.OK)
			{
				return;
			}

			string pathDocument = openFileDialog1.FileName;
			DocX document = DocX.Load(pathDocument);

			Table table = document.AddTable(dataGridView2.Rows.Count, dataGridView2.Columns.Count);
			table.Alignment = Alignment.center;
			table.Design = TableDesign.TableGrid;

			for (int i = 0; i < dataGridView2.Rows.Count; i++)
			{
				for (int j = 0; j < dataGridView2.Rows[i].Cells.Count; j++)
				{
					var value = dataGridView2.Rows[i].Cells[j].Value;
					table.Rows[i].Cells[j].Paragraphs[0].Append(
						(value != null) ? value.ToString() : ""
						);
				}
			}

			document.InsertParagraph().InsertTableAfterSelf(table);

			document.Save();

			MessageBox.Show("Таблица загружена в Word документ!");
		}
	}
}
