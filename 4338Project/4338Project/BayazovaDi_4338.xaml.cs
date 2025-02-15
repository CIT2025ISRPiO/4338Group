using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace _4338Project
{
	/// <summary>
	/// Логика взаимодействия для BayazovaDi.xaml
	/// </summary>
	public partial class BayazovaDi : Window
	{
		private Usl _currentMaterial = new Usl();
		public BayazovaDi()
		{
			InitializeComponent();
			DataContext = _currentMaterial;
			dataGridUsl.ItemsSource = isrpo_2_lbEntities.GetContext().Usl.OrderBy(x => x.IdServices).ToList();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			MessageBox.Show("Дианчик одуванчик", "Мур мур мяу мяу");
		}

		private void Button_Click_2(object sender, RoutedEventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog()
			{
				DefaultExt = "*.xlsx",
				Filter = "Excel Files (*.xlsx)|*.xlsx",
				Title = "Выберите файл для импорта"
			};

			if (!(ofd.ShowDialog() == true))
				return;

			try
			{
				using (var wb = new XLWorkbook(ofd.FileName))
				{
					var worksheet = wb.Worksheet(1);
					var range = worksheet.RangeUsed();
					var rows = range.RowsUsed().Skip(1);

					List<Usl> uslsList = new List<Usl>();

					foreach (var row in rows)
					{
						if (int.TryParse(row.Cell(1).GetString(), out int id) &&
							decimal.TryParse(row.Cell(5).GetString(), out decimal cost))
						{
							uslsList.Add(new Usl()
							{
								IdServices = id,
								NameServices = row.Cell(2).GetString(),
								TypeOfService = row.Cell(3).GetString(),
								CodeService = row.Cell(4).GetString(),
								Cost = cost
							});
						}
					}

					using (var context = new isrpo_2_lbEntities())
					{
						context.Usl.AddRange(uslsList);
						context.SaveChanges();
					}
				}

				MessageBox.Show("Данные успешно импортированы!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
				dataGridUsl.ItemsSource = isrpo_2_lbEntities.GetContext().Usl.OrderBy(x => x.IdServices).ToList();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Ошибка при импорте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void Button_Click_1(object sender, RoutedEventArgs e)
		{
			SaveFileDialog sfd = new SaveFileDialog()
			{
				DefaultExt = "*.xlsx",
				Filter = "Excel Files (*.xlsx)|*.xlsx",
				Title = "Выберите место для сохранения файла"
			};

			if (sfd.ShowDialog() != true)
				return;

			try
			{
				using (var context = new isrpo_2_lbEntities())
				{
					var services = context.Usl.OrderBy(s => s.Cost).ToList();
					var groupedServices = services.GroupBy(s => s.TypeOfService);

					using (var workbook = new XLWorkbook())
					{
						foreach (var group in groupedServices)
						{
							CreateSheet(workbook, group.Key, group.ToList());
						}
						workbook.SaveAs(sfd.FileName);
					}
				}

				MessageBox.Show("Экспорт выполнен успешно!", "Успешно", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void CreateSheet(XLWorkbook workbook, string sheetName, List<Usl> data)
		{
			var worksheet = workbook.Worksheets.Add(sheetName);

			worksheet.Cell(1, 1).Value = "ID";
			worksheet.Cell(1, 2).Value = "Название услуги";
			worksheet.Cell(1, 3).Value = "Вид услуги";
			worksheet.Cell(1, 4).Value = "Код услуги";
			worksheet.Cell(1, 5).Value = "Стоимость";

			int row = 2;
			foreach (var service in data)
			{
				worksheet.Cell(row, 1).Value = service.IdServices;
				worksheet.Cell(row, 2).Value = service.NameServices;
				worksheet.Cell(row, 3).Value = service.TypeOfService;
				worksheet.Cell(row, 4).Value = service.CodeService;
				worksheet.Cell(row, 5).Value = service.Cost;
				row++;
			}

			worksheet.Columns().AdjustToContents();
		}
	}
}
