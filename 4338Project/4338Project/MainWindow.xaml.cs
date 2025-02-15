using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace _4338Project
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void myButton_Click(object sender, RoutedEventArgs e)
        {
            Kuznetsova_4338 albina = new Kuznetsova_4338();
            albina.Show();
        }

        private void myButton_Click_Adel(object sender, RoutedEventArgs e)
        {
            Yusupov_4338 yusupov_4338 = new Yusupov_4338();
            yusupov_4338.Show();
            this.Hide();
        }
        
        private void AidarInfoButton_Click(object sender, RoutedEventArgs e)
        {
            Salikhzyanov_4338 aidar = new Salikhzyanov_4338(); 
            aidar.Show();
        }

        private void AmirInfoButton_Click(object sender, RoutedEventArgs e)
        {
            Nasyybullin_4338 nasyybullin_4338 = new Nasyybullin_4338();
            nasyybullin_4338.Show();
            this.Hide();
        }

        private void Button_Click_Ksen(object sender, RoutedEventArgs e)
        {
            Borisova_4338 ksen = new Borisova_4338();
            ksen.Show();
        }

		private void KsenButton_Click(object sender, RoutedEventArgs e)
		{

        }

		private void KsenButton_Click_1(object sender, RoutedEventArgs e)
		{

        }

		private void Button_Click_Di(object sender, RoutedEventArgs e)
		{
			BayazovaDi bayazovaDi = new BayazovaDi();
			bayazovaDi.Show();
			this.Hide();
		}
    }
}
