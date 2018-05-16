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
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string name1 = "";
        string path = Directory.GetCurrentDirectory();
        int N = 35;
        string [] data=new string[35];
        public MainWindow()
        {
            InitializeComponent();
          
           
        }

            private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {

            }

            private void Grid_Loaded(object sender, RoutedEventArgs e)
            {

            }
        private void Data()
        {
            try
            {
                int i = 0;
                data[i] = name.Text;
                i++;
                data[i] = combobox1.SelectedItem.ToString();
                i++;
                data[i] = Patient.Text;
                i++;
                data[i] = Age.Text;
                i++;
                data[i] = Date.Text;
                i++;
                data[i] = combobox2.SelectedItem.ToString();
                i++;
                data[i] = combobox3.SelectedItem.ToString();
                i++;
                data[i] = combobox4.SelectedItem.ToString();
                i++;
                data[i] = combobox5.SelectedItem.ToString();
                i++;
                data[i] = combobox6.SelectedItem.ToString();
                i++;
                data[i] = combobox7.SelectedItem.ToString();
                i++;
                data[i] = combobox8.SelectedItem.ToString();
                i++;
                data[i] = combobox9.SelectedItem.ToString();
                i++;
                data[i] = Sm.Text;
                i++;
                data[i] = combobox10.SelectedItem.ToString();
                i++;
                data[i] = Comment1.Text;
                i++;
                data[i] = combobox11.SelectedItem.ToString();
                i++;
                data[i] = combobox12.SelectedItem.ToString();
                i++;
                data[i] = combobox13.SelectedItem.ToString();
                i++;
                data[i] = combobox14.SelectedItem.ToString();
                i++;
                data[i] = combobox15.SelectedItem.ToString();
                i++;
                data[i] = combobox16.SelectedItem.ToString();
                i++;
                data[i] = combobox17.SelectedItem.ToString();
                i++;
                data[i] = combobox18.SelectedItem.ToString();
                i++;
                data[i] = combobox19.SelectedItem.ToString();
                i++;
                data[i] = combobox20.SelectedItem.ToString();
                i++;
                data[i] = combobox21.SelectedItem.ToString();
                i++;
                data[i] = combobox22.SelectedItem.ToString();
                i++;
                data[i] = Comment2.Text;
                i++;
                data[i] = combobox23.SelectedItem.ToString();
                i++;
                data[i] = combobox24.SelectedItem.ToString();
                i++;
                data[i] = combobox25.SelectedItem.ToString();
                i++;
                data[i] = Comment3.Text;
                i++;
                data[i] = Conclusion.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Заполните все поля!", MessageBoxButton.OK.ToString());
            }
           
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        { 
          
            combobox1.Items.Add("Видеогастроскоп Olympus GIF H180J");
            combobox2.Items.Add("спрей 10% Sol Lidicaini");
            combobox3.Items.Add("свободно проходим");
            combobox3.Items.Add("сужен");
            combobox3.Items.Add("расширен");
            combobox4.Items.Add("эластичны");
            combobox4.Items.Add("регидны");
            combobox5.Items.Add("бледно-розовая");
            combobox5.Items.Add("гиперемирована");
            combobox6.Items.Add("гладкая");
            combobox6.Items.Add("зернистая");
            combobox7.Items.Add("выражена");
            combobox7.Items.Add("не выражена");
            combobox8.Items.Add("ровная");
            combobox8.Items.Add("неровная");
            combobox9.Items.Add("смыкается");
            combobox9.Items.Add("несмыкается");
            combobox9.Items.Add("смыкается неполностью");
            combobox10.Items.Add("плотно");
            combobox10.Items.Add("неплотно");
            combobox10.Items.Add("не");
            combobox11.Items.Add("обычных размеров");
            combobox11.Items.Add("расширен");
            combobox11.Items.Add("уменьшен в размерах");
            combobox11.Items.Add("деформирован");
            combobox12.Items.Add("ослизненная жидкость");
            combobox12.Items.Add("желчь");
            combobox12.Items.Add("кровь");
            combobox12.Items.Add("пищевые массы");
            combobox13.Items.Add("обычных размеров");
            combobox13.Items.Add("увеличены");
            combobox13.Items.Add("уплощены");
            combobox14.Items.Add("полностью");
            combobox14.Items.Add("неполностью");
            combobox14.Items.Add("не расправляются");
            combobox15.Items.Add("бледно-розовая");
            combobox15.Items.Add("гиперемирована");
            combobox16.Items.Add("не выражен");
            combobox16.Items.Add("усилен");
            combobox17.Items.Add("вялая");
            combobox17.Items.Add("активная");
            combobox18.Items.Add("округлой");
            combobox18.Items.Add("овальной");
            combobox18.Items.Add("неправильной");
            combobox19.Items.Add("проходим");
            combobox19.Items.Add("не проходим");
            combobox20.Items.Add("не деформирована");
            combobox20.Items.Add("деформирована");
            combobox21.Items.Add("бледно-розовая");
            combobox21.Items.Add("гиперемирована");
            combobox22.Items.Add("ворсинчатая");
            combobox22.Items.Add("гладкая");
            combobox22.Items.Add("зернистая");
            combobox23.Items.Add("бледно-розовая");
            combobox23.Items.Add("гиперемирована");
            combobox24.Items.Add("есть");
            combobox24.Items.Add("нет");
            combobox25.Items.Add("не изменена");
            combobox25.Items.Add("изменена");
            FileStream fs = new FileStream(path + "\\namedoc.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
         
            StreamReader sr = new StreamReader(fs);
            name1= sr.ReadLine();
            if(String.IsNullOrEmpty(name1)||name1==" "!=true)
            {
                name.Text = name1;
                name.IsEnabled = false;
            }
            else
            {
                name.IsEnabled = true;
            }
             sr.Close();
        
        }

      

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void combobox26_MouseLeave(object sender, MouseEventArgs e)
        {

            
        }

        private void combobox26_LostStylusCapture(object sender, StylusEventArgs e)
        {
        }


        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            if (String.IsNullOrEmpty(name1) || name1 == " " == true)
            {
                name1 = name.Text;
                SaveName();
            }
        }
        private void SaveName()
        {
            FileStream fr = new FileStream(path + "\\namedoc.txt", FileMode.OpenOrCreate, FileAccess.Write);
         
            StreamWriter sw = new StreamWriter(fr);
            sw.WriteLine(name1);
       
            sw.Close();
            
        }
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            name.IsEnabled = true;
            name1 = name.Text;
            SaveName();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Data();
            Word.Document doc = null;
            try
            {
                // Создаём объект приложения
                Word.Application app = new Word.Application();
                // Путь до шаблона документа
                string source = @"C:\\Users\\Lavrentiy\\Documents\\Doctor.docx";
                // Открываем
                doc = app.Documents.Open(source);
                doc.Activate();

                // Добавляем информацию
                // wBookmarks содержит все закладки
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;
               
                foreach (Word.Bookmark mark in wBookmarks)
                {

                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }

                // Закрываем документ
                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                // Если произошла ошибка, то
                // закрываем документ и выводим информацию
                doc.Close();
                doc = null;
                Console.WriteLine("Во время выполнения произошла ошибка!");
                Console.ReadLine();
            }
        }
    }
    }

