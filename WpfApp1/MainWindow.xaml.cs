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

/**

	@author Lavrentiy Leonov,Beskrovnaya Lina
	@version 1.0
	@date April-May 2018
*/
namespace WpfApp1
{
    /// <summary>
    /// Класс MainWindow
    /// </summary>

    /// <remarks>
    /// Создаем глобальные переменные 
    /// </remarks>
    /// \n path строковая переменная для пути к файлу с именем врачей
    /// \n data[] массив для данных, введенных пользователем
    /// <summary>
    /// <see cref="Data()">
    /// </see>
    /// Функция Data() для заполнения массива данныых data значениями из грида, котрые выбрал пользователь
    /// </summary>
    /// i-счетчик
    /// <exception cref="Exeption ex"> возникает, когда человек не ввел все данные в ячейки
    /// </exception>
    /// <summary>
    /// \n Стандартная функция Window_Loaded, вызывающаяся, когда загружается окно MainWindow
    /// \n Добавляются вручную элементы в Combobox, котрые будут доступны для выбора пользователю
    ///\n Данные берем из советов опытного врача, они достоверны.
    /// \n Для сохранения и дальнейшего изменения имени врача в документе будет использвоать текстовый документ, в котором будут хранится Ф.И.О. врачей
    /// </summary>
    /// <summary>
    /// Считывание с файла имя врача
    /// \n StreamReader sr = new StreamReader(fs);
    /// \n name1= sr.ReadLine();
    ///  \n if(String.IsNullOrEmpty(name1)||name1==" "!=true)
    ///  </summary>
    /// <summary>
    ///Считывание из файла имя врача
    ///
    ///   \n name.Text = name1;
    ///   \n name.IsEnabled = false;-При успешном считывании из файла делаем ячейку для ввода имени врача затемненной и без функционала
    ///  \n}
    /// \n else
    /// \n {
    /// \n     name.IsEnabled = true;-При неуспешном считывании из файла делаем ячейку активной для ввода-
    /// \n}
    ///  \nsr.Close(); 
    ///\n}
    /// </summary>
    ///  <summary> 
    /// \n Функция TextBox_TextChanged_1( при активации текстовог поля для ввода имени врача
    ///  </summary>
    ///  \nДанная функция сохраняет имя, введенное пользователем в ячейку в переменную name1 и вызывает дальше фнукцию для сохранения этого имени в файл.

    /// <summary>
    ///Код функции выглядит следующим образом:\n
    ///
    ///private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)\n
    ///  {\n
    ///    if (String.IsNullOrEmpty(name1) || name1 == " " == true)\n
    ///  {\n
    ///       name1 = name.Text;\n
    ///      SaveName();\n
    ///  }\n
    ///   }\n
    /// </summary> 
    /// <summary>
    ///Фнукция SaveName() сохранения мени врача в файл
    /// Данная функция реализует открытие файла, сохранение данных имени в файл и закрытие файла.\n
    ///Код функции выглядит следующим образом:\n
    ///  private void SaveName()\n
    ///   {\n
    ///   FileStream fr = new FileStream(path + "\\namedoc.txt", FileMode.OpenOrCreate, FileAccess.Write);\n
    ///   StreamWriter sw = new StreamWriter(fr);\n
    ///     sw.WriteLine(name1);\n
    ///
    ///        sw.Close();    \n
    ///   }\n
    /// </summary>
    ///  <summary>
    /// Событие нажатия на меню  MenuItem_Click\n
    /// </summary>
    /// <summary>
    /// Активизирует ячейку для ввода имени врача, позволяя вписать другое имя\n
    /// </summary>
    /// <summary>
    /// Код функции:
    /// </summary>
    /// <summary>
    /// private void Save_Click(object sender, RoutedEventArgs e)\n
    /// {\n
    ///  Data();\n
    /// Word.Document doc = null;\n
    ///  try\n
    /// {\n
    ///
    /// Word.Application app = new Word.Application();\n
    ///
    ///string source = @"D:\\Doctor.docx";\n
    /// doc = app.Documents.Open(source);\n
    ///doc.Activate();\n
    /// Word.Bookmarks wBookmarks = doc.Bookmarks;\n
    /// Word.Range wRange;\n
    /// int i = 0;\n
    ///    foreach (Word.Bookmark mark in wBookmarks)\n
    ///  {\n
    /// wRange = mark.Range;\n
    /// wRange.Text = data[i];\n
    /// i++;\n
    /// }\n
    ///  Object fileName = @"D:\\"+Patient.Text+".doc";\n
    ///  Object fileFormat = Word.WdSaveFormat.wdFormatDocument;\n
    ///Object lockComments = false;\n
    /// Object password = "";\n
    ///Object addToRecentFiles = false;\n
    /// Object writePassword = "";\n
    /// Object readOnlyRecommended = false;\n
    /// Object embedTrueTypeFonts = false;\n
    ///  Object saveNativePictureFormat = false;\n
    /// Object saveFormsData = false;\n
    /// Object saveAsAOCELetter = Type.Missing;\n
    /// Object encoding = Type.Missing;\n
    /// Object insertLineBreaks = Type.Missing;\n
    /// Object allowSubstitutions = Type.Missing;\n
    /// Object lineEnding = Type.Missing;\n
    /// Object addBiDiMarks = Type.Missing;\n
    ///doc.SaveAs(ref fileName,\n
    ///  ref fileFormat, ref lockComments,\n
    /// ref password, ref addToRecentFiles, ref writePassword,\n
    /// ref readOnlyRecommended, ref embedTrueTypeFonts,\n
    ///ref saveNativePictureFormat, ref saveFormsData,\n
    /// ref saveAsAOCELetter, ref encoding, ref insertLineBreaks,\n
    /// ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);\n
    /// doc.Close();\n
    ///  doc = null;\n
    /// }\n
    ///    catch (Exception ex)\n
    ///   {\n
    ///doc.Close();\n
    ///  doc = null;\n
    /// Console.WriteLine("Во время выполнения произошла ошибка!");\n
    ///  Console.ReadLine();\n
    /// }\n
    /// </summary>
    /// <summary> 
    ///Функция Save_Click вызывается по щелчку на кнопку Save в окнце документа\n
    ///Сораняет данные, выбранные пользователем в документ Microsoft Word\n
    ///1.Вызывает фнукцию Data() для сохранения всех введенных пользователем данных в массив data[].\n
    ///2.Создает экземпляр Word.Document doc\n
    ///3.Создает объект приложения Word.Application app = new Word.Application()\n
    ///4.Открывает документ из введенной директории\n
    ///5.Создает объект wBookmarks, который содержит все закладки\n
    ///6.В цикле присваивает каждой закладке текст из массива данных пользваотеля data[]\n
    ///7.Сохраняем в документ новый, именуем его Ф.И.О. пациента\n
    ///8.Закрываем документ\n
    /// </summary>
    public partial class MainWindow : Window
    {
        string name1 = "";
        string path = Directory.GetCurrentDirectory();
        string [] data=new string[35];
        string newname = "";
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
            combobox3.Items.Add("Добавить...");
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
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Data();
            Word.Document doc = null;
            try
            {
                Word.Application app = new Word.Application();
             
            //    string source = @"D:\\Doctor.docx";
                doc = app.Documents.Open(path+ "\\Doctor.docx");
                doc.Activate();
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;

                foreach (Word.Bookmark mark in wBookmarks)
                { 
                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }
                Object fileName = @"D:\\"+Patient.Text+".doc";
                Object fileFormat = Word.WdSaveFormat.wdFormatDocument;
                Object lockComments = false;
                Object password = "";
                Object addToRecentFiles = false;
                Object writePassword = "";
                Object readOnlyRecommended = false;
                Object embedTrueTypeFonts = false;
                Object saveNativePictureFormat = false;
                Object saveFormsData = false;
                Object saveAsAOCELetter = Type.Missing;
                Object encoding = Type.Missing;
                Object insertLineBreaks = Type.Missing;
                Object allowSubstitutions = Type.Missing;
                Object lineEnding = Type.Missing;
                Object addBiDiMarks = Type.Missing;
                doc.SaveAs(ref fileName,
 ref fileFormat, ref lockComments,
 ref password, ref addToRecentFiles, ref writePassword,
 ref readOnlyRecommended, ref embedTrueTypeFonts,
 ref saveNativePictureFormat, ref saveFormsData,
 ref saveAsAOCELetter, ref encoding, ref insertLineBreaks,
 ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);
                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                doc.Close();
                doc = null;
                Console.WriteLine("Во время выполнения произошла ошибка!");
                Console.ReadLine();
            }
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
    }

