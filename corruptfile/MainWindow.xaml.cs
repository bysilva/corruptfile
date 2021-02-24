using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace corruptfile
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreateDocument(string titulo,string extension)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application
                {
                    ShowAnimation = false,

                    Visible = false
                };

                object missing = System.Reflection.Missing.Value;

                Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);


                document.Content.SetRange(0, 0);

                document.Content.Text = RandomText();

                object filename = $@"c:\{titulo}.{extension}";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                CorrupFile(filename.ToString());
                MessageBox.Show($"Archivo creado correctamente {filename.ToString()}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void MoverArchivo(string pathFile,string nameDoc)
        {
            string docPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDoc‌​uments), "CorruptFile");
            if (!Directory.Exists(docPath))
                Directory.CreateDirectory(docPath);
            string destinationFile = System.IO.Path.Combine(docPath,nameDoc);
            System.IO.File.Move(pathFile, destinationFile);
            System.IO.File.Delete(pathFile);
        }
        void CorrupFile(string filename)
        {
            string text = File.ReadAllText(filename);
            text = text.Replace("PK", "K");
            File.WriteAllText(filename, text);
        }
        string RandomText()
        {
            Random r = new Random();
            int length = r.Next(100000,200000);

            StringBuilder str_build = new StringBuilder();
            Random random = new Random();
            for (int i = 0; i < length; i++)
            {
                str_build.Append(Convert.ToChar(Convert.ToInt32(Math.Floor(25 * random.NextDouble())) + 65));
            }
            return str_build.ToString();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtTitulo.Text))
            {
                MessageBox.Show("Ingresa el titulo");
            }
            else
            {
                CreateDocument(txtTitulo.Text, ((Archivos)cmbDoc.SelectedItem).Extension);
            }
        }

        readonly GenerateArchivos Generator = new GenerateArchivos();
        void PopulateArchivos()
        {
            cmbDoc.ItemsSource = Generator.GetArchivos();
            cmbDoc.SelectedIndex = 0;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtTitulo.Focus();
            PopulateArchivos();
        }
    }
}
