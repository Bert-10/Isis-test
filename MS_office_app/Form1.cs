using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace MS_office_app
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {

            
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Image Files(*.PNG;*.JPG;*.BMP;*.GIF)|*.PNG;*.JPG;.BMP;*.GIF";
            openFileDialog.Title = "Открыть изображение";
            //    /*
            object oMissing = System.Reflection.Missing.Value;
          object oEndOfDoc = "\\endofdoc"; 

          //Start Word and create a new document.
          Word._Application oWord;
          Word._Document oDoc;
          oWord = new Word.Application();
          oWord.Visible = true;
          oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
          ref oMissing, ref oMissing);

          if (openFileDialog.ShowDialog() == DialogResult.OK)
          {
              oDoc.InlineShapes.AddPicture(openFileDialog.FileName);
          }

          oDoc.Save();
//          oDoc.Close();
          //this.Close();

          //  */

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //  /*
            string find = textBox1.Text;
            string change = textBox2.Text;
            if ((find == "") | (change == ""))
            {
                MessageBox.Show("Невозможно выполнить замену текста, поля для ввода не должны быть пустыми", "Некорректный ввод");
            }
            else {
                FileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files(*.xlsx;*.xls)|*.xlsx;*.xls";
                openFileDialog.Title = "Открыть файл Excel";
                //   Excel._Application oExcel=new Excel.Application();
                var fileContent = string.Empty;
                //   openFileDialog.OpenFile();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // oDoc.InlineShapes.AddPicture(openFileDialog.FileName);
                    using (StreamReader reader = new StreamReader(openFileDialog.FileName))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }


                /*
                 public void SendKeys(object Keys, object Wait);
                var fileContent = string.Empty;
                var filePath = string.Empty;

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files(*.xlsx;*.xls)|*.xlsx;*.xls";
                  //  openFileDialog.FilterIndex = 2;
                //    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;

                        //Read the contents of the file into a stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            fileContent = reader.ReadToEnd();
                        }
                    }
                }
               // */
                MessageBox.Show("Замена текста в выбранном Excel файле была успешно произведена", "Успех");

            }

        }
    }
}
