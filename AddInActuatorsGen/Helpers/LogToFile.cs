using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Siemens.Engineering;

namespace AddinActuatorsGen.Helpers
{
    public class LogToFile
    {
        public static void Save(Exception ex)
        {
            switch (ex.GetType())
            {
                default:
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Pliki tekstowe (*.txt)|*.txt";
                    saveFileDialog.Title = "Zapisz plik log";
                    saveFileDialog.FileName = "log.txt";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        try
                        {
                            using (StreamWriter writer = new StreamWriter(filePath))
                            {
                                writer.WriteLine("Data: " + DateTime.Now);
                                writer.WriteLine("Wyjątek: " + ex.GetType());
                                writer.WriteLine("Wiadomość: " + ex.Message);
                                writer.WriteLine("Źródło: " + ex.Source);
                                writer.WriteLine("Źródło trace: " + ex.StackTrace);

                                // Dodaj dodatkowe informacje, jeśli są dostępne
                                if (ex.InnerException != null)
                                {
                                    writer.WriteLine("Wewnętrzny wyjątek:");
                                    writer.WriteLine("Wyjątek: " + ex.ToString());
                                    writer.WriteLine("Wiadomość: " + ex.InnerException.Message);
                                    writer.WriteLine("Źródło: " + ex.InnerException.Source);
                                    writer.WriteLine("Źródło trace: " + ex.InnerException.StackTrace);
                                }

                                writer.WriteLine();
                                writer.WriteLine("--- Koniec logu ---");
                            }

                            MessageBox.Show("Plik log został zapisany.");
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Wystąpił błąd podczas zapisu pliku log.");
                        }
                    }
                    return;
            }
        }
    }
}
