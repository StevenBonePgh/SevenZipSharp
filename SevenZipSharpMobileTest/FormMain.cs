using System;
using System.IO;
using System.Windows.Forms;
using SevenZip;

namespace SevenZipSharpMobileTest
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void menuItem1_Click(object sender, System.EventArgs e)
        {
            var executingPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            
            tb_Log.Text += "Performing an internal benchmark..." + Environment.NewLine;
            var features = SevenZip.SevenZipExtractor.CurrentLibraryFeatures;
            tb_Log.Text += string.Format("Finished. The score is {0}{2}{1}{2}", ((uint)features).ToString("X6"), features, Environment.NewLine);

            var fileName = executingPath + "\\gpl.7z";
            tb_Log.Text += "File name is \"" + fileName + "\"" + Environment.NewLine;
            try
            {
                using (var extr = new SevenZipExtractor(fileName))
                {
                    tb_Log.Text += "The archive was successfully identified. Ready to extract" + Environment.NewLine;
                    try
                    {
                        extr.ExtractArchive(executingPath);
                        tb_Log.Text += "Extracted successfully!" + Environment.NewLine;
                    }
                    catch (Exception exception)
                    {
                        tb_Log.Text += string.Format("Extract failed!{1}{2}{1}", exception, Environment.NewLine);
                    }
                }
            }
            catch (Exception exception)
            {
                tb_Log.Text += string.Format("Extract failed!{1}{2}{1}", exception, Environment.NewLine);
            }
        }
    }
}