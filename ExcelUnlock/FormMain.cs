using System;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using ExcelUnlock.Properties;

namespace ExcelUnlock
{
    /// <summary>
    /// The main form of application.
    /// </summary>
    public partial class fMain : Form
    {
        public fMain()
        {
            InitializeComponent();
            fMain_setText();
        }

        private void fMain_setText()
        {
            //display version in AssemblyInfo.cs
            this.Text = $"{Resources.FMTitle}{ System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Major.ToString() }.{ System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Minor.ToString() } " ;

            //set textBox message
            this.tbPath.Text = Resources.FMCoDropItemHere;
        }

        private async void bUnlock_Click(object sender, EventArgs e)
        {
            //unlock button
            bUnlock.Enabled = false;
            try
            {
                tbPath_TextCheck();
                await excelUnlockAsync();
            }
            catch (Exception exc)
            {
                lProgress_displayMessage(null, new ReportArgs { report = (Resources.FMCoError + ": " + exc.Message) });
                MessageBox.Show(Resources.FMCoError + ": " + exc.Message,Resources.FMCoError, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            bUnlock.Enabled = true;
        }

        private void fMain_DragDrop(object sender, DragEventArgs e)
        {
            //When items droped show full path of that item in textbox
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Count() > 0)
            {
                tbPath.Text = fileList[0];
            }
        }

        private void fMain_DragEnter(object sender, DragEventArgs e)
        {
            //allow drag and drop feature
            e.Effect = DragDropEffects.All;
        }
          
        private void tbPath_GhostTextEnter(object sender, EventArgs e)
        {
            //hide ghots text
            if (tbPath.Text == Resources.FMCoDropItemHere)
            {
                tbPath.Text = "";
                tbPath.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void tbPath_GhostTextLeave(object sender, EventArgs e)
        {
            //show ghots text
            if (tbPath.Text == "")
            {
                tbPath.Text = Resources.FMCoDropItemHere;
                tbPath.ForeColor = System.Drawing.Color.DarkGray;
            }
        }

        private void lProgress_displayMessage(object sender, ReportArgs e)
        {
            //display progress
            this.Invoke((MethodInvoker)delegate
            {
                lProgress.Text = e.report;
                this.Refresh();
            });
        }

        private void excelUnlock()
        {
            //create class for unlock excel
            ExcelUnlockClass excelUnlockClass = new ExcelUnlockClass(tbPath.Text);
            excelUnlockClass.reportEvent += lProgress_displayMessage;
            excelUnlockClass.process();
        }

        private async Task excelUnlockAsync()
        {
            //create class for unlock excel
            ExcelUnlockClass excelUnlockClass = new ExcelUnlockClass(tbPath.Text);
            excelUnlockClass.reportEvent += lProgress_displayMessage;
            await Task.Run(() => {
                excelUnlockClass.process();
            });
        }

        private void tbPath_TextCheck()
        {
            //check it 
            if (tbPath.Text.Equals(Resources.FMCoDropItemHere) | tbPath.Text.Equals(""))
            {
                Exception exc = new Exception(Resources.FMCoErrorSetPath);
                throw exc;
            }

        }
    }
}
