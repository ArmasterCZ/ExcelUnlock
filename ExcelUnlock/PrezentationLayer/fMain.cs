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
        /// <summary>
        /// initializes form items
        /// </summary>
        public fMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// button for unlock
        /// </summary>
        private async void bUnlock_Click(object sender, EventArgs e)
        {
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

        /// <summary>
        /// get path from droped file
        /// </summary>
        private void fMain_DragDrop(object sender, DragEventArgs e)
        {
            //When items droped show full path of that item in textbox
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Count() > 0)
            {
                tbPath.Text = fileList[0];
                tbPath.ForeColor = System.Drawing.Color.Black;
            }
        }

        /// <summary>
        /// allow drag drop
        /// </summary>
        private void fMain_DragEnter(object sender, DragEventArgs e)
        {
            //allow drag and drop feature
            e.Effect = DragDropEffects.All;
        }
          
        /// <summary>
        /// ghost Text hide
        /// </summary>
        private void tbPath_GhostTextEnter(object sender, EventArgs e)
        {
            //hide ghots text
            if (tbPath.Text == Resources.FMCoDropItemHere)
            {
                tbPath.Text = "";
                tbPath.ForeColor = System.Drawing.Color.Black;
            }
        }

        /// <summary>
        /// ghost Text show
        /// </summary>
        private void tbPath_GhostTextLeave(object sender, EventArgs e)
        {
            //show ghots text
            if (tbPath.Text == "")
            {
                tbPath.Text = Resources.FMCoDropItemHere;
                tbPath.ForeColor = System.Drawing.Color.DarkGray;
            }
        }

        /// <summary>
        /// display progress label
        /// </summary>
        private void lProgress_displayMessage(object sender, ReportArgs e)
        {
            this.Invoke((MethodInvoker)delegate
            {
                lProgress.Text = e.report;
                this.Refresh();
            });
        }

        /// <summary>
        /// create class for unlock excel
        /// </summary>
        private void excelUnlock()
        {
            //create class for unlock excel
            ExcelUnlockClass excelUnlockClass = new ExcelUnlockClass(tbPath.Text);
            excelUnlockClass.reportEvent += lProgress_displayMessage;
            excelUnlockClass.process();
        }

        /// <summary>
        /// create task for unlock excel => responsive UI.
        /// </summary>
        /// <returns></returns>
        private async Task excelUnlockAsync()
        {
            //create class for unlock excel
            ExcelUnlockClass excelUnlockClass = new ExcelUnlockClass(tbPath.Text);
            excelUnlockClass.reportEvent += lProgress_displayMessage;
            await Task.Run(() => {
                excelUnlockClass.process();
            });
        }

        /// <summary>
        /// check if tbPath contain correct text
        /// </summary>
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
