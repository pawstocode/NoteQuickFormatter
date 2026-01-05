using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Linq;

namespace NoteQuickFormatter
{
    public partial class NextMonthToDoTemplateDialog : Form
    {
        private readonly OneNoteService _oneNoteService = new OneNoteService();
        public NextMonthToDoTemplateDialog()
        {
            InitializeComponent();
            _oneNoteService.RefreshHierarchy();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            if (_oneNoteService.CurrentlyViewedNotebook is null)
            {
                MessageBox.Show("OneNote is not open.");
                Close();
            }
            else
            {
                comboBox1.Items.AddRange(_oneNoteService.NotebookNames);
                comboBox1.SelectedText = _oneNoteService.CurrentlyViewedNotebook.Attribute("nickname").Value;
                int currentMonth = DateTime.Today.Month;
                textBox1.Text = new DateTimeHelper().AbbreviatedMonthNames[currentMonth];
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _oneNoteService.CreateNewSection(textBox1.Text, comboBox1.Text, true);
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
    class Notebook
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public bool IsCurrentlyViewed { get; set; }
    }
    class ElementNameComparer : IEqualityComparer<XElement>
    {
        public bool Equals(XElement x, XElement y)
        {
            return x.Attribute("name").Value == y.Attribute("name").Value;
        }

        public int GetHashCode(XElement obj)
        {
            throw new NotImplementedException();
        }
    }
}
