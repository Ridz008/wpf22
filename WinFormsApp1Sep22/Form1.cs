namespace WinFormsApp1Sep22
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn1_Click(object sender, EventArgs e)
        {
            this.label1.Text = "I'm changed";
        }
    }
}