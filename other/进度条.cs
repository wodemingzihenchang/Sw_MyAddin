using System.Windows.Forms;

namespace Sw_MyAddin
{
    public partial class 进度条 : Form
    {
        public 进度条()
        {
            InitializeComponent();
        }
        public void 数字显示(int i, int Totality)
        {
            this.label1.Text = i + "/" + Totality;
        }
        public void 文字显示(string s)
        {
            this.label1.Text = s;
        }
    }
}
