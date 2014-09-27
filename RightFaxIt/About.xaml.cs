using System.Windows.Input;

namespace RightFaxIt
{
    /// <summary>
    /// Interaction logic for About.xaml
    /// </summary>
    public partial class About
    {
        public About()
        {
            InitializeComponent();
        }

        private void Url_Link_Clicked(object sender, MouseButtonEventArgs e)
        {

            System.Diagnostics.Process.Start("http://www.github.com/zKarp");
        }
    }
}
