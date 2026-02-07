using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();
    }

    private void Button_Yana_Click(object sender, RoutedEventArgs e)
    {
        var window = new _4333_Kramarovskaya(); 
        window.Show();
    }
}
