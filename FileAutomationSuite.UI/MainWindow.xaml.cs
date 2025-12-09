using FileAutomationSuite.UI.Views;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FileAutomationSuite.UI;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void NavButton_Click(object sender, RoutedEventArgs e)
    {
        Button btn = sender as Button;
        string tag = btn.Tag.ToString();

        switch (tag)
        {
            case "ExcelFilter":
                MainContent.Navigate(new ExcelFilterControl());
                break;

            case "ExcelToBCP":
                MainContent.Content = new TextBlock { Text = "Excel → BCP File Page (Coming Soon)", FontSize = 20 };
                break;

            case "BCPToDB":
                MainContent.Content = new TextBlock { Text = "BCP File → DB Page (Coming Soon)", FontSize = 20 };
                break;

            case "DBTableToBCP":
                MainContent.Content = new TextBlock { Text = "Database Table → BCP Page (Coming Soon)", FontSize = 20 };
                break;

            case "ExcelToDB":
                MainContent.Content = new TextBlock { Text = "Excel → Database Page (Coming Soon)", FontSize = 20 };
                break;

            case "DBToExcel":
                MainContent.Content = new TextBlock { Text = "Database → Excel Page (Coming Soon)", FontSize = 20 };
                break;

            case "DBToBCP":
                MainContent.Content = new TextBlock { Text = "Database → BCP Page (Coming Soon)", FontSize = 20 };
                break;
        }
    }
}