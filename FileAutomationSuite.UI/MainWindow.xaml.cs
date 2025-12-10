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
                MainContent.Navigate(new ExcelToBCP());
                break;

            case "BCPToDB":
                MainContent.Navigate(new BCPToDataBase());
                break;

            case "DBTableToBCP":
                MainContent.Navigate(new DBTableToBCP());
                break;

            case "ExcelToDBTable":
                MainContent.Navigate(new ExcelToDBTable());
                break;

            case "DBTableToExcel":
                MainContent.Navigate(new DBTableToExcel());
                break;

            case "DBToBCP":
                MainContent.Navigate(new DBToBCP());
                break;
        }
    }
}