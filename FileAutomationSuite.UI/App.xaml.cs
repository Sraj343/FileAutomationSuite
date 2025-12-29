using FileAutomationSuite.Core.BCP;
using FileAutomationSuite.Core.Excel;
using FileAutomationSuite.Core.SQL;
using FileAutomationSuite.Infrastructure.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using System.Configuration;
using System.Data;
using System.Windows;

namespace FileAutomationSuite.UI;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    public static IServiceProvider ServiceProvider { get; private set; }

    protected override void OnStartup(StartupEventArgs e)
    {
        var services = new ServiceCollection();

        services.AddSingleton<IExcelService, ProcessExcel>();
        //services.AddSingleton<ISQLService, SQLProcess>();
        services.AddSingleton<IBCPService, BCPProcess>();

        ServiceProvider = services.BuildServiceProvider();

        base.OnStartup(e);
    }
}

