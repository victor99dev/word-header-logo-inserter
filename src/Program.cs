using Cabecalho.LogoInserter.Controllers;
using Cabecalho.LogoInserter.Services;
using Cabecalho.LogoInserter.Views;

namespace Cabecalho.LogoInserter;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();

        var view = new MainForm();
        var service = new WordLogoService();
        _ = new MainController(view, service);

        Application.Run(view);
    }
}
