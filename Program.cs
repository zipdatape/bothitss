using System.Text;
using NotificadorBajasHitssApp;

namespace NotificadorBajasHitssApp;

static class Program
{
    [STAThread]
    static void Main()
    {
        // Necesario para Encoding.GetEncoding("iso-8859-15") en .NET 5+
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}
