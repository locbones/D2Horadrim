namespace D2_Horadrim
{
    internal static class Program
    {
        // The main entry point for the application.
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new D2Horadrim());
        }
    }
}