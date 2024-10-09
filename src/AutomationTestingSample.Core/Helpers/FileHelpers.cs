using System.Reflection;

namespace AutomationTestingSample.Core.Helpers
{
    public class FileHelpers
    {
        public static string ProjectPath => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"..\..\..\..\";

        public static string Actuals_Downloads { get; private set; } 
        public static string Actuals_Screenshots { get; private set; } 
        public static string Baselines_Screenshots { get; private set; } 
        public static string Uploads { get; private set; }

        public FileHelpers()
        {
            Actuals_Downloads = ProjectPath + @"Resources\Actuals\Downloads\";
            Actuals_Screenshots = ProjectPath + @"Resources\Actuals\Screenshots\";
            Baselines_Screenshots = ProjectPath + @"Resources\Baselines\Screenshots\";
            Uploads = ProjectPath + @"Resources\Uploads";
            
            CreateDirectoryIfNotExist(Actuals_Downloads);
            CreateDirectoryIfNotExist(Actuals_Screenshots);
            CreateDirectoryIfNotExist(Baselines_Screenshots);
            CreateDirectoryIfNotExist(Uploads);
        }

        public static void Initialize()
        {
            _ = new FileHelpers();
        }

        private static void CreateDirectoryIfNotExist(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
