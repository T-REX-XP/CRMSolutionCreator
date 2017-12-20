using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Reflection;
using System.IO;
using System.Globalization;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;

namespace DynamicsCRMSolutionCreator
{
        /// <summary>
        /// App
        /// </summary>
        public partial class App : System.Windows.Application
        {

            /////// <summary>
            /////// InitializeComponent
            /////// </summary>
            ////[System.Diagnostics.DebuggerNonUserCodeAttribute()]
            ////[System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
            ////public void InitializeComponent()
            ////{

            ////    #line 4 "..\..\App.xaml"
            ////    this.StartupUri = new System.Uri("MainWindow.xaml", System.UriKind.Relative);

            ////    #line default
            ////    #line hidden
            ////}

            /////// <summary>
            /////// Application Entry Point.
            /////// </summary>
            ////[System.STAThreadAttribute()]
            ////[System.Diagnostics.DebuggerNonUserCodeAttribute()]
            ////[System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
            ////public static void Main()
            ////{
            ////    MainWindow.App app=new MainWindow.App();
            ////    app.InitializeComponent();
            ////    app.Run();
            ////}
        }
    
    public class Program
    {
        static Dictionary<string, Assembly> assemblyDictionary = new Dictionary<string, Assembly>();
        [STAThreadAttribute]
        public static void Main()
        {
            
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
            //Assembly executingAssembly = Assembly.GetExecutingAssembly();
            //string[] resources = executingAssembly.GetManifestResourceNames();
            //foreach (string resource in resources)
            //{
            //    if (resource.EndsWith(".dll"))
            //    {
            //        using (Stream stream = executingAssembly.GetManifestResourceStream(resource))
            //        {
            //            if (stream == null)
            //                continue;

            //            byte[] assemblyRawBytes = new byte[stream.Length];
            //            stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
            //            try
            //            {
            //                assemblyDictionary.Add(resource, Assembly.Load(assemblyRawBytes));
            //            }
            //            catch (Exception ex)
            //            {
            //                System.Diagnostics.Debug.Print("Failed to load: " + resource + " Exception: " + ex.Message);
            //            }
            //        }
            //    }
            //}
            App.Main();
        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly parentAssembly = Assembly.GetExecutingAssembly();
 
            var name = args.Name.Substring(0, args.Name.IndexOf(',')) + ".dll";
            var resourceName = parentAssembly.GetManifestResourceNames()
                .First(s => s.EndsWith(name));

            using (Stream stream = parentAssembly.GetManifestResourceStream(resourceName))
            {
                byte[] block = new byte[stream.Length];
                stream.Read(block, 0, block.Length);
                return Assembly.Load(block);
            }
            //Assembly executingAssembly = Assembly.GetExecutingAssembly();
            //AssemblyName assemblyName = new AssemblyName(args.Name);

            //string path = assemblyName.Name + ".dll";

            //if (assemblyDictionary.ContainsKey(path))
            //{
            //    return assemblyDictionary[path];
            //}
            //return null;
        }
    }
}


