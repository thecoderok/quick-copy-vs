//------------------------------------------------------------------------------
// <copyright file="CopyFullyQualifiedNamePackage.cs" company="Tableau">
//     Copyright (c) Tableau.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using EnvDTE;
using System.Windows.Forms;

namespace VitaliiGanzha.VisualStudio.CopyFullyQualifiedNameExtension
{
    //https://msdn.microsoft.com/en-us/library/envdte.filecodemodel.codeelementfrompoint.aspx
    //http://www.diaryofaninja.com/blog/2014/02/18/who-said-building-visual-studio-extensions-was-hard
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}")]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(CopyFullyQualifiedNamePackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class CopyFullyQualifiedNamePackage : Package, IDisposable
    {
        public static CopyFullyQualifiedNamePackage instance;
        /// <summary>
        /// CopyFullyQualifiedNamePackage GUID string.
        /// </summary>
        public const string PackageGuidString = "29a7ddd4-1490-4959-9781-5dfd1bd0957a";

        /// <summary>
        /// Initializes a new instance of the <see cref="CopyFullyQualifiedNamePackage"/> class.
        /// </summary>
        /// 
        private static DTE2 dte;
        public CopyFullyQualifiedNamePackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
            
        }

        public void RunShit()
        {
            

            try
            {
                TextSelection sel =
                    (TextSelection)dte.ActiveDocument.Selection;
                TextPoint pnt = (TextPoint)sel.ActivePoint;

                // Discover every code element containing the insertion point.
                FileCodeModel fcm =
                    dte.ActiveDocument.ProjectItem.FileCodeModel;
                string elems = "";
                vsCMElement scopes = 0;

                foreach (vsCMElement scope in Enum.GetValues(scopes.GetType()))
                {
                    CodeElement elem = null;
                    try
                    {
                        elem = fcm.CodeElementFromPoint(pnt, scope);
                    }
                    catch(Exception ex)
                    {

                    }
                    

                    if (elem != null)
                        elems += elem.Name + " (" + scope.ToString() + ")\n";
                }

                MessageBox.Show(
                    "The following elements contain the insertion point:\n\n" +
                    elems);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
            dte = (DTE2)GetService(typeof(DTE));
        }

        public void Dispose()
        {
        }

        #endregion
    }
}
