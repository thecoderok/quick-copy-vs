//------------------------------------------------------------------------------
// <copyright file="CopyFullyQualifiedNamePackage.cs">
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using EnvDTE80;
using EnvDTE;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;

namespace VitaliiGanzha.VisualStudio.CopyQualifiedNameExtension
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
        private static CopyFullyQualifiedNamePackage instance;
        /// <summary>
        /// CopyFullyQualifiedNamePackage GUID string.
        /// </summary>
        public const string PackageGuidString = "29a7ddd4-1490-4959-9781-5dfd1bd0957a";

        private readonly Lazy<string> NameForLogging;

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
            instance = this;
            NameForLogging = new Lazy<string>(() => this.GetType().FullName);
        }

        public static CopyFullyQualifiedNamePackage Get()
        {
            return instance;
        }

        public void PerformCopy()
        {
            var scheduler = TaskScheduler.FromCurrentSynchronizationContext();
            System.Threading.Tasks.Task.Factory.StartNew(CopyFullyQualifiedNameInternal, CancellationToken.None, TaskCreationOptions.None, scheduler);
        }

        private void CopyFullyQualifiedNameInternal()
        {
            try
            {
                TextSelection sel =
                    (TextSelection)dte.ActiveDocument.Selection;
                TextPoint pnt = sel.ActivePoint as TextPoint;
                if (pnt == null)
                {
                    // Not in editor? Exit.
                    ActivityLog.LogWarning(NameForLogging.Value, "Can't obtain TextPoint, returning.");
                }

                // Discover every code element containing the insertion point.
                FileCodeModel fcm =
                    dte.ActiveDocument.ProjectItem.FileCodeModel;
                vsCMElement scopes = 0;

                // Super-dumb way of getting fully qualified name - find longest full name
                // Seems to be working. It is late and I don't want to invent some heuristic that depend on the Scope type
                string longestName = string.Empty;

                foreach (vsCMElement scope in Enum.GetValues(scopes.GetType()))
                {
                    CodeElement elem = null;
                    try
                    {
                        elem = fcm.CodeElementFromPoint(pnt, scope);

                        if (elem != null)
                        {
                            var name = elem.FullName;
                            if (name.Length > longestName.Length)
                            {
                                longestName = name;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ActivityLog.LogInformation(this.GetType().FullName, "Unable to get code element, exception: " + ex.Message);
                        continue;
                    }
                }

                if (!string.IsNullOrWhiteSpace(longestName))
                {
                    Clipboard.SetText(longestName);
                }
            }
            catch (Exception ex)
            {
                ActivityLog.LogError(this.GetType().FullName, "Error when trying to copy fully qualified name: " + ex.Message);
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
