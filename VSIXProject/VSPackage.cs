using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.LanguageServices;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace VSIXProject
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(VSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string)]
    public sealed class VSPackage : AsyncPackage
    {
        /// <summary>
        /// VSPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "20f94b8a-cd36-451c-bf00-f18f7a3654f8";

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage"/> class.
        /// </summary>
        public VSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <inheritdoc />
        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            var componentModel = (IComponentModel)await this.GetServiceAsync(typeof(SComponentModel));
            var workspace = componentModel.GetService<VisualStudioWorkspace>();

            workspace.DocumentOpened += Workspace_DocumentOpenedAsync;
            workspace.DocumentClosed += Workspace_DocumentClosedAsync;

            await this.WriteSolutionInfoAsync(workspace.CurrentSolution);
        }

        private async System.Threading.Tasks.Task WriteSolutionInfoAsync(Solution currentSolution)
        {
            await this.WriteToOutputWindowAsync("Projects in the solution: " + Environment.NewLine);

            foreach (var project in currentSolution.Projects)
            {
                await this.WriteToOutputWindowAsync("File path: " + project.FilePath + Environment.NewLine);
                await this.WriteToOutputWindowAsync("Language : " + project.Language + Environment.NewLine);
                await this.WriteToOutputWindowAsync("Output   : " + project.OutputFilePath + Environment.NewLine);
                await this.WriteToOutputWindowAsync(Environment.NewLine);
            }
        }

        private async void Workspace_DocumentClosedAsync(object sender, Microsoft.CodeAnalysis.DocumentEventArgs e)
        {
            await this.WriteToOutputWindowAsync("Closed document: " + e.Document.Name + Environment.NewLine);
            await this.WriteToOutputWindowAsync("Owner project  : " + e.Document.Project.FilePath + Environment.NewLine);
        }

        private async void Workspace_DocumentOpenedAsync(object sender, Microsoft.CodeAnalysis.DocumentEventArgs e)
        {
            await this.WriteToOutputWindowAsync("Opened document: " + e.Document.Name + Environment.NewLine);
            await this.WriteToOutputWindowAsync("Owner project  : " + e.Document.Project.FilePath + Environment.NewLine);
        }

        private async System.Threading.Tasks.Task WriteToOutputWindowAsync(string output)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync();

            var outputWindow = await this.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;

            Guid paneGuid = VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            int hr = outputWindow.GetPane(ref paneGuid, out IVsOutputWindowPane pane);

            if (ErrorHandler.Failed(hr) || (pane == null))
            {
                if (ErrorHandler.Succeeded(outputWindow.CreatePane(ref paneGuid, "General", 1, 1)))
                {
                    hr = outputWindow.GetPane(ref paneGuid, out pane);
                }
            }

            if (ErrorHandler.Succeeded(hr))
            {
                pane?.Activate();
                pane?.OutputString(output);
            }
        }
    }
}
