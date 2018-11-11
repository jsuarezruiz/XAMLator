using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace XAMLator.Client.VisualStudio
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    // The following line will schedule the package to be initialized when a solution is being opened
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class XAMLatorPackage : AsyncPackage
    {
        public const string PackageGuidString = "c3dd450f-35c5-44e5-ab40-51960e8c9acb";

        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // Since this package might not be initialized until after a solution has finished loading,
            // we need to check if a solution has already been loaded and then handle it.
            bool isSolutionLoaded = await IsSolutionLoadedAsync();

            if (isSolutionLoaded)
            {
                HandleOpenSolution();
            }

            // Listen for subsequent solution events.
            Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnAfterOpenSolution += HandleOpenSolution;
        }

        private async Task<bool> IsSolutionLoadedAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            // Check if a solution has loaded.
            var solService = await GetServiceAsync(typeof(SVsSolution)) as IVsSolution;

            Assumes.Present(solService);
            ErrorHandler.ThrowOnFailure(solService.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object value));

            return value is bool isSolOpen && isSolOpen;
        }

        private async void HandleOpenSolution(object sender = null, EventArgs e = null)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            DTE dte = (DTE)await GetServiceAsync(typeof(DTE));

            // Initialize Visual Studio XAMLatorMonitor.
            XAMLatorMonitor.Init(new VisualStudioIDE(dte));

            // Start monitoring.
            XAMLatorMonitor.Instance.StartMonitoring();
        }
    }
}
