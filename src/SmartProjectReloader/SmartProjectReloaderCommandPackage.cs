using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using Microsoft.Build.Evaluation;

namespace SmartProjectReloader
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
    [InstalledProductRegistration("#2110", "#2112", "1.0", IconResourceID = 2400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(SmartProjectReloaderCommandPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed partial class SmartProjectReloaderCommandPackage : Package
    {
        /// <summary>
        /// SmartProjectLoaderCommandPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "d24ede67-c86e-4aa6-9bb4-92978d3f5955";

        /// <summary>
        /// Initializes a new instance of the <see cref="SmartProjectLoaderCommand"/> class.
        /// </summary>
        public SmartProjectReloaderCommandPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            UnloadAllProjectsCommand.Initialize(this);
            ReloadProjectWithReferencesCommand.Initialize(this);
            base.Initialize();
            ReloadAllProjectsCommand.Initialize(this);
        }

        #endregion

        public void UnloadAllProjects()
        {
            var solution = (IVsSolution)GetService(typeof(SVsSolution));
            if (solution == null) return;
            var solution4 = (IVsSolution4)solution;

            var guid = Guid.Empty;
            solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, ref guid, out var enumerator);
            IVsHierarchy[] hierarchy = new IVsHierarchy[1] { null };
            for (enumerator.Reset(); enumerator.Next(1, hierarchy, out var fetched) == VSConstants.S_OK && fetched == 1;)
            {
                solution.GetGuidOfProject(hierarchy[0], out guid);
                solution4.UnloadProject(ref guid, (uint)_VSProjectUnloadStatus.UNLOADSTATUS_UnloadedByUser);
            }
        }

        public void ReloadAllProjects()
        {
            var solution = (IVsSolution)GetService(typeof(SVsSolution));
            if (solution == null) return;
            var solution4 = (IVsSolution4)solution;

            foreach (var project in GetUnloadedProjects())
            {
                solution.GetGuidOfProject(project, out var guid);
                solution4.ReloadProject(ref guid);
            }
        }

        public void ReloadSelectedProjectsWithReferences()
        {
            var solution = (IVsSolution)GetService(typeof(SVsSolution));
            if (solution == null) return;
            var solution4 = (IVsSolution4)solution;

            var selectedProject = GetSelectedProjectHierarchy();
            selectedProject.GetCanonicalName((uint)VSConstants.VSITEMID.Root, out var selectedProjectFilePath);

            var reloadProjectFiles = new HashSet<string>();
            GetReferenceProjectFilesRecursive(selectedProjectFilePath, reloadProjectFiles);
            reloadProjectFiles.Add(selectedProjectFilePath);

            var unloadedProjects = GetUnloadedProjectHierarchiesWithFilePath();
            foreach (var reloadProjectFile in reloadProjectFiles)
            {
                if (unloadedProjects.TryGetValue(reloadProjectFile, out var reloadProject))
                {
                    solution.GetGuidOfProject(reloadProject, out var guid);
                    solution4.ReloadProject(ref guid);
                }
            }
        }

        private IVsHierarchy GetSelectedProjectHierarchy()
        {
            var selectionMonitor = GetService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            if (selectionMonitor == null) return null;
            selectionMonitor.GetCurrentSelection(out var hierarchyPtr, out var itemID, out var multiSelect, out var containerPtr);
            if (IntPtr.Zero != containerPtr)
            {
                Marshal.Release(containerPtr);
                containerPtr = IntPtr.Zero;
            }
            return (IVsHierarchy)Marshal.GetTypedObjectForIUnknown(hierarchyPtr, typeof(IVsHierarchy));
        }

        private IEnumerable<IVsHierarchy> GetUnloadedProjects()
        {
            var solution = (IVsSolution)GetService(typeof(SVsSolution));
            if (solution == null) yield break;

            var results = new Dictionary<string, IVsHierarchy>();
            var guid = Guid.Empty;
            solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_UNLOADEDINSOLUTION, ref guid, out var enumerator);
            IVsHierarchy[] hierarchy = new IVsHierarchy[1] { null };
            for (enumerator.Reset(); enumerator.Next(1, hierarchy, out var fetched) == VSConstants.S_OK && fetched == 1;)
            {
                hierarchy[0].GetCanonicalName((uint)VSConstants.VSITEMID.Root, out var projFileName);
                yield return hierarchy[0];
            }
        }

        private IDictionary<string, IVsHierarchy> GetUnloadedProjectHierarchiesWithFilePath()
            => GetUnloadedProjects().ToDictionary(proj => { proj.GetCanonicalName((uint)VSConstants.VSITEMID.Root, out var projFileName); return projFileName; });

        private void GetReferenceProjectFilesRecursive(string projectFile, HashSet<string> referenceProjectFiles, ProjectCollection projectCollection = null)
        {
            var requiresDispose = false;
            if (projectCollection == null)
            {
                projectCollection = new ProjectCollection();
                requiresDispose = true;
            }

            try
            {
                var baseUri = new Uri(projectFile);
                var proj = projectCollection.LoadProject(projectFile);
                foreach (var refereceProjectPropItem in proj.Items.Where(item => item.ItemType == "ProjectReference"))
                {
                    var refProjFile = new Uri(baseUri, refereceProjectPropItem.EvaluatedInclude).LocalPath;
                    if (referenceProjectFiles.Add(refProjFile))
                    {
                        GetReferenceProjectFilesRecursive(refProjFile, referenceProjectFiles, projectCollection);
                    }
                }
            }
            finally
            {
                if (requiresDispose) projectCollection.Dispose();
            }
        }
    }
}
