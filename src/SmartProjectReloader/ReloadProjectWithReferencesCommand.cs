using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Build.Execution;

namespace SmartProjectReloader
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ReloadProjectWithReferencesCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4130;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("9449a163-323a-420f-a09d-561b0e5f8a65");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly SmartProjectReloaderCommandPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReloadProjectWithReferencesCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ReloadProjectWithReferencesCommand(SmartProjectReloaderCommandPackage package)
        {
            this.package = package ?? throw new ArgumentNullException("package");

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ReloadProjectWithReferencesCommand Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => this.package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(SmartProjectReloaderCommandPackage package)
        {
            Instance = new ReloadProjectWithReferencesCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            package.ReloadSelectedProjectsWithReferences();
        }
    }
}
