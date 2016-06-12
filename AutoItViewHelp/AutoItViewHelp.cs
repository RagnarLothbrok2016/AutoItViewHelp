//------------------------------------------------------------------------------
//
// Copyright 2016 Christian Lang
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//  http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace AutoItViewHelp
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AutoItViewHelp
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3d8a0665-1f2a-42ef-b478-9e1451762bdb");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutoItViewHelp"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private AutoItViewHelp(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

            if (null != commandService)
            {
                CommandID menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += menuCommand_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AutoItViewHelp Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new AutoItViewHelp(package);
        }

        /// <summary>
        /// This function gets executed before the context menu inside the editor is visible.
        /// Purpose of this is to be able to determine if the "wanted" file type inside the editor was started and
        /// to set the dynamic visibility property for the context menu.
        /// </summary>
        void menuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;

            if (menuCommand != null)
            {
                DTE dte = (DTE)ServiceProvider.GetService(typeof(DTE));
                Document documents = dte.ActiveDocument;
           
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                if (string.IsNullOrEmpty(documents.Name) || documents.Name.Length < 4)
                {
                    return;
                }
                
                string returnedActiveDocumentType = documents.Name.ToLower().Remove(0, documents.Name.Length - 4);
                bool isAutoItFile = returnedActiveDocumentType.Contains(".au3");

                if (!isAutoItFile)
                {
                    return;
                }

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
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
            RegistryKey checkArchKey = Registry.LocalMachine.OpenSubKey(@"Software\Wow6432Node");
            string autoItInstallDir = null;

            if (checkArchKey == null)
            {
                RegistryKey autoItKey = Registry.LocalMachine.OpenSubKey(@"Software\AutoIt v3\AutoIt");

                if (autoItKey == null)
                {
                    VsShellUtilities.ShowMessageBox(
                        this.ServiceProvider,
                        "Detecting your AutoIt version failed!.",
                        "Error.",
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }
                else
                {
                    autoItInstallDir = autoItKey.GetValue("InstallDir").ToString() + @"\AutoIt.chm";
                    autoItKey.Close();
                }
            }
            else
            {
                checkArchKey.Close();
                RegistryKey autoItKey = Registry.LocalMachine.OpenSubKey(@"Software\Wow6432Node\AutoIt v3\AutoIt");
                if (autoItKey == null)
                {
                    VsShellUtilities.ShowMessageBox(
                        this.ServiceProvider,
                        "Detecting your AutoIt version failed!.",
                        "Error.",
                        OLEMSGICON.OLEMSGICON_INFO,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }
                else
                {
                    autoItInstallDir = autoItKey.GetValue("InstallDir").ToString() + @"\AutoIt.chm";
                    autoItKey.Close();
                }
            }

            DTE dte = (DTE)ServiceProvider.GetService(typeof(DTE));
            TextDocument document = (TextDocument)(dte.ActiveDocument.Object("TextDocument"));
            string manualSelectedEditorText = document.Selection.Text.ToString();
            manualSelectedEditorText = System.Text.RegularExpressions.Regex.Replace(manualSelectedEditorText, @"\r\n?|\n", "");

            if (string.IsNullOrEmpty(manualSelectedEditorText) || string.IsNullOrWhiteSpace(manualSelectedEditorText))
            {
                manualSelectedEditorText = "AutoIt";
            }

            try
            {
                System.Windows.Forms.Help.ShowHelp(new System.Windows.Forms.Control(), autoItInstallDir, System.Windows.Forms.HelpNavigator.KeywordIndex, manualSelectedEditorText);
            }
            catch
            {
                VsShellUtilities.ShowMessageBox(
                    this.ServiceProvider,
                    "Launching your AutoIt.chm help file failed!.",
                    "Error.",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }
        }
    }
}
