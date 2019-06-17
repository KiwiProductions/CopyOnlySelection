using System;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.VisualStudio.OLE.Interop;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace CopyOnlySelection
{
    static class Helpers
    {
        private static Dictionary<string, string> keyRemap = new Dictionary<string, string>()
        {
            { "Ctrl+", "^"},
            { "Alt+", "%"},
            { "Bkspce", "{BKSP}" },
            { "Break", "{BREAK}" },
            { "Del", "{DEL}" },
            { "Down Arrow", "{DOWN}" },
            { "End", "{END}" },
            { "Enter", "{ENTER}" },
            { "Esc", "{ESC}" },
            { "F1", "{F1}" },
            { "F2", "{F2}" },
            { "F3", "{F3}" },
            { "F4", "{F4}" },
            { "F5", "{F5}" },
            { "F6", "{F6}" },
            { "F7", "{F7}" },
            { "F8", "{F8}" },
            { "F9", "{F9}" },
            { "F10", "{F10}" },
            { "F11", "{F11}" },
            { "F12", "{F12}" },
            { "F13", "{F13}" },
            { "F14", "{F14}" },
            { "F15", "{F15}" },
            { "F16", "{F16}" },
            { "Home", "{HOME}" },
            { "Ins", "{Ins}" },
            { "Left Arrow", "{LEFT}" },
            { "PgDn", "{PGDN}" },
            { "PgUp", "{PGUP}" },
            { "Right Arrow", "{RIGHT}" },
            { "Tab", "{TAB}" },
            { "Up Arrow", "{UP}" },
            { "Num +", "{ADD}" },
            { "Num -", "{SUBTRACT}" },
            { "Num *", "{MULTIPLY}" },
            { "Num /", "{DIVIDE}" },
            { "Shift+", "+" },
            { ", ", "" }
        };

        private static Regex keyRegex = new Regex(@"Ctrl\+|Alt\+|Bkspce|Break|Del|Down Arrow|End|Enter|Esc|F1|F2|F3|F4|F5|F6|F7|F8|F9|F10|F11|F12|F13|F14|F15|F16|Home|Ins|Left Arrow|PgDn|PgUp|Right Arrow|Tab|Up Arrow|Num \+|Num \-|Num \*|Num /|Shift\+|\r ");

        private static MatchEvaluator matchEvaluator = new MatchEvaluator(ReplaceSpecialKeys);

        private static string ReplaceSpecialKeys(Match m)
        {
            try
            {
                return keyRemap[m.Value];
            }
            catch (Exception) { return string.Empty; }
        }

        private static DTE2 _dte = null;
        public static DTE2 GetDTE()
        {
            if (_dte == null)
            {
                _dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
            }
            return _dte;
        }

        public static void ExecuteCommand(string command, string commandArgs = "")
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE2 dte = GetDTE();
            try
            {
                dte.ExecuteCommand(command, commandArgs);
            }
            catch (COMException)
            {
                Debug.WriteLine("Could not execute command.");
                try
                {
                    EnvDTE.Command cmd = dte.Commands.Item(command);
                    if (cmd == null) return;
                    if (cmd.Bindings is object[] bindings)
                    {
                        if (bindings.Length > 0)
                        {
                            string[] split = ((string)bindings[0]).Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
                            if (split.Length > 1)
                            {
                                string keys = keyRegex.Replace(split[1], matchEvaluator);
                                SendKeys.Send(keys);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }
        }

        public static IWpfTextView GetCurentTextView()
        {
            var componentModel = GetComponentModel();
            if (componentModel == null) return null;
            ThreadHelper.ThrowIfNotOnUIThread();
            var vsTextView = GetCurrentNativeTextView();
            if (vsTextView == null) return null;
            var editorAdapter = componentModel.GetService<IVsEditorAdaptersFactoryService>();
            return editorAdapter.GetWpfTextView(vsTextView);
        }

        public static IVsTextView GetCurrentNativeTextView()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var textManager = (IVsTextManager)ServiceProvider.GlobalProvider.GetService(typeof(SVsTextManager));
            Assumes.Present(textManager);
            IVsTextView activeView = null;
            textManager.GetActiveView(1, null, out activeView);
            return activeView;
        }

        public static IComponentModel GetComponentModel()
        {
            return (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
        }
    }
}
