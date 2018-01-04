using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.VCCodeModel;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using System.Windows.Forms;

namespace RemoveUnusedCodeVSPackage
{
    public class RemoveUnusedCode
    {
        public RemoveUnusedCode(DTE dte)
        {
            this.dte = dte;
            DebugInfoFileName = string.Empty;
            Writer = null;
        }

        public VCCodeModel GetVCCodeModel()
        {
            CodeModel codeModel;
            VCCodeModel vcCodeModel = null;
            Solution solution;
            ErrorMessage = string.Empty;
            if (dte != null)
            {
                solution = dte.Solution;
                if (solution.Count != 0)
                {
                    codeModel = dte.Solution.Item(1).CodeModel;
                    vcCodeModel = (VCCodeModel) codeModel;
                    if (vcCodeModel == null)
                        ErrorMessage = ("The first project is not a VC++ project.");
                }
                else
                    ErrorMessage = "A solution is not open";
            }
            return vcCodeModel;
        }

        public void RemoveUnusedMethods()
        {
            {
                TextDocument ActiveDocument = (TextDocument) dte.ActiveDocument.Object("TextDocument");
                VCCodeModel vcCM = GetVCCodeModel();
                if (vcCM != null)
                {
                    try
                    {
                        if (DebugInfoFileName != string.Empty)
                            Writer = new StreamWriter(DebugInfoFileName);
                        else
                            Writer = null;
                        DateTime StartTime = DateTime.Now;
                        List<VCCodeElement> Elements = ScanElements((VCCodeElements)vcCM.CodeElements);
                        string ProgramText = ActiveDocument.StartPoint.CreateEditPoint().GetText(ActiveDocument.EndPoint);
                        string[] FragmentNames = new string[Elements.Count];
                        string[] FragmentTexts = new string[Elements.Count];
                        Dictionary<string, bool> IsNotTrash = new Dictionary<string, bool>();
                        bool[,] Dependencies = new bool[Elements.Count, Elements.Count];
                        for (int i = 0; i < Elements.Count; ++i)
                        {
                            FragmentNames[i] = Elements[i].Name;
                            FragmentTexts[i] = Elements[i].StartPoint.CreateEditPoint().GetText(Elements[i].EndPoint);
                            if (!IsNotTrash.ContainsKey(Elements[i].Name))
                                IsNotTrash.Add(Elements[i].Name, Elements[i].Name == "main");

                            if (Writer != null)
                            {
                                Writer.WriteLine("FragmentNames[{0}] = {1}", i, FragmentNames[i]);
                                Writer.WriteLine("FragmentTexts[{0}] = {1}", i, FragmentTexts[i]);
                                Writer.WriteLine("Elements[{0}].Kind = {1}", i, Elements[i].Kind.ToString());
                            }
                        }
                        if (Writer != null)
                        {
                            Writer.Close();
                            Writer = null;
                        }
                        ScanElement(ref Dependencies, ref FragmentNames, ref FragmentTexts);
                        Queue<string> q = new Queue<string>();
                        q.Enqueue("main");
                        while (q.Count > 0)
                        {
                            string s = q.Dequeue();
                            for (int i = 0; i < FragmentNames.Length; ++i)
                                if (FragmentNames[i] == s)
                                    for (int j = 0; j < Elements.Count; ++j)
                                        if (Dependencies[i, j] && !IsNotTrash[FragmentNames[j]])
                                        {
                                            IsNotTrash[FragmentNames[j]] = true;
                                            q.Enqueue(FragmentNames[j]);
                                        }
                        }
                        for (int i = 0; i < Elements.Count; ++i)
                            if (!IsNotTrash[FragmentNames[i]])
                                ProgramText = ProgramText.Replace(FragmentTexts[i], "");
                        Regex R = new Regex(@"(?<fragment>\r*\n)\s+\r*\n");
                        ProgramText = R.Replace(ProgramText, @"${fragment}");
                        TimeSpan ElapsedTime = DateTime.Now.Subtract(DateTime.Now);
                        EditPoint StartPoint = ActiveDocument.StartPoint.CreateEditPoint();
                        StartPoint.Delete(ActiveDocument.EndPoint);
                        StartPoint.Insert(ProgramText);
                        dte.StatusBar.Text = "Trash was removed in " + ElapsedTime.TotalMilliseconds.ToString() + " msec";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private List<VCCodeElement> ScanElements(VCCodeElements elements)
        {
            Queue<VCCodeElement> QueueElements = new Queue<VCCodeElement>();
            foreach (VCCodeElement Element in elements)
                QueueElements.Enqueue(Element);
            List<VCCodeElement> result = new List<VCCodeElement>();
            while (QueueElements.Count != 0 )
            {
                VCCodeElement Element = QueueElements.Dequeue();
                switch (Element.Kind)
                {
                    case vsCMElement.vsCMElementStruct:
                        result.Add(Element);
                        VCCodeElements children = (VCCodeElements)Element.Children;
                        foreach (VCCodeElement childElement in children)
                            QueueElements.Enqueue(childElement);
                        break;
                    case vsCMElement.vsCMElementClass:
                    case vsCMElement.vsCMElementEnum:
                    case vsCMElement.vsCMElementFunction:
                    case vsCMElement.vsCMElementMacro:
                    case vsCMElement.vsCMElementTypeDef:
                    case vsCMElement.vsCMElementVariable:
                        result.Add(Element);
                        break;
                    case vsCMElement.vsCMElementIncludeStmt:
                        break;
                    default:
                        MessageBox.Show(Element.Kind.ToString());
                        break;
                }
            }
            return result;
        }

        private void ScanElement(ref bool[,] Dependencies, ref string[] FragmentNames, ref string[] FragmentTexts)
        {
            Regex IsAlphaNum = new Regex(@"[^\w\d]", RegexOptions.Compiled);
            for (int j = 0; j < FragmentNames.Length; ++j)
            {
                string S = FragmentTexts[j];
                HashSet<string> H = new HashSet<string>();
                for (int i = 0; i < FragmentNames.Length; ++i)
                {
                    string Pattern = FragmentNames[i];
                    int k = S.IndexOf(Pattern), n = Pattern.Length;
                    while (k >= 0)
                    {
                        if (k == 0 || IsAlphaNum.IsMatch(S[k - 1].ToString()) && (k + n == S.Length || IsAlphaNum.IsMatch(S[k + n].ToString())))
                        {
                            H.Add(Pattern);
                            k = -1;
                        }
                        else
                            k = S.IndexOf(Pattern, k + 1);
                    }
                }
                for (int i = 0; i < FragmentNames.Length; ++i)
                    if (H.Contains(FragmentNames[i]))
                        Dependencies[j, i] = true;
            }
        }

        private DTE dte;
        public string ErrorMessage
        {
            get;
            private set;
        }

        public string DebugInfoFileName
        {
            get;
            set;
        }
        private StreamWriter Writer;
    }
}
