using System;
using System.Collections.Generic;
using System.Security;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ResumeBuilder();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ResumeBuilder
{
    [ComVisible(true)]
    public class ResumeBuilder : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ResumeBuilder()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ResumeBuilder.ResumeBuilder.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        public void OnBuildResume(Office.IRibbonControl control)
        {

            using (ClientContext context = new ClientContext("https://contoso.sharepoint.com"))
            {
                SecureString Password = new SecureString();
                foreach (char c in ("P@ssW0rd!").ToCharArray()) Password.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials("admin@contoso.onmicrosoft.com", Password);
                PeopleManager peoplemanager = new PeopleManager(context);
                PersonProperties myprops = peoplemanager.GetMyProperties();
                context.Load(myprops);
                context.ExecuteQuery();

                var props = myprops.UserProfileProperties;

                foreach (Word.Section section in Globals.ThisAddIn.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    headerRange.Text = "Internal Job Posting";
                }

                //Insert Display Name and Title from SharePoint Online Profile
                Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.set_Style("Intense Quote");
                selection.TypeText(myprops.DisplayName + " | " + myprops.Title + Environment.NewLine);
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                selection.TypeText("Email: " + myprops.Email);

                //Insert Profile Picture from SharePoint Online Profile

                selection.TypeParagraph();
                selection.set_Style("Normal");
                selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //selection.InlineShapes.AddPicture(myprops.PictureUrl);
                selection.InlineShapes.AddPicture(myprops.PictureUrl);

                selection.TypeParagraph();
                selection.set_Style("Heading 1");
                selection.TypeText("About Me");

                selection.TypeParagraph();
                selection.set_Style("Normal");
                string AboutMe = "";
                if (props.TryGetValue("AboutMe", out AboutMe))
                {
                    selection.TypeText(AboutMe);
                }

                selection.TypeParagraph();
                selection.set_Style("Heading 2");
                selection.TypeText("Skill");

                selection.TypeParagraph();
                selection.Range.ListFormat.ApplyBulletDefault();
                // selection.set_Style("Normal");
                var SPSSkills = "";
                if (props.TryGetValue("SPS-Skills", out SPSSkills))
                {
                    if (SPSSkills.Contains("|"))
                    {
                        selection.TypeText(SPSSkills.Replace("|", "\n"));
                    }
                    else
                    {
                        selection.TypeText(SPSSkills);
                    }
                }
                selection.TypeParagraph();
                selection.set_Style("Heading 2");
                selection.TypeText("Past Projects");

                selection.TypeParagraph();
                selection.Range.ListFormat.ApplyBulletDefault();
                var PastProject = "";
                if (props.TryGetValue("SPS-PastProjects", out PastProject))
                {
                    if (PastProject.Contains("|"))
                    {
                        selection.TypeText(PastProject.Replace("|", "\n"));
                    }
                    else
                    {
                        selection.TypeText(PastProject);
                    }
                }

                selection.TypeParagraph();
                selection.set_Style("Heading 2");
                selection.TypeText("Employee Information");

                selection.TypeParagraph();
                selection.Range.ListFormat.ApplyBulletDefault();

                var CellPhone = ""; var School = ""; var IM = "";
                var Department = ""; var Languages = "";
                if (props.TryGetValue("CellPhone", out CellPhone)
                    && props.TryGetValue("SPS-School", out School)
                    && props.TryGetValue("SPS-SipAddress", out IM)
                    && props.TryGetValue("SPS-Department", out Department)
                    && props.TryGetValue("SPS-MUILanguages", out Languages))
                {
                    selection.TypeText(string.Concat("Email Address", "\t", "\t", myprops.Email, Environment.NewLine));
                    selection.TypeText(string.Concat("Cell Phone", "\t", "\t", selection.Text = CellPhone, Environment.NewLine));
                    selection.TypeText(string.Concat("School Name", "\t", "\t", selection.Text = School.Replace('.', ' '), Environment.NewLine));
                    selection.TypeText(string.Concat("Instant Messenger", "\t", selection.Text = IM, Environment.NewLine));
                    selection.TypeText(string.Concat("Department", "\t", "\t", selection.Text = Department, Environment.NewLine));
                    selection.TypeText(string.Concat("Languages", "\t", "\t", selection.Text = Languages));
                }
                
                //selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            }
        }
    }
}
