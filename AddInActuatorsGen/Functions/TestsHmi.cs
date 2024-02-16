using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW;
using Siemens.Engineering;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TiaHelperLibrary.TiaHelper;
using TiaHelperLibrary.Models;
using Siemens.Engineering.Hmi;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;
using TiaXmlGenerator.Models;
using TiaXmlGenerator;
using static System.Net.Mime.MediaTypeNames;

namespace AddInActuatorsGen.Functions
{
    public class TestsHmi
    {
        TiaPortal _tiaportal;
        public ExclusiveAccess tiaMessage;
        public List<CultureInfo> projectLanguages;


        public TestsHmi(TiaPortal tiaportal)
        {
            _tiaportal = tiaportal;
        }

        public void TestsHmiGen(MenuSelectionProvider<PlcBlock>
            menuSelectionProvider)
        {
            /// 
            ///
            /// Getting PlcSoftware object from menu selection
            ///
            ///

            tiaMessage = _tiaportal.ExclusiveAccess("Odczytywanie tagów...");

            PlcSoftware plcSoftware = GetPlcSoftware(menuSelectionProvider);
            HmiTarget hmiSoftware = GetHmiTarget(_tiaportal);

            /// Active languages
            /// 
            projectLanguages = GetProjectCultures(_tiaportal.Projects.FirstOrDefault());

            /// Chosen DB block
            /// 
            PlcBlock block = menuSelectionProvider.GetSelection<PlcBlock>().FirstOrDefault();


            /// Exporting block
            /// 
            FileInfo xmlFilePath = new FileInfo(Path.GetTempFileName() + ".xml");

            if (!block.IsConsistent)
            {
                MessageBox.Show("Blok nieskompilowany!");
                return;
            }

            if (block is InstanceDB)
            {
                MessageBox.Show("Instancje nie są obsługiwane!");
                return;
            }

            block.Export(xmlFilePath, ExportOptions.WithDefaults);

            /// Reading xml file to string
            ///
            string xmlData = File.ReadAllText(xmlFilePath.FullName);
            xmlFilePath.Delete();

            /// Serialize xml to class
            /// 
            XmlSerializer serializer = new XmlSerializer(typeof(DbModel));


            /// text - all texts from db for debugging
            /// prefix - if nested structure occurs prefix from parent objects name is used
            string text = string.Empty;
            string prefix = string.Empty;


            using (StringReader reader = new StringReader(xmlData))
            {
                DbModel document = (DbModel)serializer.Deserialize(reader);


                MemberRecurrence(
                    document.SWBlocksGlobalDB.AttributeList.Interface.Sections.Section.Member,
                    ref text,
                    ref prefix
                    );
            }

            MessageBox.Show(text);
        }


        private void MemberRecurrence(SectionsSectionMember[] members, ref string membersText, ref string prefix)
        {
            string entry_text = string.Empty;
            if (members.Length > 0)
            {
                tiaMessage.Text = $"Kopiowanie zawartości DB... {prefix}";
                CheckCancellation();
            }


            foreach (SectionsSectionMember member in members)
            {
                membersText += member.Name + "\t" + member.Datatype + "\n";

                string subprefix = prefix;

                if (member.Member != null)
                {
                    subprefix += member.Name + " - ";
                    MemberRecurrence(member.Member, ref membersText, ref subprefix);
                }

                if (member.Subelement != null)
                {
                    subprefix += member.Name;
                    MemberRecurrence(member.Subelement, ref membersText, ref subprefix);
                }

                if (member.Sections != null)
                {
                    subprefix += member.Name + " - ";
                    MemberRecurrence(member.Sections.Section.Member, ref membersText, ref subprefix);
                }
            }
        }



        private void MemberRecurrence(SectionsSectionMemberSubelement[] members, ref string membersText, ref string prefix)
        {
            string entry_text = string.Empty;
            //CheckCancellation();

            foreach (SectionsSectionMemberSubelement subElement in members)
            {
                membersText += prefix + "[" + subElement.Path + "]" + "\t" + subElement.StartValue + "\n";
                
                if (subElement.Comment != null)
                {

                }
                else
                {
                    entry_text = SeparateWords(prefix + "[" + subElement.Path + "]");
                }
            }
        }


        public bool CheckCancellation()
        {
            if (tiaMessage.IsCancellationRequested)
            {
                throw new Exception("Addin canceled");
            }
            return false;
        }

    }
}
