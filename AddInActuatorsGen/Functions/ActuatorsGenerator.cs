using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW;
using Siemens.Engineering;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static TiaHelperLibrary.Models.Tia.TagTableXml;
using static TiaHelperLibrary.TiaHelper;
using System.Xml.Serialization;
using TiaHelperLibrary.Models.Tia;
using TiaXmlGenerator.Models;
using TiaXmlGenerator;
using System.Globalization;

namespace AddInActuatorsGen.Functions
{
    public class ActuatorsGenerator
    {
        TiaPortal _tiaportal;
        public ExclusiveAccess tiaMessage;

        public ActuatorsGenerator(TiaPortal tiaportal)
        {
            _tiaportal = tiaportal;
        }

        public void ActuatorsGen(MenuSelectionProvider<PlcBlockGroup>
            menuSelectionProvider)
        {
            /// 
            ///
            /// Getting PlcSoftware object from menu selection
            ///
            ///

            tiaMessage = _tiaportal.ExclusiveAccess("Odczytywanie tagów...");

            PlcSoftware plcSoftware = GetPlcSoftware(menuSelectionProvider);


            /// 
            ///
            /// Reading all tags, finding tags connected to actuators
            ///
            ///

            List<PlcTag> Tags = new List<PlcTag>(1000);
            List<PlcConstant> Constants = new List<PlcConstant>(1000);

            CheckCancellation();
            GetTagsConstantsLists(plcSoftware, ref Tags, ref Constants);


            string expression = @"^Y\d{1,3}";
            Regex regex = new Regex(expression);
            var ActuatorsConstants = Constants.Where(c => regex.IsMatch(c.Name));


            ///
            /// 
            /// Exporting tag table with actuator constants.
            /// 
            /// 

            PlcTagTable actConstants = (PlcTagTable)Constants.FirstOrDefault(c => regex.IsMatch(c.Name)).Parent;

            List<string> descriptions = new List<string>(1000);

            // File to export
            FileInfo xmlFile = new FileInfo(Path.GetTempFileName() + ".Xml");
            actConstants.Export(new FileInfo(xmlFile.FullName), ExportOptions.None);

            string xmlData = File.ReadAllText(xmlFile.FullName);
            xmlFile.Delete();

            /// 
            /// Serialize xml to class, reading actuators descriptions
            /// 

            XmlSerializer serializer = new XmlSerializer(typeof(TagTableXml.Document));
            using (StringReader reader = new StringReader(xmlData))
            {
                TagTableXml.Document tagTable = (TagTableXml.Document)serializer.Deserialize(reader);

                foreach (DocumentSWTagsPlcTagTableObjectListSWTagsPlcUserConstant constant in tagTable.SWTagsPlcTagTable.ObjectList.SWTagsPlcUserConstant)
                {
                    foreach (DocumentSWTagsPlcTagTableObjectListSWTagsPlcUserConstantObjectListMultilingualTextMultilingualTextItem multitext in constant.ObjectList.MultilingualText.ObjectList)
                    {
                        if (multitext.AttributeList.Text != "") descriptions.Add(multitext.AttributeList.Text);
                    }
                }
            }


            /// 
            /// 
            /// Save only tags with Y* text
            ///
            ///

            expression = @"\w*Y\d{1,3}\w*";
            regex = new Regex(expression);

            Dictionary<int, Actuator> Actuators = new Dictionary<int, Actuator>(1000);

            List<PlcTag> ActuatorsTags = Tags.Where(t => regex.IsMatch(t.Name)).ToList();



            /// 
            /// 
            /// Assign tags to actuators IO
            ///
            ///

            tiaMessage.Text = $"Tworzenie listy siłowników: ";
            foreach (PlcConstant c in ActuatorsConstants)
            {
                CheckCancellation();
                //Console.WriteLine(c.Name);
                Actuator actuator = new Actuator();

                actuator.Name = c.Name;
                try
                {
                    actuator.Description = descriptions.First(d => d.Contains(c.Name));
                }
                catch (Exception ex)
                {
                    actuator.Description = c.Name;
                }
                actuator.Constant = int.Parse(c.Value);

                List<int> numbersInName = FindNumbersInString(c.Name);

                if (numbersInName.Count > 0) actuator.Number = numbersInName[0];
                else continue;

                //expression = $"^I_ActY\d{1,3}Ret";

                int st = 0;
                int.TryParse(c.Name.Substring(1, 1), out st);
                actuator.Station = st >= 1 ? (EnumStations)st - 1 : 0;
                Actuators.Add(actuator.Number, actuator);
                tiaMessage.Text = $"Tworzenie listy siłowników: " + Actuators.Count;
            }

            AssingTagsToActuators(Actuators, ActuatorsTags);


            /// 
            /// 
            /// Make folders and FC_Actuators file
            ///
            ///

            // Check if !!!Devices folder exists
            PlcBlockGroup devicesGroup;
            devicesGroup = GetGroupByGroupName(plcSoftware.BlockGroup, "!!!Devices");
            if (devicesGroup == null)
            {
                devicesGroup = plcSoftware.BlockGroup.Groups.Create("!!!Devices");
            }

            // Check if actuators folder exists
            PlcBlockGroup actuatorsGroup;
            actuatorsGroup = GetGroupByGroupName(devicesGroup, "Actuators");
            if (actuatorsGroup == null)
            {
                actuatorsGroup = devicesGroup.Groups.Create("Actuators");
            }

            string nameBase = "FC_Actuators";
            string actuatorsBlockName = nameBase;
            int nameIter = 0;

            while (actuatorsGroup.Blocks.Any(b => b.Name == actuatorsBlockName))
            {
                actuatorsBlockName = nameBase + ++nameIter;
            }

            CheckCancellation();


            /// 
            /// 
            /// Generating XML file
            ///
            ///
            #region generating xml

            // File to export
            xmlFile = new FileInfo(Path.GetTempFileName() + ".Xml");
            //string xmlFilePath = Environment.CurrentDirectory + "Actuators.Xml";
            string xmlContant = XmlHelper.ActuatorsHeader.Contant;

            xmlContant = XmlHelper.InsertName(xmlContant, actuatorsBlockName);
            string tempConatant = string.Empty;

            // Xml header elements - 11
            int id = 12;
            int actuatorsNetworks = Actuators.Count * 3 + 4;
            int networkCount = 4;

            // ADD NETWORKS
            // Adding actuators

            foreach (KeyValuePair<int, Actuator> act in Actuators)
            {
                CheckCancellation();
                tiaMessage.Text = $"Generowanie networków {++networkCount} / {actuatorsNetworks}";

                tempConatant = XmlHelper.ActuatorsMovement.Contant;

                tempConatant = XmlHelper.InsertActuator(tempConatant, act.Value, ref id);

                xmlContant += tempConatant;
            }


            // Adding comment subnet
            Comment parametersComment = new Comment("--------------------Parameters--------------------");
            tempConatant = XmlHelper.SubnetComment.Contant;
            tempConatant = XmlHelper.InsertComment(tempConatant, parametersComment);
            tempConatant = XmlHelper.InsertIds(tempConatant, ref id);
            xmlContant += tempConatant;


            // Adding safety network
            tempConatant = XmlHelper.ActuatorsSafety.Contant;
            tempConatant = XmlHelper.InsertIds(tempConatant, ref id);
            xmlContant += tempConatant;


            // Adding parameters networks
            foreach (KeyValuePair<int, Actuator> act in Actuators)
            {
                CheckCancellation();
                tiaMessage.Text = $"Generowanie networków {++networkCount} / {actuatorsNetworks}";

                tempConatant = XmlHelper.ActuatorsParameters.Contant;

                tempConatant = XmlHelper.InsertActuator(tempConatant, act.Value, ref id);

                xmlContant += tempConatant;
            }


            // Adding handling network
            tempConatant = XmlHelper.ActuatorsHandling.Contant;
            tempConatant = XmlHelper.InsertIds(tempConatant, ref id);
            xmlContant += tempConatant;


            // Adding comment subnet
            Comment OutputsComment = new Comment("--------------------Outputs--------------------");
            tempConatant = XmlHelper.SubnetComment.Contant;
            tempConatant = XmlHelper.InsertComment(tempConatant, OutputsComment);
            tempConatant = XmlHelper.InsertIds(tempConatant, ref id);
            xmlContant += tempConatant;


            // Adding outputs network
            foreach (KeyValuePair<int, Actuator> act in Actuators)
            {
                CheckCancellation();
                tiaMessage.Text = $"Generowanie networków {++networkCount} / {actuatorsNetworks}";
                // Outputs template
                tempConatant = XmlHelper.ActuatorsOutputs.Contant;

                tempConatant = XmlHelper.InsertActuator(tempConatant, act.Value, ref id);

                xmlContant += tempConatant;
            }


            // Adding footer
            xmlContant += XmlHelper.ActuatorsFooter.Contant;

            #endregion

            File.WriteAllText(xmlFile.FullName, xmlContant);

            CheckCancellation();



            /// 
            /// 
            /// Import XML file
            ///
            ///

            // Import generated block
            SWImportOptions importOptions = SWImportOptions.None;

            actuatorsGroup.Blocks.Import(new FileInfo(xmlFile.FullName), ImportOptions.Override, importOptions);


            _tiaportal.Dispose();

            //return true;
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
