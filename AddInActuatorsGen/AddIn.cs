using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.HW;
using Siemens.Engineering.SW;
using System.Windows.Forms;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System;
using System.Text.RegularExpressions;
using static TiaHelperLibrary.TiaHelper;
using TiaHelperLibrary.Models.Tia;
using AddinActuatorsGen.Helpers;
using TiaXmlGenerator;
using TiaXmlGenerator.Models;
using System.Xml.Linq;
using TiaHelperLibrary.Models.Tia;
using System.Xml.Serialization;
using Siemens.Engineering.Hmi.Tag;
using static TiaHelperLibrary.Models.Tia.TagTableXml;

namespace AddInActuatorsGen
{
    public class AddIn : ContextMenuAddIn
    {
        Siemens.Engineering.
        /// <summary>
        ///The global TIA Portal Object 
        ///<para>It will be used in the TIA Add-In.</para>
        /// </summary>
        TiaPortal _tiaportal;
        public ExclusiveAccess tiaMessage;

        /// <summary>
        /// The display name of the Add-In.
        /// </summary>
        private const string s_DisplayNameOfAddIn = "Generatory";

        /// <summary>
        /// The constructor of the AddIn.
        /// Creates an object of the class AddIn
        /// Called from AddInProvider, when the first
        /// right-click is performed in TIA
        /// Motherclass' constructor of ContextMenuAddin
        /// will be executed, too. 
        /// </summary>
        /// <param name="tiaportal">
        /// Represents the actual used TIA Portal process.
        /// </param>
        public AddIn(TiaPortal tiaportal) : base(s_DisplayNameOfAddIn)
        {
            /*
             * The acutal TIA Portal process is saved in the 
             * global TIA Portal variable _tiaportal
             * tiaportal comes as input Parameter from the
             * AddInProvider
            */
            _tiaportal = tiaportal;
        }

        /// <summary>
        /// The method is supplemented to include the Add-In
        /// in the Context Menu of TIA Portal.
        /// Called when a right-click is performed in TIA
        /// and a mouse-over is performed on the name of the Add-In.
        /// </summary>
        /// <typeparam name="addInRootSubmenu">
        /// The Add-In will be displayed in 
        /// the Context Menu of TIA Portal.
        /// </typeparam>
        /// <example>
        /// ActionItems like Buttons/Checkboxes/Radiobuttons
        /// are possible. In this example, only Buttons will be created 
        /// which will start the Add-In program code.
        /// </example>
        protected override void BuildContextMenuItems(ContextMenuAddInRoot
            addInRootSubmenu)
        {
            /* Method addInRootSubmenu.Items.AddActionItem
             * Will Create a Pushbutton with the text 'Start Add-In Code'
             * 1st input parameter of AddActionItem is the text of the 
             *          button
             * 2nd input parameter of AddActionItem is the clickDelegate, 
             *          which will be executed in case the button 'Start 
             *          Add-In Code' will be clicked/pressed.
             * 3rd input parameter of AddActionItem is the 
             *          updateStatusDelegate, which will be executed in 
             *          case there is a mouseover the button 'Start 
             *          Add-In Code'.
             * in <placeholder> the type of AddActionItem will be 
             *          specified, because AddActionItem is generic 
             * AddActionItem<DeviceItem> will create a button that will be 
             *          displayed if a rightclick on a DeviceItem will be 
             *          performed in TIA Portal
             * AddActionItem<Project> will create a button that will be 
             *          displayed if a rightclick on the project name 
             *          will be performed in TIA Portal
            */
            addInRootSubmenu.Items.AddActionItem<PlcBlockGroup>(
                "FC_Actuators", OnDoSomething, OnCanSomething);

            addInRootSubmenu.Items.AddActionItem<PlcBlock>(
                "FC_Actuators - Folder", OnDoSomething, OnCanSomething);

            addInRootSubmenu.Items.AddActionItem<Project>(
                "FC_Actuators - Folder", OnDoSomething, OnCanSomething);

            addInRootSubmenu.Items.AddActionItem<DeviceItem>(
                "FC_Actuators - Folder", OnDoSomething, OnCanSomething);
        }

        /// <summary>
        /// The method contains the program code of the TIA Add-In.
        /// Called when the button 'Start Add-In Code' will be pressed.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <DeviceItem>
        /// </typeparam>
        private void OnDoSomething(MenuSelectionProvider<PlcBlockGroup>
            menuSelectionProvider)
        {
            try
            {
                //_tiaportal.Projects.First().Save();
                ActuatorsGenerator(menuSelectionProvider);
            }
            catch (Exception ex)
            {
                if (ex.Message == "Addin canceled") _tiaportal.Dispose();
                else
                {
                    LogToFile.Save(ex);
                    _tiaportal.Dispose();
                }
            }
        }

        private void OnDoSomething(MenuSelectionProvider<PlcBlock>
            menuSelectionProvider)
        {
        }

        private void OnDoSomething(MenuSelectionProvider<Project>
            menuSelectionProvider)
        {
        }

        private void OnDoSomething(MenuSelectionProvider<DeviceItem>
            menuSelectionProvider)
        {
        }


        /// <summary>
        /// Called when there is a mousover the button at a DeviceItem.
        /// It will be used to enable the button.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <DeviceItem>
        /// </typeparam>
        private MenuStatus OnCanSomething(MenuSelectionProvider
            <PlcBlockGroup> menuSelectionProvider)
        {
            //enable the button
            return MenuStatus.Enabled;
        }

        private MenuStatus OnCanSomething(MenuSelectionProvider
            <PlcBlock> menuSelectionProvider)
        {
            //enable the button
            return MenuStatus.Disabled;
        }

        private MenuStatus OnCanSomething(MenuSelectionProvider
            <Project> menuSelectionProvider)
        {
            //enable the button
            return MenuStatus.Disabled;
        }

        private MenuStatus OnCanSomething(MenuSelectionProvider
            <DeviceItem> menuSelectionProvider)
        {
            //enable the button
            return MenuStatus.Disabled;
        }

        /// <summary>
        /// The method contains the program code of the TIA Add-In.
        /// Called when the button will be pressed on project level.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <Project>
        /// </typeparam>
        private void OnClickProject(MenuSelectionProvider<Project>
            menuSelectionProvider)
        {
            //Do Nothing on Project level
        }

        /// <summary>
        /// Called when there is a mousover the button at the Project 
        /// Level. It will be used to disable the button because no 
        /// action should be performed on project level.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <Project>
        /// </typeparam>
        private MenuStatus OnStatusUpdateProject(MenuSelectionProvider
            <Project> menuSelectionProvider)
        {
            //disable the button
            return MenuStatus.Disabled;
        }


        /// <summary>
        /// This method will be invoked by the TIA Add-In Tester. The return value of this
        /// method will be provided in the Click- and UpdateDelegate's MenuSelectionProvider
        /// </summary>
        /// <param name="label">Label of the context menu item</param>
        /// <returns>Objects to provide for the MenuSelectionProvider</returns>
        public IEnumerable<IEngineeringObject> GetSelection(string label)
        {
            // go to project settings -> Debug -> command line arguments
            // specify which context menu item to test at --item
            PlcSoftware software = GetPlcSoftware(_tiaportal);

            switch (label)
            {
                case "FC_Actuators":
                    /*var myPlc = project
                         .Devices[1]
                        .DeviceItems.FirstOrDefault(plc => plc.Name.Length > 0)
                        .GetService<SoftwareContainer>().Software as PlcSoftware;
                    // return the program blocks folder*/
                    return software.BlockGroup.Groups;

                default:
                    return Enumerable.Empty<IEngineeringObject>();
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



        public void ActuatorsGenerator(MenuSelectionProvider<PlcBlockGroup>
            menuSelectionProvider)
        {
            /// 
            ///
            /// Getting PlcSoftware object from menu selection
            ///
            ///

            tiaMessage = _tiaportal.ExclusiveAccess("Odczytywanie tagów...");

            IEnumerable<PlcBlockGroup> selection = menuSelectionProvider.GetSelection<PlcBlockGroup>();

            var item = selection.FirstOrDefault();
            var parent = item.Parent;
            int i = 0;
            while (!(parent is PlcSoftware))
            {
                if (i > 10) throw new Exception("Can't find PLC software");
                parent = parent.Parent;
                i++;
            }

            PlcSoftware plcSoftware = (PlcSoftware)parent;


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
    }
}