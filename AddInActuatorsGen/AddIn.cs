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
using AddInActuatorsGen.Functions;

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
                ActuatorsGenerator actuatorsGenerator = new ActuatorsGenerator(_tiaportal);
                actuatorsGenerator.ActuatorsGen(menuSelectionProvider);
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



        
    }
}