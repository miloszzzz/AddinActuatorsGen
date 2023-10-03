using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddInActuatorsGen.Models
{
    /// <summary>
    /// 
    /// </summary>
    public class Template
    {
        public string _contant { get; set; }
        private EnumTemplates _templateType;
        public string Contant = "";


        public Template(EnumTemplates enumTemplate)
        { 
            _templateType = enumTemplate; 
            ReadTemplate();
        }


        /// <summary>
        /// Reading file with template text
        /// </summary>
        public void ReadTemplate()
        {
            switch (_templateType)
            {
                case EnumTemplates.Comment:
                    _contant = AddInActuatorsGen.Properties.Resources.EmptySubnetComment;
                    break;

                case EnumTemplates.ActuatorsHeader:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsHeader;
                    break;

                case EnumTemplates.ActuatorsFooter:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsFooter;
                    break;

                case EnumTemplates.ActuatorsMovement:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsMovement;
                    break;

                case EnumTemplates.ActuatorsSafety:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsSafety;
                    break;

                case EnumTemplates.ActuatorsParameters:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsParameters;
                    break;

                case EnumTemplates.ActuatorsHandling:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsHandling;
                    break;

                case EnumTemplates.ActuatorsOutputs:
                    _contant = AddInActuatorsGen.Properties.Resources.FC_ActuatorsOutputs;
                    break;

                case EnumTemplates.TextlistHeader:
                    _contant = AddInActuatorsGen.Properties.Resources.Textlist_Header;
                    break;

                case EnumTemplates.TextlistComment:
                    _contant = AddInActuatorsGen.Properties.Resources.Textlist_Comment;
                    break;

                case EnumTemplates.TextlistCommentMulti:
                    _contant = AddInActuatorsGen.Properties.Resources.Textlist_CommentMulitlingual;
                    break;

                case EnumTemplates.TextlistEntry:
                    _contant = AddInActuatorsGen.Properties.Resources.Textlist_Entry;
                    break;

                case EnumTemplates.TextlistEntryMulti:
                    _contant = AddInActuatorsGen.Properties.Resources.Textlist_EntryMultilingual;
                    break;

                case EnumTemplates.TextlistFooter:
                    _contant = AddInActuatorsGen.Properties.Resources.Textlist_Footer;
                    break;

                default: 
                    _contant = string.Empty;
                    break;
            }
            Contant = _contant;
        }
    }


    
}
