using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AddInActuatorsGen.Models;

namespace AddInActuatorsGen.Helpers
{
    public class XmlHelper
    {

        public static Template ActuatorsHeader = new Template(EnumTemplates.ActuatorsHeader);
        public static Template ActuatorsFooter = new Template(EnumTemplates.ActuatorsFooter);
        public static Template SubnetComment = new Template(EnumTemplates.Comment);
        public static Template ActuatorsMovement = new Template(EnumTemplates.ActuatorsMovement);
        public static Template ActuatorsSafety = new Template(EnumTemplates.ActuatorsSafety);
        public static Template ActuatorsParameters = new Template(EnumTemplates.ActuatorsParameters);
        public static Template ActuatorsHandling = new Template(EnumTemplates.ActuatorsHandling);
        public static Template ActuatorsOutputs = new Template(EnumTemplates.ActuatorsOutputs);


        public static Template TextlistHeader = new Template(EnumTemplates.TextlistHeader);
        public static Template TextlistComment = new Template(EnumTemplates.TextlistComment);
        public static Template TextlistCommentMulti = new Template(EnumTemplates.TextlistCommentMulti);
        public static Template tplTextlistEntry = new Template(EnumTemplates.TextlistEntry);
        public static Template TextlistEntryMulti = new Template(EnumTemplates.TextlistEntryMulti);
        public static Template TextlistFooter = new Template(EnumTemplates.TextlistFooter);


        public static string InsertIds(string xmlContant, ref int id)
        {
            int _id = id;
            while (xmlContant.Contains("{id}"))
            {
                xmlContant = Regex.Replace(xmlContant, @"\{id\}", match => _id++.ToString());
            }
            id = _id;

            return xmlContant;
        }


        public static string InsertName(string xmlContant, string name)
        {
            return xmlContant.Replace("{name}", name);
        }


        public static string InsertNumber(string  xmlContant, int number) 
        { 
            return xmlContant.Replace("{number}", number.ToString());
        }


        public static string InsertNumber(string xmlContant, string number)
        {
            return xmlContant.Replace("{number}", number);
        }


        public static string InsertText(string xmlContant, string text)
        {
            return xmlContant.Replace("{text}", text);
        }


        public static string InsertLang(string xmlContant, string lang)
        {
            return xmlContant.Replace("{lang}", lang);
        }


        public static string InsertActuator(string xmlContant, Actuator actuator, ref int id)
        {
            xmlContant = InsertName(xmlContant, actuator.Name);
            xmlContant = InsertNumber(xmlContant, actuator.Number);
            xmlContant = xmlContant.Replace("{constant}", actuator.Constant.ToString());
            xmlContant = xmlContant.Replace("{station}", actuator.Station.ToString());
            xmlContant = xmlContant.Replace("{station_name}", GetTypeDescription(actuator.Station));
            xmlContant = xmlContant.Replace("{input_retract}", actuator.InputRetract);
            xmlContant = xmlContant.Replace("{input_extend}", actuator.InputExtend);
            xmlContant = xmlContant.Replace("{output_retract}", actuator.OutputRetract);
            xmlContant = xmlContant.Replace("{output_extend}", actuator.OutputExtend);
            xmlContant = InsertIds(xmlContant, ref id);
            return xmlContant;
        }


        public static string InsertTextlistCommentMulti(List<CultureInfo> cultures, List<string> comments) 
        {
            string result = "";
            foreach (CultureInfo culture in cultures)
            {
                string tempContant = TextlistCommentMulti.Contant;
                tempContant = InsertLang(tempContant, culture.Name);
                tempContant = InsertText(tempContant, comments[cultures.IndexOf(culture)]);
                result += tempContant;
            }

            return result;
        }


        public static string AddTextlistComment(List<CultureInfo> cultures, List<string> texts, ref int id)
        {
            string tempContant = TextlistComment.Contant;
            tempContant = tempContant.Replace("{multilingual}", 
                InsertTextlistCommentMulti(cultures, texts));
            tempContant = InsertIds(tempContant, ref id);
            return tempContant;
        }


        public static string InsertTextlistEntryMulti(List<CultureInfo> cultures, List<string> texts)
        {
            string result = "";
            foreach (CultureInfo culture in cultures)
            {
                string tempContant = TextlistEntryMulti.Contant;
                tempContant = InsertLang(tempContant, culture.Name);
                tempContant = InsertText(tempContant, texts[cultures.IndexOf(culture)]);
                result += tempContant;
            }

            return result;
        }


        public static string AddTextlistEntry(TextlistEntry entry, List<CultureInfo> cultures, ref int id)
        {
            string tempContant = tplTextlistEntry.Contant;
            tempContant = InsertNumber(tempContant, entry.Number);
            tempContant = tempContant.Replace("{multilingual}",
                InsertTextlistEntryMulti(cultures, entry.Texts));
            tempContant = InsertIds(tempContant, ref id);
            return tempContant;
        }


        public static string InsertComment(string xmlContant, Comment comment)
        {
            while (xmlContant.Contains("{comment}"))
            {
                xmlContant = Regex.Replace(xmlContant, @"\{comment\}", match => comment.CommentText);
            }
            return xmlContant;
        }


        public static string ProcessString(string input)
        {
            StringBuilder output = new StringBuilder();

            for (int i = 0; i < input.Length; i++)
            {
                char currentChar = input[i];
                char nextChar = i + 1 < input.Length ? input[i+1] : '\0';
                char prevChar = i - 1 >= 0 ? input[i-1] : '\0';

                if (char.IsUpper(currentChar) && i > 0 && !char.IsUpper(nextChar) && nextChar != ' ' && !char.IsUpper(prevChar) && prevChar != ' ' && prevChar != '_')
                {
                    output.Append(' ');
                    output.Append(char.ToLower(currentChar));
                }
                else if (currentChar == '_')
                {
                    output.Append(" - ");
                }
                else
                {
                    output.Append(currentChar);
                }
            }

            return output.ToString();
        }


        /// <summary>
        /// Get description of defined program element
        /// </summary>
        /// <param name="defGroup">Group from DefGroup enum</param>
        /// <returns>String with group name</returns>
        private static string GetTypeDescription(EnumStations defGroup)
        {
            Type enumType = typeof(EnumStations);

            var groupInfo = enumType.GetField(defGroup.ToString());
            var attribute = groupInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
            var description = (DescriptionAttribute)attribute[0];

            return description.Description;
        }
    }
}
