using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Drawing;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Text.RegularExpressions;

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
  public class Script
  {

    const string SCRIPT_NAME = "AutoStructures Script";

    public Script()
    {
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
    {
            // TODO : Add here the code that is called when the script is launched from Eclipse.
            if (context.Patient == null || context.StructureSet == null)
            {
                MessageBox.Show("Please load a patient, 3D image, and structure set before running this script.", SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            
            StructureSet ss = context.StructureSet;

            context.Patient.BeginModifications();

                        //Structure zPTV = ss.AddStructure("CONTROL", "z_PTV1A");
            Structure tmpring;
            Structure zptv;
            Structure tmp2;
            Structure ptvges;
            Structure ptvmax;
            Structure tmp = ss.AddStructure("CONTROL", "tmp");
            StructureCodeDictionary scd = context.StructureCodes.VmsStructCode;
            Color ringColor = Color.FromArgb(255, 255, 165, 0);

            //Create loopable list of PTVs
            IEnumerable<Structure> ptvs = ss.Structures.Where(x => x.Id.StartsWith("PTV")).OrderBy(y => y.Volume).ToList();

            //create merged PTV 
            try
            {
                ptvges = ss.AddStructure("PTV", "z_PTVges");
            }
            catch
            {
                ptvges = ss.Structures.Single(x => x.Id == "z_PTVges");
            }

            ptvmax = ptvs.FirstOrDefault(); //start with the first PTV in List
            if ( ptvs.Count() > 1 && ptvges.IsEmpty) 
            {
                foreach (Structure tptv in ptvs)
                {
                    //MessageBox.Show("PTVmax ID: " + ptvmax.Id + " (" + ptvmax.Volume + ")" + "\ntPTV Id: " + tptv.Id + " (" + tptv.Volume + ")");
                    if (tptv.Volume > ptvmax.Volume)
                    {
                        ptvmax = tptv;
                    }
                    tmp.SegmentVolume = ptvges.Or(tptv);
                    if (ptvges.Volume < tmp.Volume)
                    {
                        ptvges.SegmentVolume = tmp.SegmentVolume;
                    }
                }
                if (ptvges.Volume > ptvmax.Volume)
                {
                    ptvmax = ptvges;
                    ptvs = ptvs.Concat(ss.Structures.Where(x => x.Id.StartsWith("z_PTVges")).ToList());
                }
                else { ss.RemoveStructure(ptvges); } //delete if ptvges is not bigger than the largest ptv
            }
            else if (!ptvges.IsEmpty) // if there is a user generated z_PTVges use that one
            {
                ptvmax = ptvges;
                ptvs = ptvs.Concat(ss.Structures.Where(x => x.Id.StartsWith("z_PTVges")).ToList());
            }
            else // if there is only one PTV and no user provided z_PTVges
            {
                ss.RemoveStructure(ptvges);
            }


            //Regex that matches on PTV number and removes trailing date etc, also matches z_PTVges
            Regex ptvreg1 = new Regex(@"PTV_?(\d?[A-z]+)", RegexOptions.Compiled);

            //Create ring structures from PTVs
            foreach (Structure tptv in ptvs)
            {
                Match ptvmatch = ptvreg1.Match(tptv.Id);
                string ptvid = ptvmatch.Groups[1].Value;
                if (ptvid == "")
                {
                    MessageBox.Show("Falsch benanntes PTV gefunden!");
                    break;
                }
                try
                {
                    tmpring = ss.AddStructure("CONTROL", "z_Ring_" + ptvid); //does not exist yet
                }
                catch
                {
                    tmpring = ss.Structures.FirstOrDefault(x => x.Id == "z_Ring_" + ptvid); //already exists
                }
                if (tmpring.IsEmpty)
                {
                    tmpring.SegmentVolume = tmpring.Or(tptv.Margin(20.0));
                    tmpring.SegmentVolume = tmpring.Sub(tptv.Margin(3.0));
                    tmpring.Color = ringColor;
                    tmpring.StructureCode = scd["Ring"];
                }
            }


            //Regex that matches on PTV number and removes trailing date etc
            Regex ptvreg2 = new Regex(@"^PTV_?(\d[A-Z]+)", RegexOptions.Compiled);

            //Create optimization structures from PTVs
            if (ptvs.Count() > 1)
            {
                foreach (Structure tptv in ptvs)
                {
                    //build potential parent PTV name (a.e. PTV_1A is parent for PTV_1BA)
                    Match ptvmatch = ptvreg2.Match(tptv.Id);

                    if (ptvmatch.Success)
                    {
                        string tmpname = ptvmatch.Groups[1].Value;
                        tmpname = tmpname.Remove(1, 1); //remove first character after number
                        tmpname = "PTV_" + tmpname;

                        string pattern = @tmpname + "[^A-Z]"; //ignore parents without a letter at the end

                        //test if there is a parent PTV
                        //Structure parentPtv = ss.Structures.FirstOrDefault(x => x.Id.Equals(tmpname));
                        Structure parentPtv = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, pattern));
                        if (parentPtv != null)
                        {
                            //MessageBox.Show(tptv.Id + ": Parent = " + tmpname);
                            try
                            {
                                zptv = ss.AddStructure("PTV", "z_" + tmpname);
                                zptv.StructureCode = parentPtv.StructureCode;
                            }
                            catch
                            {
                                zptv = ss.Structures.FirstOrDefault(x => x.Id == "z_" + tmpname);
                            }
                            if (zptv.IsEmpty)
                            {
                                zptv.SegmentVolume = parentPtv.Sub(tptv.Margin(3.0));
                            }
                        }
                    }
                }
            }

            //Create OAR optimization help structures (3mm cropped)
            //Create loopable list of OARs
            IEnumerable<Structure> oars = ss.Structures.Where(x => x.Id.StartsWith("OAR")).ToList();
            foreach (Structure str in oars)
            {
                //Test if volumes overlap with or are within 3mm of PTV
                tmp.SegmentVolume = ptvmax.And(str.Margin(3.0));
                if (tmp.Volume != 0.0 && !(str.Id.Contains("Spinal") || str.Id.Contains("HS") || str.Id.Contains("Opt") || str.Id.Contains("Chia"))) //nerves do not get cropped!
                {
                    //MessageBox.Show("Struktur " + str.Id + " überlappt");
                    try
                    {
                        tmp2 = ss.AddStructure("CONTROL", "z_" + str.Id.Substring(4));
                        tmp2.Color = str.Color;
                        tmp2.StructureCode = scd["Control"];
                    }
                    catch
                    {
                        tmp2 = ss.Structures.FirstOrDefault(x => x.Id == "z_" + str.Id.Substring(4));
                    }
                    if (tmp2.IsEmpty)
                    {
                        tmp2.SegmentVolume = str.Sub(ptvmax.Margin(3.0));                        
                        if (tmp2.IsEmpty) { ss.RemoveStructure(tmp2); } //if help structure is empty we can remove it
                    }
                }
            }

            // warn user if one or more PRVs overlap with PTV, he may need to crop
            int scount = 0;
            string message = "";
            //Create loopable list of PRVs
            IEnumerable<Structure> prvs = ss.Structures.Where(x => x.Id.Contains("PRV")).ToList();
            foreach (Structure str in prvs)
            {
                tmp.SegmentVolume = ptvmax.And(str);
                if (tmp.Volume != 0.0)
                {
                    if (scount > 0) { message += ", "; }
                    message += str.Id;
                    scount++;
                }
            }
            if (message != "" && scount > 1) { MessageBox.Show("Strukturen " + message + " überlappen mit einem PTV"); }
            else if (message != "" && scount == 1) { MessageBox.Show("Struktur " + message + " überlappt mit einem PTV"); }

            ss.RemoveStructure(tmp);
        }
  }
}
