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

            context.Patient.BeginModifications();

            Structure ptvmax;
            StructureSet ss = context.StructureSet;
            StructureCodeDictionary scd = context.StructureCodes.VmsStructCode;

            //Plannames
            //IEnumerable<ExternalPlanSetup> ps = context.Course.ExternalPlanSetups;
            
            //List<Structure> l = GetPTVsFromPlanname(ps, ss);
            //foreach (Structure il in l)
            //{
            //    MessageBox.Show("Struktur mit Id: " + il.Id.ToString() + " in Liste gefunden.");
            //}

            //Create loopable list of PTVs
            //IEnumerable<Structure> ptvs = ss.Structures.Where(x => x.Id.StartsWith("PTV")).OrderBy(y => y.Volume).ToList();
            IEnumerable<Structure> ptvs = GetPTVsFromPlanname(context.Course.ExternalPlanSetups, ss);

            // Create merged PTV and/or find largest PTV
            if (ptvs.Count() > 1)
            {
                ptvmax = CreateMaxMergedPTV(ptvs, ss);
                ptvs = ptvs.Concat(ss.Structures.Where(x => x.Id.StartsWith(ptvmax.Id)).ToList());
            }
            else
            {
                ptvmax = ptvs.FirstOrDefault();
            }

            //Create Cropped PTVs if there are more than just one
            if (ptvs.Count() > 1)
            {
                CreateCroppedPTVs(ptvs, ss, scd, 0.3);
            }

            //Create Rings around PTVs
            CreateRings(ptvs, ss, scd, 0.3);

            //Create loopable list of OARs
            IEnumerable<Structure> oars = ss.Structures.Where(x => x.Id.StartsWith("OAR")).ToList();
            //Create OAR optimization help structures (3mm cropped)
            CreateCroppedOARs(oars, ss, ptvmax, scd, 0.3);
            
            
            //Create loopable list of PRVs
            IEnumerable<Structure> prvs = ss.Structures.Where(x => x.Id.StartsWith("PRV")).ToList();
            // warn user if one or more PRVs overlap with PTV, he may need to crop
            WarnOnPRVs(prvs, ss, ptvmax);

    }


        ///<summary>Combines all PTV in the given list to create a merged PTV</summary>
        ///<param name="ptvs">IEnumerable containing structures that should get a ring</param>
        ///<param name="ss">Structure set to operate on</param>
        ///<returns>Merged PTV or largest one.</returns>
        static Structure CreateMaxMergedPTV(IEnumerable<Structure> ptvs, StructureSet ss)
        {
            Structure ptvges;
            Structure ptvmax;
            Structure tmp;

            //create or use tmp Structure 
            try { tmp = ss.AddStructure("CONTROL", "tmp"); }
            catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }

            //create merged PTV 
            try
            {
                ptvges = ss.AddStructure("PTV", "z_PTV_ges");
            }
            catch
            {
                ptvges = ss.Structures.Single(x => x.Id == "z_PTV_ges");
            }


            ptvmax = ptvs.FirstOrDefault(); //start with the first PTV in List

            if (ptvges.IsEmpty) //if there is a user generated merged PTV do not change that
            {
                foreach (Structure tptv in ptvs)
                {
                    //search for biggest Volume (should be Last Element in List)
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
                    //ptvs = ptvs.Concat(ss.Structures.Where(x => x.Id.StartsWith("z_PTV_ges")).ToList());
                }
                else { ss.RemoveStructure(ptvges); } //delete if merged ptvges is not bigger than the largest ptv
            }
            else if (!ptvges.IsEmpty) // if there is a user generated z_PTVges use that one
            {
                ptvmax = ptvges;
            }

            ss.RemoveStructure(tmp);
            //Test on more than one PTV is now done before method call
            //else // if there is only one PTV and no user provided z_PTVges
            //{
            //    ss.RemoveStructure(ptvges);
            //}

            return ptvmax; //which is either a merged PTV or a user generated z_PTV_ges
        }

        /// <summary>
        /// Create Rings around Structures (PTVs)
        /// </summary>
        /// <param name="ptvs">IEnumerable containing structures that should get a ring</param>
        /// <param name="ss">Structure set to operate on</param>
        /// <param name="scd">Structure Code Dictionary to create Structure Code from</param>
        static void CreateRings(IEnumerable<Structure> ptvs, StructureSet ss, StructureCodeDictionary scd, double cropmm)
        {
            Structure tmpring;
            Color ringColor = Color.FromArgb(255, 255, 165, 0);

            //Regex that matches on PTV number and removes trailing date etc, also matches z_PTV_ges
            Regex ptvreg1 = new Regex(@"PTV_(\d?[A-z]+)", RegexOptions.Compiled);
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
                    tmpring.SegmentVolume = tmpring.Sub(tptv.Margin(cropmm));
                    tmpring.Color = ringColor;
                    tmpring.StructureCode = scd["Ring"];
                }
            }
        }
    
        /// <summary>
        /// Creates a cropped Structure for each PTV that has child. The resulting structure is a ringlike structure.
        /// </summary>
        /// <param name="ptvs"></param>
        /// <param name="ss"></param>
        /// <param name="scd"></param>
        /// <param name="cropmm"></param>
        static void CreateCroppedPTVs(IEnumerable<Structure> ptvs, StructureSet ss, StructureCodeDictionary scd, double cropmm)
        {
            Structure zptv;

            //Regex that matches on PTV number and removes trailing date etc
            Regex ptvreg2 = new Regex(@"^PTV_?(\d[A-Z]+)", RegexOptions.Compiled);

            //Create optimization structures from PTVs
            foreach (Structure tptv in ptvs)
            {
                //build potential parent PTV name (a.e. PTV_1A is parent for PTV_1BA)
                Match ptvmatch = ptvreg2.Match(tptv.Id);

                if (ptvmatch.Success)
                {
                    string tmpname = ptvmatch.Groups[1].Value;
                    tmpname = tmpname.Remove(1, 1); //remove first character after number
                    
                    //ignore PTV_1
                    if (tmpname.Length == 1)
                    { 
                        continue;
                    }
                    
                    tmpname = "PTV_" + tmpname;

                    //test if there is a parent PTV
                    //Structure parentPtv = ss.Structures.FirstOrDefault(x => x.Id.Equals(tmpname));
                    Structure parentPtv = ss.Structures.FirstOrDefault(x => Regex.IsMatch(x.Id, @tmpname));
                    if (parentPtv != null)
                    {
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

        static void CreateCroppedOARs(IEnumerable<Structure> oars, StructureSet ss, Structure ptvcrop, StructureCodeDictionary scd, double cropmm)
        {
            Structure tmpoar;
            Structure tmp;
            try { tmp = ss.AddStructure("CONTROL", "tmp"); }
            catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }

            foreach (Structure str in oars)
            {
                //Test if volumes overlap with or are within 3mm of PTV
                tmp.SegmentVolume = ptvcrop.And(str.Margin(cropmm));
                if (tmp.Volume != 0.0 && !(str.Id.Contains("Spinal") || str.Id.Contains("HS") || str.Id.Contains("Opt") || str.Id.Contains("Chia"))) //nerves do not get cropped!
                {
                    //MessageBox.Show("Struktur " + str.Id + " überlappt");
                    try
                    {
                        tmpoar = ss.AddStructure("CONTROL", "z_" + str.Id.Substring(4));
                        tmpoar.Color = str.Color;
                        tmpoar.StructureCode = scd["Control"];
                    }
                    catch
                    {
                        tmpoar = ss.Structures.FirstOrDefault(x => x.Id == "z_" + str.Id.Substring(4));
                    }
                    if (tmpoar.IsEmpty)
                    {
                        tmpoar.SegmentVolume = str.Sub(ptvcrop.Margin(cropmm));
                        if (tmpoar.IsEmpty) { ss.RemoveStructure(tmpoar); } //if help structure is empty we can remove it
                    }
                }
            }
            ss.RemoveStructure(tmp);
        }


        static void WarnOnPRVs(IEnumerable<Structure> prvs, StructureSet ss, Structure ptvmax)
        {
            string message = "";
            int scount = 0;
            Structure tmp;
            try { tmp = ss.AddStructure("CONTROL", "tmp"); }
            catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }

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

        static List<Structure> GetPTVsFromPlanname(IEnumerable<ExternalPlanSetup> ps, StructureSet ss)
        {
            List<Structure> ptvs = new List<Structure>();
            List<string> ids = new List<string>();
            string pattern = @"\d[A-Z]+";
            foreach (ExternalPlanSetup eps in ps)
            {
                foreach(Match m in Regex.Matches(eps.Id, pattern))
                {
                    ids.Add(m.Value.ToString());
                }    
            }
            foreach(string id in ids)
            {
                if (ss.Structures.Any(x => x.Id.StartsWith("PTV_" + id)))
                {
                    Structure smatch = ss.Structures.Single(x => x.Id.StartsWith("PTV_" + id));
                    if (!ptvs.Contains(smatch))
                    {
                        ptvs.Add(smatch);
                    }
                }
                else
                {
                    MessageBox.Show("Keine Struktur mit Id PTV_" + id + " gefunden!");
                }

            }
            return ptvs;
        }
    }
}
