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

/* The script adds a ring structure around the largest of the PTVs of the current plan or the
 * combined/merged PTVs if there is more than one PTV referenced in the plan name. The ring 
 * has 3mm spacing between the largest PTV, 5mm between the first boost PTV and 7mm between the
 * second boost PTV.
 * The script creates cropped PTVs if the plan contains an integrated boost. The space between 
 * parent PTV and boost PTV is also 3mm. 
 * Helper structures for OARs are also created by cropping them 3mm from the largest or merged PTV.
 * The Script displays a warning message if any PRV overlaps with a PTV.
 */


/* Das Skript fügt einen Ring um das größte PTV des jeweiligen Plans bzw. die Vereinigung aller
 * vorhandenen PTVs ein. Der Ring hat einen Abstand von 3 mm vom größten PTV, 5mm vom ersten
 * Boost PTV und 7 mm vom zweiten Boost PTV.
 * Es erzeugt Hilfs PTVs, wenn es sich um einen Plan mit integriertem Boost handelt. Diese haben
 * jeweils einen Abstand von 3mm zum Eltern PTV. Ein evtl. vorhandener Überlapp von PTV_1A und 
 * PTV_2A oder ähnlichem wird nicht berücksichtigt.
 * Hilfs OARs werden am größten PTV bzw. dem Summen PTV des jeweiligen Plans erstellt, indem 
 * die jeweiligen OARs um 3mm gecroppt werden.
 * Außerdem wird auf einen Überlapp von PRVs und dem jeweils größten PTV geprüft und ggf. eine
 * Warnung ausgegeben. Evtl. muss der Benutzer hier von Hand ein passendes z_PTV erzeugen.
 */


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
            if (context.ExternalPlanSetup == null)
            {
                MessageBox.Show("Please load a plan before running this script.", SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
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
            IEnumerable<Structure> ptvs = GetPTVsFromPlanname(context.ExternalPlanSetup, ss);

            // Create merged PTV and/or find largest PTV
            ptvmax = CreateMaxMergedPTV(ptvs, ss);
            if (!ptvs.Contains(ptvmax))
            {
                ptvs = ptvs.Concat(ss.Structures.Where(x => x.Id.StartsWith(ptvmax.Id)).ToList());
            }

            //Create loopable list of OARs to crop
            IEnumerable<Structure> oars = ss.Structures.Where(x => x.Id.StartsWith("OAR")).ToList();
            //Create OAR optimization help structures (3mm cropped)
            CreateCroppedOARs(oars, ss, ptvmax, scd, 3.0);

            //Create loopable list of PRVs
            IEnumerable<Structure> prvs = ss.Structures.Where(x => x.Id.StartsWith("PRV")).ToList();
            // warn user if one or more PRVs overlap with PTV, he may need to crop
            WarnOnPRVs(prvs, ss, ptvmax);

            //Create Ring around largest or compound PTV
            CreateRing(ptvs, ss, scd, 3.0);

            //Create Cropped PTVs if there are more than just one (only after cropping OARs and testing for PRVs)
            CreateCroppedPTVs(ptvs, ss, scd, 3.0);
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
        /// Creates Ring around largest PTV Structure and crops from child PTVs with additional margin.
        /// </summary>
        /// <param name="ptvs">IEnumerable containing structures that should get a ring</param>
        /// <param name="ss">Structure set to operate on</param>
        /// <param name="scd">Structure Code Dictionary to create Structure Code from</param>
        static void CreateRing(IEnumerable<Structure> ptvs, StructureSet ss, StructureCodeDictionary scd, double cropmm)
        {
            Structure tmpring;
            Color ringColor = Color.FromArgb(255, 255, 165, 0);

            //Regex that matches on PTV number and removes trailing date etc, also matches z_PTV_ges
            Regex ptvreg1 = new Regex(@"PTV_(\d?[A-z]+)", RegexOptions.Compiled);

            //Use only largest (compound) PTV to create a ring for the plan.
            Structure tptv = ptvs.OrderByDescending(x => x.Volume).FirstOrDefault();

            Match ptvmatch = ptvreg1.Match(tptv.Id);
            string ptvid = ptvmatch.Groups[1].Value;
            if (ptvid == "")
            {
                MessageBox.Show("Falsch benanntes PTV gefunden!");
                return;
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
                //Crop PTV and any child PTVs from ring with larger distance
                foreach (Structure cptv in ptvs)
                {
                    tmpring.SegmentVolume = tmpring.Sub(cptv.Margin(cropmm + 2.0 * GetParentCount(ptvs, cptv.Id)));
                }
                tmpring.Color = ringColor;
                tmpring.StructureCode = scd["Ring"];
            }

            //Codeblock might be interessting for SRS Structures or 
            //    Match ptvmatch = ptvreg1.Match(tptv.Id);
            //    string ptvid = ptvmatch.Groups[1].Value;
            //    if (ptvid == "")
            //    {
            //        MessageBox.Show("Falsch benanntes PTV gefunden!");
            //        break;
            //    }
            //    try
            //    {
            //        tmpring = ss.AddStructure("CONTROL", "z_Ring_" + ptvid); //does not exist yet
            //    }
            //    catch
            //    {
            //        tmpring = ss.Structures.FirstOrDefault(x => x.Id == "z_Ring_" + ptvid); //already exists
            //    }
            //    if (tmpring.IsEmpty)
            //    {
            //        tmpring.SegmentVolume = tmpring.Or(tptv.Margin(20.0));
            //        tmpring.SegmentVolume = tmpring.Sub(tptv.Margin(cropmm));
            //        tmpring.Color = ringColor;
            //        tmpring.StructureCode = scd["Ring"];
            //    }
            //}
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

            //Create optimization structures from PTVs
            foreach (Structure tptv in ptvs)
            {
                //get parent id
                string tmpname = GetParentId(ptvs, tptv.Id);
                //if partent was found crop it
                if (tmpname != "")
                {
                    //Structure parentPtv = ss.Structures.FirstOrDefault(x => x.Id.Equals(tmpname));
                    Structure parentPtv = ptvs.FirstOrDefault(x => Regex.IsMatch(x.Id, @tmpname));
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
                        zptv.SegmentVolume = parentPtv.Sub(tptv.Margin(cropmm));
                        if (ptvs.Any(x => x.Id.Equals("z_PTV_ges")) && Regex.IsMatch(tptv.Id, @"^PTV_\d[A-Z]{2}"))
                        {
                            Structure tges = ptvs.FirstOrDefault(x => x.Id.Equals("z_PTV_ges"));
                            tges.SegmentVolume = tges.Sub(tptv.Margin(cropmm));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Creates cropped z_OARname structures by cropping them from the given ptvcrop. 
        /// </summary>
        /// <param name="oars">IEnumerable of OARs</param>
        /// <param name="ss">structure set</param>
        /// <param name="ptvcrop">PTV to crop with</param>
        /// <param name="scd">Structure Code Dictionary for Structure Code</param>
        /// <param name="cropmm">Millimeter to crop</param>
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


        /// <summary>
        /// Returns the number of parents the given ptv has in the provided list of PTVs.
        /// </summary>
        /// <param name="ptvs">List of PTVs to search</param>
        /// <param name="ptvid">PTV Id whose number of parents are asked for</param>
        /// <returns></returns>
        static double GetParentCount(IEnumerable<Structure> ptvs, string ptvid)
        {
            //Regex ptvreg = new Regex(@"^(PTV_\d[A-Z]+)", RegexOptions.Compiled);
            //string parent = Regex.Match(ptvid, @"^(PTV_\d[A-Z]+)").Groups[0].Value;
            //parent = parent.Remove(5, 1);
            string parent = GetParentId(ptvs, ptvid);
            if (parent != "")
            {
                double count = 1.0 + GetParentCount(ptvs, parent);
                return count;
            }
            else return 0.0;
        }

        /// <summary>
        /// Returns Id of parent volume if there is one in the provided list of PTVs. If there is no parent to be found it returns an empty string.
        /// </summary>
        /// <param name="ptvs">List of PTVs to search</param>
        /// <param name="childId">Id of child</param>
        /// <returns></returns>
        static string GetParentId(IEnumerable<Structure> ptvs, string childId)
        {
            Regex ptvreg = new Regex(@"^(PTV_\d[A-Z]+)", RegexOptions.Compiled);
            Match parentMatch = ptvreg.Match(childId); // test if child id is valid ptv id
            string parentId = parentMatch.Groups[0].Value; // parentId equals childId without additional characters
            if (parentMatch.Success)
            {
                parentId = parentId.Remove(5, 1);
                if (ptvs.Any(x => x.Id.Contains(parentId)) && Regex.IsMatch(parentId, @"^PTV_\d[A-Z]+"))
                {
                    return parentId;
                }
                //else search recursive for grandparentId etc.
                else
                {
                    return GetParentId(ptvs, parentId); 
                }
            }
            else { return ""; }
        }

        /// <summary>
        /// Prints a warning message if one or more PRVs overlap with PTVs.
        /// </summary>
        /// <param name="prvs">List of PRV Structures</param>
        /// <param name="ss">Used Structureset</param>
        /// <param name="ptvmax">Largest PTV or compound PTV from Plan</param>
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

        /// <summary>
        /// Returns List of PTVs according to Planname found in structure Set
        /// </summary>
        /// <param name="ps">PlanSetup from context</param>
        /// <param name="ss">StructureSet from context</param>
        /// <returns></returns>
        static List<Structure> GetPTVsFromPlanname(ExternalPlanSetup ps, StructureSet ss)
        {
            List<Structure> ptvs = new List<Structure>();
            List<string> ids = new List<string>();

            // if plan id containts "-" (a.e. c1A-1DCBA)
            Regex psIdReg = new Regex(@"\d[A-Z]+?-(\d[A-Z]{3,})", RegexOptions.Compiled);
            Match psIdMatch = psIdReg.Match(ps.Id); 
            string childId = psIdMatch.Groups[1].Value;
            if (psIdMatch.Success)
            {
                while (childId.Length > 1)
                {
                    ids.Add(childId);
                    childId = childId.Remove(1, 1);
                }
            }
            // else if plan id lists all PTVs individually (c1A1BA1CBA1DCBA)
            else
            {
                string pattern = @"\d[A-Z]+";
                foreach (Match m in Regex.Matches(ps.Id, pattern))
                {
                    ids.Add(m.Value.ToString());
                }
            }

            foreach(string id in ids)
            {
                if (ss.Structures.Any(x => x.Id.StartsWith("PTV_" + id)))
                {
                    Structure smatch;
                    try
                    {
                        smatch = ss.Structures.Single(x => x.Id.StartsWith("PTV_" + id));
                    }
                    catch (Exception)
                    {
                        //MessageBox.Show("Mehr als ein PTV mit Id PTV_" + id + " gefunden.");
                        throw new Exception("Mehr als ein PTV mit Id PTV_" + id + " gefunden. Bitte altes PTV umbenennen (z.B. z_PTV_" + id + ")");
                    }
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
