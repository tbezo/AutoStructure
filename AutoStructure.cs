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
            StructureCodeDictionary scdvms = context.StructureCodes.VmsStructCode;

            //Create loopable list of PTVs
            IEnumerable<Structure> ptvs = GetPTVsFromPlanname(context.ExternalPlanSetup, ss);

            // Create merged PTV and/or find largest PTV
            ptvmax = CreateMaxMergedPTV(ptvs, ss, scdvms);
            if (!ptvs.Contains(ptvmax))
            {
                ptvs = ptvs.Concat(ss.Structures.Where(x => x.Id.StartsWith(ptvmax.Id)).ToList());
            }

            //Create loopable list of OARs to crop
            IEnumerable<Structure> oars = ss.Structures.Where(x => x.Id.StartsWith("OAR")).ToList();

            //Create OAR optimization help structures (3mm cropped)
            CreateCroppedOARs(oars, ss, ptvmax, scdvms, 3.0);

            //Create loopable list of PRVs
            IEnumerable<Structure> prvs = ss.Structures.Where(x => x.Id.StartsWith("PRV")).ToList();
            // warn user if one or more PRVs overlap with PTV, he may need to crop
            WarnOnPRVs(prvs, ss, ptvmax);

            //Create Ring around largest or compound PTV
            CreateRing(ptvs, ss, scdvms, 3.0);

            //Create Cropped PTVs if there are more than just one (only after cropping OARs and testing for PRVs)
            CreateCroppedPTVs(ptvs, ss, scdvms, 3.0);
        }


        ///<summary>Combines all PTV in the given list to create a merged PTV</summary>
        ///<param name="ptvs">IEnumerable containing structures that should get a ring</param>
        ///<param name="ss">Structure set to operate on</param>
        ///<returns>Merged PTV or largest one.</returns>
        static Structure CreateMaxMergedPTV(IEnumerable<Structure> ptvs, StructureSet ss, StructureCodeDictionary scd)
        {
            Structure ptvges;
            Structure ptvmax;
            Structure tmpMerge;
            Structure tmpPtv;

            // test if only single PTV is in the list skip the rest
            if (ptvs.Count() == 1)
            {
                return ptvs.FirstOrDefault();
            }

            //create or use tmp Structures 
            try { tmpMerge = ss.AddStructure("CONTROL", "tmpMerge"); }
            catch { tmpMerge = ss.Structures.Single(x => x.Id == "tmpMerge"); }
            try { tmpPtv = ss.AddStructure("CONTROL", "tmpPtv"); }
            catch { tmpPtv = ss.Structures.Single(x => x.Id == "tmpPtv"); }

            //create merge PTV 
            try
            {
                ptvges = ss.AddStructure("PTV", "z_PTV_ges");
                ptvges.StructureCode = scd["PTV_Low"];
            }
            catch { ptvges = ss.Structures.Single(x => x.Id == "z_PTV_ges"); }

            if (ptvs.Where(x => x.IsHighResolution).Any())
            {
                ptvges.ConvertToHighResolution();
                tmpMerge.ConvertToHighResolution();
                tmpPtv.ConvertToHighResolution();
            }

            // select largest PTV 
            ptvmax = ptvs.OrderByDescending(x => x.Volume).FirstOrDefault();

            if (ptvges.IsEmpty) //if there is a user generated merged PTV do not change that
            {
                foreach (Structure ptv in ptvs)
                {
                    //merge ptv with ptvges
                    if (ptvges.IsHighResolution && !ptv.IsHighResolution)
                    {
                        tmpPtv.SegmentVolume = ptv.SegmentVolume;
                        tmpPtv.ConvertToHighResolution();
                        tmpMerge.SegmentVolume = ptvges.Or(tmpPtv);
                    }
                    else { tmpMerge.SegmentVolume = ptvges.Or(ptv); }     
                    
                    //if merged structure is bigger use new structure as ptvges
                    if (tmpMerge.Volume > ( ptvges.Volume - ptvges.Volume * 0.001)) // 0.001: ignore rounding errors - maybe a fixed value would be better
                    {
                        ptvges.SegmentVolume = tmpMerge.SegmentVolume;
                    }
                }                
                if (ptvmax.Volume < (ptvges.Volume - ptvges.Volume * 0.001) ) // 0.001: ignore rounding errors - maybe a fixed value would be better
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
            ss.RemoveStructure(tmpMerge);
            ss.RemoveStructure(tmpPtv);
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
            Structure tmp;
            Structure tmpring;
            Color ringColor = Color.FromArgb(255, 255, 165, 0);

            //Regex that matches on PTV number and removes trailing date etc, also matches z_PTV_ges
            Regex ptvreg1 = new Regex(@"PTV_(\d?[A-Za-z]+)", RegexOptions.Compiled);

            //Use only largest (compound) PTV to create a ring for the plan.
            Structure tptv = ptvs.OrderByDescending(x => x.Volume).FirstOrDefault();

            Match ptvmatch = ptvreg1.Match(tptv.Id);
            string ptvid = ptvmatch.Groups[1].Value;
            if (ptvid == "")
            {
                MessageBox.Show("Found wrongly named PTV!");
                return;
            }
            
            try { tmpring = ss.AddStructure("CONTROL", "z_Ring_" + ptvid); } //does not exist yet
            catch { tmpring = ss.Structures.FirstOrDefault(x => x.Id == "z_Ring_" + ptvid); } //already exists

            if (tmpring.IsEmpty)
            {
                //if there are high resolution PTVs, create a high resolution ring
                if (ptvs.Where(x => x.IsHighResolution).Any())
                {
                    try { tmp = ss.AddStructure("CONTROL", "tmp"); }
                    catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }
                    tmpring.ConvertToHighResolution();

                    tmp.SegmentVolume = tptv.SegmentVolume;
                    if (!tmp.IsHighResolution) { tmp.ConvertToHighResolution(); }
                    tmpring.SegmentVolume = tmpring.Or(tmp.Margin(20.0));

                    //Crop PTV and any child PTVs from ring with larger distance
                    foreach (Structure cptv in ptvs)
                    {
                        tmp.SegmentVolume = cptv.SegmentVolume; // how timeconsuming is it to convert to high resolution?
                        if (!tmp.IsHighResolution) { tmp.ConvertToHighResolution(); }
                        tmpring.SegmentVolume = tmpring.Sub(tmp.Margin(cropmm + 2.0 * GetParentCount(ptvs, cptv.Id)));
                    }
                    ss.RemoveStructure(tmp);
                }
                else
                {
                    tmpring.SegmentVolume = tmpring.Or(tptv.Margin(20.0));
                    //Crop PTV and any child PTVs from ring with larger distance
                    foreach (Structure cptv in ptvs)
                    {
                        tmpring.SegmentVolume = tmpring.Sub(cptv.Margin(cropmm + 2.0 * GetParentCount(ptvs, cptv.Id)));
                    }
                }
                tmpring.Color = ringColor;
                tmpring.StructureCode = scd["Ring"];                
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
            Structure tmp;

            //Create optimization structures from PTVs
            foreach (Structure tptv in ptvs)
            {
                //get parent id
                string tmpname = GetParentId(ptvs, tptv.Id);
                //if parent was found crop it
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
                        // handle situation where one ptv is in high resolution and not the other
                        if (parentPtv.IsHighResolution ^ tptv.IsHighResolution)
                        {
                            try { tmp = ss.AddStructure("CONTROL", "tmp"); }
                            catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }

                            if (!tptv.IsHighResolution)
                            {
                                zptv.SegmentVolume = tptv.SegmentVolume;
                                zptv.SegmentVolume = tptv.Margin(cropmm);
                                zptv.ConvertToHighResolution();
                            }
                            else 
                            {
                                zptv.ConvertToHighResolution();
                                zptv.SegmentVolume = tptv.Margin(cropmm);
                            }
                            if (!parentPtv.IsHighResolution) 
                            {
                                tmp.SegmentVolume = parentPtv.SegmentVolume;
                                tmp.ConvertToHighResolution();
                            }
                            else
                            {
                                tmp.ConvertToHighResolution();
                                tmp.SegmentVolume = parentPtv.SegmentVolume;
                            }
                            zptv.SegmentVolume = tmp.Sub(zptv);
                            ss.RemoveStructure(tmp);
                        }
                        else { zptv.SegmentVolume = parentPtv.Sub(tptv.Margin(cropmm)); } // both have same resolution
                        
                        if (ptvs.Any(x => x.Id.Equals("z_PTV_ges")) && Regex.IsMatch(tptv.Id, @"^PTV_\d[A-Z]{2}")) // would not crop PTV_1CBA from z_PTV_ges!
                        {
                            Structure tges = ptvs.FirstOrDefault(x => x.Id.Equals("z_PTV_ges"));
                            if (tges.IsHighResolution ^ tptv.IsHighResolution)
                            {
                                try { tmp = ss.AddStructure("CONTROL", "tmp"); }
                                catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }

                                if (tges.IsHighResolution)
                                {
                                    tmp.SegmentVolume = tptv.SegmentVolume;
                                    tmp.ConvertToHighResolution();
                                    tges.SegmentVolume = tges.Sub(tmp.Margin(cropmm));
                                }
                                if (tptv.IsHighResolution)
                                {
                                    tges.ConvertToHighResolution();
                                    tges.SegmentVolume = tges.Sub(tptv.Margin(cropmm));
                                }
                            }
                            else { tges.SegmentVolume = tges.Sub(tptv.Margin(cropmm)); }
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
            Structure tmp;
            Structure tmpPtv;

            try { tmp = ss.AddStructure("CONTROL", "tmp"); }
            catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }
            try { tmpPtv = ss.AddStructure("CONTROL", "tmpPtv"); }
            catch { tmpPtv = ss.Structures.Single(x => x.Id == "tmpPtv"); }

            // loop for all non high resolution OARs
            if(oars.Where(x => !x.IsHighResolution).Any())
            {
                if (ptvcrop.IsHighResolution)
                {
                    tmpPtv.SegmentVolume = HightoLow(ss, ptvcrop);
                }
                else { tmpPtv.SegmentVolume = ptvcrop.SegmentVolume; }

                foreach (Structure str in oars.Where(x => !x.IsHighResolution))
                {
                    CreateCroppedStr(tmp, tmpPtv, str, ss, scd);
                }
            }

            // loop for high resolution OARs
            if(oars.Where(x => x.IsHighResolution).Any())
            {
                if (!ptvcrop.IsHighResolution)
                {
                    tmpPtv.SegmentVolume = ptvcrop.SegmentVolume;
                    tmpPtv.ConvertToHighResolution();
                }
                else { tmpPtv.SegmentVolume = ptvcrop.SegmentVolume; }
                tmp.ConvertToHighResolution();

                foreach (Structure str in oars.Where(x => x.IsHighResolution))
                {
                    CreateCroppedStr(tmp, tmpPtv, str, ss, scd);                   
                }
            }
            ss.RemoveStructure(tmp);
            ss.RemoveStructure(tmpPtv);
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
        static void WarnOnPRVs(IEnumerable<Structure> prvs, StructureSet ss, Structure ptvmax, bool silent=false)
        {
            List<string> warn_prv = new List<string>();
            Structure tmp;
            Structure ptvmax_alt;

            try { tmp = ss.AddStructure("CONTROL", "tmp"); }
            catch { tmp = ss.Structures.Single(x => x.Id == "tmp"); }

            try { ptvmax_alt = ss.AddStructure("CONTROL", "ptvmax_alt"); }
            catch { ptvmax_alt = ss.Structures.Single(x => x.Id == "ptvmax_alt"); }

            // handle different resolutions of PTV and PRVs, keep number of high resolution structures low
            if (prvs.Where(x => !x.IsHighResolution).Any())
            {
                if (ptvmax.IsHighResolution)
                {
                    ptvmax_alt.SegmentVolume = HightoLow(ss, ptvmax);
                }
                else { ptvmax_alt.SegmentVolume = ptvmax.SegmentVolume; }
                // check for overlap
                foreach (Structure prv in prvs.Where(x => !x.IsHighResolution))
                {
                    tmp.SegmentVolume = ptvmax_alt.And(prv);
                    if (tmp.Volume != 0.0)
                    {
                        warn_prv.Add(prv.Id);
                    }
                }
            }
            if (prvs.Where(x => x.IsHighResolution).Any())
            {
                tmp.ConvertToHighResolution();
                if (!ptvmax.IsHighResolution)
                {
                    ptvmax_alt.SegmentVolume = ptvmax.SegmentVolume;
                    ptvmax_alt.ConvertToHighResolution();
                }
                else { ptvmax_alt.SegmentVolume = ptvmax.SegmentVolume; }
                // check for overlap
                foreach (Structure hrprv in prvs.Where(x => x.IsHighResolution))
                {
                    tmp.SegmentVolume = ptvmax_alt.And(hrprv);
                    if (tmp.Volume != 0.0)
                    {
                        warn_prv.Add(hrprv.Id);
                    }
                }
            }

            if(!silent)
            {
                if (warn_prv.Any())
                {
                    if (warn_prv.Count() > 1)
                    {
                        MessageBox.Show("Structures " + string.Join(", ", warn_prv) + " overlap with PTV.", "PRV Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Structure " + string.Join(", ", warn_prv) + " overlaps with PTV.", "PRV Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }

            //remove structures not used anymore
            ss.RemoveStructure(ptvmax_alt);
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
                string pattern = @"\d[A-Z]+(?![a-z])"; // (?![a-z]) keeps from matching on XGy (X = Number) in the plan id
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
                        throw new Exception("Found more than one PTV with Id PTV_" + id + ". Please rename old PTV (a.e. to z_PTV_" + id + ")");
                    }
                    if (!ptvs.Contains(smatch))
                    {
                        ptvs.Add(smatch);
                    }
                }
                else
                {
                    MessageBox.Show("No structure found with Id PTV_" + id + "!");
                }

            }

            try
            {
                ptvs.First();
            }
            catch (Exception)
            {
                throw new Exception("No PTVs found matching plan name!");
            }
            return ptvs;
        }

        /// <summary>
        /// Function to convert high resolution structures to normal/low resolution. Function takes some time, try to avoid.
        /// Originally from https://www.reddit.com/r/esapi/comments/mbbxa3/low_and_high_resolution_structures/
        /// </summary>
        /// <param name="ss">structureset to operate on</param>
        /// <param name="str">structure to convert</param>
        /// <returns>Structure SegementVolume in normal resolution</returns>
        static SegmentVolume HightoLow(StructureSet ss, Structure str)
        {
            Structure lowresstructure = ss.AddStructure("CONTROL", str.Id + "_lr");
            SegmentVolume lowresvolume;

            System.Windows.Media.Media3D.Rect3D mesh = str.MeshGeometry.Bounds;
            int meshLow = GetSlice(mesh.Z, ss);
            int meshUp = GetSlice(mesh.Z + mesh.SizeZ, ss) + 1;

            for (int j = meshLow; j <= meshUp; j=j+3) // do not use all segments for more speed
            {
                VVector[][] contours = str.GetContoursOnImagePlane(j);
                if (contours.Any())
                {
                    foreach (VVector[] segment in contours) { lowresstructure.AddContourOnImagePlane(segment, j); }
                }
            }

            lowresvolume = lowresstructure.SegmentVolume;
            ss.RemoveStructure(lowresstructure);
            return lowresvolume;
        }

        /// <summary>
        /// In which Slice this VVector is? It depends on the image and Structure set you are accessing. 
        /// From https://jhmcastelo.medium.com/tips-for-vvectors-and-structures-in-esapi-575bc623074a
        /// </summary>
        /// <param name="z">z-coordinate from vector</param>
        /// <param name="SS">Structureset</param>
        /// <returns>Slicenumber</returns>
        static int GetSlice(double z, StructureSet ss)
        {
            var imageRes = ss.Image.ZRes;
            return Convert.ToInt32((z - ss.Image.Origin.z) / imageRes);
        }

        /// <summary>
        /// Method to create a cropped structure if two structures overlap or are within margin of each other. Nerves are ignored.
        /// </summary>
        /// <param name="tmp">Temporary structure to work with</param>
        /// <param name="tmpPtv">PTV to test overlap with</param>
        /// <param name="str">Structure to test overlap</param>
        /// <param name="ss">Structureset to work with</param>
        /// <param name="scd">Structure code Dictionary</param>
        /// <param name="cropmm">Additional distance in mm from PTV to test</param>
        /// <returns>True if structure was created, false if Structures do not overlap or are not within margin.</returns>
        static bool CreateCroppedStr(Structure tmp, Structure tmpPtv, Structure str, StructureSet ss, StructureCodeDictionary scd, double cropmm=3)
        {
            Structure zoar;          

            tmp.SegmentVolume = tmpPtv.And(str.Margin(cropmm));
            if (tmp.Volume != 0.0 && !(str.Id.Contains("Spinal") || str.Id.Contains("HS") || str.Id.Contains("Opt") || str.Id.Contains("Chia"))) //nerves do not get cropped!
            {
                try
                {
                    zoar = ss.AddStructure("CONTROL", "z_" + str.Id.Substring(4));
                    zoar.Color = str.Color;
                    zoar.StructureCode = scd["Control"];
                }
                catch
                {
                    zoar = ss.Structures.FirstOrDefault(x => x.Id == "z_" + str.Id.Substring(4));
                }
                if (zoar.IsEmpty && !zoar.IsApproved)
                {
                    zoar.SegmentVolume = str.Sub(tmpPtv.Margin(cropmm));
                    if (zoar.IsEmpty) 
                    { 
                        ss.RemoveStructure(zoar);
                        return false;
                    } //if z_OAR is still empty we can remove it
                }
                return true;
            }
            return false;
        }

    }
}
