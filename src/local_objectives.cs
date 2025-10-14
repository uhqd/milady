using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using Newtonsoft.Json;
namespace catchDose
{
    internal class local_objectives
    {

        public List<Ligne> lignes { get; set; }
        public int nObjectifs { get { return lignes.Count; } }
        public local_objectives(ScriptContext _ctx, PreliminaryInformation _pinfo) //constructor
        {

            lignes = new List<Ligne>();
            DateTime dh = DateTime.Now;
            String nf = ((int)_ctx.PlanSetup.NumberOfFractions).ToString();
            String ipp = _ctx.Patient.Id;
            string cid = _ctx.Course.Id;
            string pid = _ctx.PlanSetup.Id;
            string statut = "OK";

            int nTotal = 0;
            if (_ctx.PlanSetup.GetClinicalGoals() != null)
                nTotal = _ctx.PlanSetup.GetClinicalGoals().Count();
            if (nTotal != 0)
            {

                string msg = string.Empty;
                foreach (ClinicalGoal cg in _ctx.PlanSetup.GetClinicalGoals())
                {

                    string dvhpoint = stringTheObjective(cg);

                    string comparateur = "";
                    if (cg.Objective.Operator == ObjectiveOperator.LessThan)
                    {
                        comparateur = "<";
                    }
                    if (cg.Objective.Operator == ObjectiveOperator.LessThanOrEqual)
                    {
                        comparateur = "<=";
                    }
                    if (cg.Objective.Operator == ObjectiveOperator.GreaterThan)
                    {
                        comparateur = ">";
                    }
                    if (cg.Objective.Operator == ObjectiveOperator.GreaterThanOrEqual)
                    {
                        comparateur = ">=";
                    }
                    if (cg.Objective.Operator == ObjectiveOperator.Equals)
                    {
                        comparateur = "#";
                    }


                    double d = cg.Objective.Limit;
                    double va = cg.VariationAcceptable;
                    double av = cg.ActualValue;
                    if (cg.Objective.Type == ObjectiveGoalType.Volume)
                        if (cg.Objective.LimitUnit.ToString().Contains("Absolute"))
                        {
                            d = d / 1000;// mm3 --> cc
                            va = va / 1000;// mm3 --> cc
                            //av = av / 1000;// mm3 --> cc
                        }
                    string objectif = d.ToString("F1");
                    string variation = va.ToString("F1");
                    string valeur_plan = av.ToString("F2");


                    statut = "ok";
                    if (cg.EvaluationResult.ToString().ToLower().Contains("variation"))
                        statut = "var";
                    if (cg.EvaluationResult.ToString().ToLower().Contains("failed"))
                        statut = "false";


                    Ligne l = new Ligne(dh, nf, ipp, cid, pid, cg.StructureId, dvhpoint, comparateur, objectif, variation, valeur_plan, statut);


                    lignes.Add(l);




                }


            }



        }

        public void writeJson(string fullpath)
        {
            try
            {
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(this.lignes, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(fullpath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error writing json file " + fullpath + "\n" + ex.Message);
            }
        }
        public string stringTheObjective(ClinicalGoal cg)
        {
            Objective obj = cg.Objective;

            // MessageBox.Show("Objective : " + cg.StructureId + " value: " + obj.Value + "     " + cg.ObjectiveAsString);

            string s = "";
            if (obj.Type == ObjectiveGoalType.Dose)
            {
                string valueUnit = obj.ValueUnit.ToString().Replace("Relative", "%").Replace("Absolute", "cc");
                string limitUnit = obj.LimitUnit.ToString().Replace("Relative", "%").Replace("Absolute", "Gy");

                double v = obj.Value;
                if (obj.LimitUnit.ToString().Contains("Absolute"))
                    v = v / 1000; // mm3 --> cc

                s = "D" + v.ToString("F1") + valueUnit + "[" + limitUnit + "]";

            }
            if (obj.Type == ObjectiveGoalType.Volume)
            {

                string valueUnit = obj.ValueUnit.ToString().Replace("Relative", "%").Replace("Absolute", "Gy");
                string limitUnit = obj.LimitUnit.ToString().Replace("Relative", "%").Replace("Absolute", "cc");


                double d = obj.Value;

                s = "V" + d.ToString("F1") + valueUnit + "[" + limitUnit + "]";

            }

            if (obj.Type == ObjectiveGoalType.Prescription)
            {
                s = "Presc[" + obj.LimitUnit.ToString() + "]";
            }
            if (obj.Type == ObjectiveGoalType.GradientMeasure)
            {
                s = "GI";
            }
            if (obj.Type == ObjectiveGoalType.ConformityIndex)
            {
                s = "CI";
            }
            if (obj.Type == ObjectiveGoalType.Minimum_Dose)
            {
                s = "Min[" + obj.LimitUnit.ToString() + "]";
                s = s.Replace("Relative", "%");
                s = s.Replace("Absolute", "Gy");
            }
            if (obj.Type == ObjectiveGoalType.Maximum_Dose)
            {
                s = "Max[" + obj.LimitUnit.ToString() + "]";
                s = s.Replace("Relative", "%");
                s = s.Replace("Absolute", "Gy");
            }
            if (obj.Type == ObjectiveGoalType.Mean_Dose)
            {
                s = "Mean[" + obj.LimitUnit.ToString() + "]";
                s = s.Replace("Relative", "%");
                s = s.Replace("Absolute", "Gy");
            }

            // s += "\nvariation " + cg.VariationAcceptable.ToString("F2");
            // MessageBox.Show(s);
            return s;
        }
    }
}
