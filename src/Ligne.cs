using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace catchDose
{


    public class Ligne
    {
        //  constructor for reading
        public Ligne()
        {
            
        }
        //  constructor for writing
        


        public Ligne(DateTime dh, string loc, string ipp, string course, string plan, string structure, string dvhpoint, string comparateur, string objectif, string variation, string valeur_plan, string statut)
        {
            this.date_heure = dh;
            this.loc_nbrefractions = loc;
            this.ipp = ipp;
            this.course = course;
            this.plan = plan;
            this.structure = structure;
            this.dvhpoint = dvhpoint;
            this.comparateur = comparateur;
            this.objectif = objectif;
            this.variation = variation;
            this.valeur_plan = valeur_plan;
            this.statut = statut;
        }

        public DateTime date_heure { get; set; }
        public string loc_nbrefractions { get; set; }
        public string ipp { get; set; }
        public string course { get; set; }
        public string plan { get; set; }
        public string structure { get; set; }
        public string dvhpoint { get; set; }
        public string comparateur { get; set; }
        public string objectif { get; set; } // Peut être null
        public string variation { get; set; }
        public string valeur_plan { get; set; }
        public string statut { get; set; }
    }


}
