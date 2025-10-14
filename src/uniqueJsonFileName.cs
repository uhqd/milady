using System;


namespace catchDose
{
    internal class uniqueJsonFileName
    {

        public uniqueJsonFileName()
        {
            filename = string.Empty;
            IPP = string.Empty;
            courseID = string.Empty;
            planID = string.Empty;

        }

        public string makefilename(string ipp, string courseID, string planID)
        {
            filename = Uri.EscapeDataString(planID + "_" + courseID + "_" + ipp + ".json");
            return filename;
        }
        public void unmakefilename(string afilename)
        {
            this.filename = afilename;
            string decoded = Uri.UnescapeDataString(filename);
            string[] parts = decoded.Split(new char[] { '_' }, 3);
            if (parts.Length == 3)
            {
                this.IPP = parts[0];
                this.courseID = parts[1];
                this.planID = parts[2].Replace(".json", "");
            }
            else
            {
                this.planID = string.Empty;
                this.courseID = string.Empty;
                this.IPP = string.Empty;
            }
        }
        public string filename { get; set; }
        public string IPP { get; set; }
        public string courseID { get; set; }
        public string planID { get; set; }


        /*
         *** exemple creation filename
        
        uniqueJsonFileName ujfn = new uniqueJsonFileName();
        string fname = ujfn.makefilename(context.Patient.Id,context.Course.Id,context.PlanSetup.Id);


        **** exemple parsing filename
        uniqueJsonFileName ujfn = new uniqueJsonFileName();
        ujfn.unmakefilename("PLAN_A_COURSE1_12345.json");
        string ipp = ujfn.IPP; // "12345"
        string course = ujfn.courseID; // "COURSE1"
        string plan = ujfn.planID; // "PLAN_A"
            

         
         */


    }
}
