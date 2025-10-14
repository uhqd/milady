//using PlanCheck.Users;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Runtime.Remoting.Contexts;
using System.Runtime.InteropServices;
using System.Text;


namespace catchDose
{
    public class PreliminaryInformation
    {
        #region --------------------------------VARIABLES-----------------------------------------------
        public bool tomoFailedToFindReport { get; set; }
        public string tomoApproverName { get; set; }
        public string tomoApprovalDate { get; set; }
        private ScriptContext _ctx;
        private string _patientname;
        private string _patientdob;
        private DateTime _patientdob_dt;
        private string _coursename;
        private string _planname;
        private string _treatmentType;
        private bool _NOVA;
        private bool _HALCYON;
        private bool _SRS;
        private bool _HYPERARC;
        private bool _isModulated;
        private bool _isFE;
        private bool _isDCA;
        private bool _isDIBH;
        private int _nFraction;
        private string _machine;
        private List<DateTime> dosimetrie = new List<DateTime>();
        private List<DateTime> dosecheck = new List<DateTime>();
        private List<DateTime> ficheDePosition = new List<DateTime>();
        private List<DateTime> autres = new List<DateTime>();
        List<string> listOfTargets = new List<string>();
        List<string> listOfStructures = new List<string>();
        List<(string, string)> targetsAndStructList = new List<(string, string)>();

        public string _creator { get; set; }
        public string _user { get; set; }
        public string _doctor { get; set; }
        public bool _isSTIC { get; set; }
        public bool _isSTEC { get; set; }
        public bool _isTOMO { get; set; }
        public string approbationDate { get; set; }
        #endregion

        #region -----------------------------------------CONSTRUCTOR----------------------------------------------------------------------------------
        public PreliminaryInformation(ScriptContext ctx)  //Constructor
        {
           
            _ctx = ctx;
    
      
            if (_ctx.Patient.Name != null)
                _patientname = _ctx.Patient.Name;
            else
                _patientname = "no name";

            if (_ctx.Patient.DateOfBirth.HasValue)
            {
                _patientdob_dt = (DateTime)_ctx.Patient.DateOfBirth;
                _patientdob = _patientdob_dt.Day + "/" + _patientdob_dt.Month + "/" + _patientdob_dt.Year;
            }
            else
                _patientdob = "no DoB";
          




            

            docFinder trf = new docFinder(ctx);



           

        }
        #endregion





        #region ------------------------------------GETS/SETS ---------------------------------------------------------------------------------------------
        public string PatientName
        {
            get { return _patientname; }
        }


        public string PatientDOB
        {
            get { return _patientdob; }
        }

        public DateTime PatientDOB_dt
        {
            get { return _patientdob_dt; }
        }
        public string CourseName
        {
            get { return _coursename; }
        }
        public string PlanName
        {
            get { return _planname; }
        }
        public bool isModulated
        {
            get { return _isModulated; }
        }
        public string treatmentType
        {
            get { return _treatmentType; }
        }

        public void setTreatmentType(string type)
        {
            _treatmentType = type;
        }

        public bool isFE
        {
            get { return _isFE; }
        }
        public bool isNOVA
        {
            get { return _NOVA; }
        }
        public bool isDCA
        {
            get { return _isDCA; }
        }
        public bool isDIBH
        {
            get { return _isDIBH; }
        }

        public bool isSRS
        {
            get { return _SRS; }
        }
        public bool isHALCYON
        {
            get { return _HALCYON; }
        }
        public bool isHyperArc
        {
            get { return _HYPERARC; }
        }

        public string machine
        {
            get { return _machine; }
        }



        public List<string> mylistOfTargets
        {
            get { return listOfTargets; }
        }
        public List<string> mylistOfStructures
        {
            get { return listOfStructures; }
        }
        public List<(string, string)> mytargetsAndStructList
        {
            get { return targetsAndStructList; }
        }


        public int nFractions
        {
            get { return _nFraction; }
        }



        #endregion


    }


}

