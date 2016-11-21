/**************************************************************************************************
 *                             ___   ___   _____  ___       _       
 *                            / __\ / __\ /   \ \/ / | __ _| |_ ___ *
 *                           / /   / /   / /\ /\  /| |/ _` | __/ _ \
 *                          / /___/ /___/ /_// /  \| | (_| | ||  __/
 *                          \____/\____/___,' /_/\_\_|\__,_|\__\___|
 *                                       
 *
 *  CCDXlate provides batch translation of patient Continuing Care Document (CCD) XML files
 *  into several specific CSV files to be used for import into an Electronic Medical Records
 *  (EMR) system.
 *  
 *  CCDXlate requires that all the CCD files be placed in a single directory (pending by default)
 *  and have a common file specification (ccd*.xml by default).  All of the CCD files are 
 *  translated with each section from each CCD file translated to the following CSV files:
 *  
 *      - allergies.csv
 *      - medications.csv
 *      - problems.csv
 *      - procedures.csv
 *      - vitals.csv
 *  
 *  Once processed, CCD files successfully translated are moved to a target directory (processed
 *  by default) while those files that could not be translated due to an error are moved to
 *  an error directory (errory by default).
 *  
 *  A report of the batch translation is included in the file:
 *  
 *      - report.csv
 *    
 *  Lastly, the batch translation may be performed in stages by placing additional files in the
 *  pending directory and then re-running the CCDXLATE executable with the /a (append) parameter.
 *  When run in the append mode, CCDXlate will append records to the *.csv output files.
 *  
 **************************************************************************************************  
 *      * Ascii art generator: http://patorjk.com/software/taag/#p=display&f=Ogre&t=CCDXlate
 *      
 *                                   (c) TJRTech, Inc. 2016
 *************************************************************************************************/
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Xml;
using System.Globalization;

namespace ccdxlate
{
    class Program
    {
        //*************************************************************************************
        //*** CONSTANTS                                                                     ***
        //*************************************************************************************
        public const string moduleName = "ccdxlate";    // Module name
        public const String ALLERGIES_SECTION_TITLE = "Allergies, Adverse Reactions, Alerts";
        public const String MEDICATIONS_SECTION_TITLE = "Medications";
        public const String PROBLEMS_SECTION_TITLE = "Problems";
        public const String PROCEDURES_SECTION_TITLE = "Procedures";
        public const String VITALS_SECTION_TITLE = "Vital Signs";
        public const String ALLERGIES_FILE = "allergies.csv";
        public const String MEDICATIONS_FILE = "medications.csv";
        public const String PROBLEMS_FILE = "problems.csv";
        public const String PROCEDURES_FILE = "procedures.csv";
        public const String VITALS_FILE = "vitals.csv";
        public const String REPORT_FILE = "report.csv";
        public const String SOURCE_DIR = "pending";
        public const String TARGET_DIR = "processed";
        public const String ERROR_DIR = "error";
        public const String FILE_SPEC = "CCD*.XML";
        public const String ALLERGIES_CSV_HEADER = "UniqueId,PatientId,FileInfo,AllergyName,Status,StartDate,ICD9,Reaction,AllergyCategory,Note";
        public const String MEDICATIONS_CSV_HEADER = "UniqueId,PatientId,FileInfo,DrugName,Status,MedId,StartDate,SIG,CareProviderName,MedicationStrength,MedicationStrengthUnit,Refills,DispenseAmount,DispenseUnit,DoseForm,Route,DispenseAsWritten,DispensableDrugName,Frequency";
        public const String PROBLEMS_CSV_HEADER = "UniqueId,PatientId,FileInfo,ProblemName,Status,StartDate,ICD9,Comment";
        public const String PROCEDURES_CSV_HEADER = "UniqueId,PatientId,FileInfo,SurgeryName,StartDate,ICD9,Comment";
        public const String VITALS_CSV_HEADER = "UniqueId,PatientId,FileInfo,Date,WeightInGrams,HeightInCm,TemperatureInCelsius,BP-Sys,BP-Dia,Pulse,RespRate,O2";
        public const String REPORT_CSV_HEADER = "FileNum,File,Status,NumAllergies,NumMedications,NumProblems,NumProcedures,NumVitals";


        //*************************************************************************************
        //*** CLASSES                                                                       ***
        //*************************************************************************************
  
        /*** Allergy class ***/
        public class Allergy
        {
            // Public properties
            public bool isInitialized = false;
            public String AllergyName;
            public String Status;
            public DateTime StartDate;
            public String ICD9;
            public String Reaction;
            public String AllergyCategory;
            public String Note;

            // Constructor
            public Allergy()
            {
                isInitialized = true;
                AllergyName = "";
                Status = "";
                StartDate = DateTime.MinValue;
                ICD9 = "";
                Reaction = "";
                AllergyCategory = "";
                Note = "";
            }

            // isEmpty - returns true if record is empty and should not be displayed, imported.
            public bool isEmpty()
            {
                return ( string.IsNullOrWhiteSpace(AllergyName) );
            }

            // toCSVString - returns a CSV formatted string suitable for writing to import file
            public String toCSVString(String UniqueId, String PatientId, String FileInfo)
            {
                return UniqueId + "," +
                       PatientId + "," +
                       FileInfo + "," +
                       QuoteCSVString(AllergyName) + "," +
                       QuoteCSVString(Status) + "," +
                       FormatCSVDate(StartDate) + "," +
                       QuoteCSVString(ICD9) + "," +
                       QuoteCSVString(Reaction) + "," +
                       QuoteCSVString(AllergyCategory) + "," +
                       QuoteCSVString(Note);
            }
        };


        /*** Medication class ***/
        public class Medication
        {
            // Public properties
            public bool isInitialized = false;
            public String DrugName;
            public String Status;
            public String MedId;
            public DateTime StartDate;
            public DateTime EndDate;
            public String SIGNum;
            public String SIG;
            public String CareProviderName;
            public String MedicationStrength; 
            public String MedicationStrengthUnit;
            public String Refills;
            public String DispenseAmount;
            public String DispenseUnit;
            public String DoseForm;
            public String Route;
            public String DispenseAsWritten;
            public String DispensableDrugName;
            public String Frequency;

            // Constructor
            public Medication()
            {
                isInitialized = true;
                DrugName = "";
                Status = "";
                MedId = "";
                StartDate = DateTime.MinValue;
                EndDate = DateTime.MinValue;
                SIGNum = "";
                SIG = "";
                CareProviderName = "";
                MedicationStrength = "";
                MedicationStrengthUnit = "";
                Refills = "";
                DispenseAmount = "";
                DispenseUnit = "";
                DoseForm = "";
                Route = "";
                DispenseAsWritten = "";
                DispensableDrugName = "";
                Frequency = "";
            }

            // isEmpty - returns true if record is empty and should not be displayed, imported.
            public bool isEmpty()
            {
                return ( string.IsNullOrWhiteSpace(DrugName) );
            }

            // toCSVString - returns a CSV formatted string suitable for writing to import file
            public String toCSVString(String UniqueId, String PatientId, String FileInfo)
            {
                return UniqueId + "," +
                       PatientId + "," +
                       FileInfo + "," +
                       QuoteCSVString(DrugName) + "," +
                       QuoteCSVString(Status) + "," +
                       QuoteCSVString(MedId) + "," +
                       FormatCSVDate(StartDate) + "," +
                       QuoteCSVString(SIG) + "," +
                       QuoteCSVString(CareProviderName) + "," +
                       QuoteCSVString(MedicationStrength) + "," +
                       QuoteCSVString(MedicationStrengthUnit) + "," +
                       QuoteCSVString(Refills) + "," +
                       QuoteCSVString(DispenseAmount) + "," +
                       QuoteCSVString(DispenseUnit) + "," +
                       QuoteCSVString(DoseForm) + "," +
                       QuoteCSVString(Route) + "," +
                       QuoteCSVString(DispenseAsWritten) + "," +
                       QuoteCSVString(DispensableDrugName) + "," +
                       QuoteCSVString(Frequency);
            }
        };


        /*** Problem class ***/
        public class Problem
        {
            // Public properties
            public bool isInitialized = false;
            public DateTime StartDate;
            public String ProblemName;
            public String Status;
            public String ICD9;
            public String Comment;

            // Constructor
            public Problem()
            {
                isInitialized = true;
                StartDate = DateTime.MinValue;
                ProblemName = "";
                Status = "";
                ICD9 = "";
                Comment = "";
            }

            // isEmpty - returns true if record is empty and should not be displayed, imported.
            public bool isEmpty()
            {
                return ( string.IsNullOrWhiteSpace(ProblemName) );
            }

            // toCSVString - returns a CSV formatted string suitable for writing to import file
            public String toCSVString(String UniqueId, String PatientId, String FileInfo)
            {
                return UniqueId + "," +
                       PatientId + "," +
                       FileInfo + "," +
                       QuoteCSVString(ProblemName) + "," +
                       QuoteCSVString(Status) + "," +
                       FormatCSVDate(StartDate) + "," +
                       QuoteCSVString(ICD9) + "," +
                       QuoteCSVString(Comment);
            }
        }


        /*** Procedure class ***/
        public class Procedure
        {
            // Public properties
            public bool isInitialized = false;
            public DateTime StartDate;
            public String SurgeryName;
            public String ICD9;
            public String Comment;

            // Constructor
            public Procedure()
            {
                isInitialized = true;
                StartDate = DateTime.MinValue;
                SurgeryName = "";
                ICD9 = "";
                Comment = "";
            }

            // isEmpty - returns true if record is empty and should not be displayed, imported.
            public bool isEmpty()
            {
                return ( string.IsNullOrWhiteSpace(SurgeryName) );
            }

            // toCSVString - returns a CSV formatted string suitable for writing to import file
            public String toCSVString(String UniqueId, String PatientId, String FileInfo)
            {
                return UniqueId + "," +
                       PatientId + "," +
                       FileInfo + "," +
                       QuoteCSVString(SurgeryName) + "," +
                       FormatCSVDate(StartDate) + "," +
                       QuoteCSVString(ICD9) + "," +
                       QuoteCSVString(Comment);
            }
        }


        /*** VitalsReading class ***/
        public class VitalsReading
        {
            // Public properties
            public bool isInitialized = false;
            public DateTime StartDate;
            public float WeightInLbs;
            public float HeightInInches;
            public float TemperatureInFahrenheit;
            public int BPSys;
            public int BPDia;
            public int Pulse;
            public int RespRate;
            public int O2;

            // Constructor
            public VitalsReading()
            {
                isInitialized = true;
                StartDate = DateTime.MinValue;
                WeightInLbs = 0.0F;
                WeightInGrams = 0.0F;
                HeightInCm = 0.0F;
                TemperatureInCelsius = 0.0F;
                BPSys = 0;
                BPDia = 0;
                Pulse = 0;
                RespRate = 0;
                O2 = 0;
            }

            // isEmpty - returns true if record is empty and should not be displayed, imported.
            public bool isEmpty()
            {
                return (StartDate == DateTime.MinValue);
            }

            // toCSVString - returns a CSV formatted string suitable for writing to import file.
            // UniqueId,PatientId,Date,WeightInGrams,HeightInCm,TemperatureInCelsius,BP-Sys,BP-Dia,Pulse,RespRate,O2";
            public String toCSVString(String UniqueId, String PatientId, String FileInfo)
            {
                return UniqueId + "," +
                       PatientId + "," +
                       FileInfo + "," +
                       FormatCSVDate(StartDate) + "," +
                       ZeroToBlankString(WeightInGrams) + "," +
                       ZeroToBlankString(HeightInCm) + "," +
                       ZeroToBlankString(TemperatureInCelsius) + "," +
                       ZeroToBlankString(BPSys) + "," +
                       ZeroToBlankString(BPDia) + "," +
                       ZeroToBlankString(Pulse) + "," +
                       ZeroToBlankString(RespRate) + "," +
                       ZeroToBlankString(O2);
            }

            // WeightInGrams - property that converts between US and Metric units
            public float WeightInGrams
            {
                get
                {
                    return WeightInLbs * 453.592F;
                }
                set
                {

                    WeightInLbs = value / 453.592F;
                }
            }

            // HeightInCm - property that converts between US and Metric units
            public float HeightInCm
            {
                get
                {
                    return HeightInInches * 2.54F;
                }
                set
                {

                    HeightInInches = value / 2.54F;
                }
            }

            // TemperatureInCelsius - property that converts between US and Metric units
            public float TemperatureInCelsius
            {
                get
                {
                    return (TemperatureInFahrenheit - 32.0F) / 1.8F;
                }
                set
                {

                    TemperatureInFahrenheit = (value * 1.8F) + 32.0F;
                }
            }

            // ZeroToBlankString - Returns string version of the numeric if not zero else returns blank string if zero
            private static String ZeroToBlankString(float val)
            {
                if (val <= 0.0F)
                {
                    return "";
                }
                else
                {
                    return val.ToString();
                }
            }

            // ZeroToBlankString - Returns string version of the numeric if not zero else returns blank string if zero
            private static String ZeroToBlankString(int val)
            {
                if (val <= 0)
                {
                    return "";
                }
                else
                {
                    return val.ToString();
                }
            }
        };


        /*** CCDFile class ***/
        public class CCDFile
        {
            public bool isInitialized = false;
            public bool isError = false;
            public String errorMessage = "Success";
            public String FileName = "";
            public String PatientId = "";
            public String FileInfo = "";
            public String PatientName = "";
            public List<Allergy> Allergies;
            public List<Medication> Medications;
            public List<Problem> Problems;
            public List<Procedure> Procedures;
            public List<VitalsReading> VitalSigns;
            private XmlDocument xmlDoc;
            private XmlNamespaceManager nsmgr;

            // Constructor
            public CCDFile(String filepath)
            {
                isInitialized = true;
                FileName = Path.GetFileName(filepath);   // Save file name to object 

                // Load XmlDocument object and the XML file
                xmlDoc = new XmlDocument();

                // xmlDoc can fail to load which is an error.
                try
                {
                    xmlDoc.Load(filepath);
                }
                catch
                {
                    isError = true;
                    errorMessage = "ERROR: Unexpected XML document loading error.";
                    return;
                }

                // Get document namespace, add
                nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);  // Load namespace    
                nsmgr.AddNamespace("ccd", "urn:hl7-org:v3");        // Add in namespace. XML parsing via XMLDocument does not work without this.

                // Look for title element in header and if not found assume an error
                XmlNode docTitleNode = xmlDoc.SelectSingleNode("/ccd:ClinicalDocument/ccd:title", nsmgr);
                if (docTitleNode != null)
                {
                    PatientId = GetPatientId(FileName);
                    FileInfo = GetFileInfo(FileName);
                    PatientName = docTitleNode.InnerText;
                    isError = false;
                    LoadAllergies();
                    LoadMedications();
                    LoadProblems();
                    LoadProcedures();
                    LoadVitals();
                }
                else
                {
                    isError = true;
                    errorMessage = "ERROR: Document not of expected namespace (urn:hl7-org:v3) or not having /ClinicalDocument/title path.";
                }
            }


            //** CCDFile - Private Member Functions

            // GetPatientId
            //
            // Purpose:  Gets the patient id from a file name of format:  CCD<PatientId>_*, as in CCD10002-1_...
            private String GetPatientId(String filename)
            {
                int pos = filename.IndexOf("_"); 
                if (filename.Length >= 9 && filename.Substring(0, 3) == "CCD" && pos>=3)
                {
                    return filename.Substring(3,pos-3);
                }
                else
                {
                    return "";
                }   
            }

            // GetFileInfo
            //
            // Purpose:  Gets the patient id from a file name of format:  CCD10002-1_<fileInfo>.xml*, as in CCD10002-1_DOE_JANE_19660125.xml...
            private String GetFileInfo(String filename)
            {
                int pos1 = filename.IndexOf("_");
                int pos2 = filename.IndexOf(".");
                if ( pos1 >= 3 && pos2 > 3)
                {
                    return filename.Substring(pos1+1, pos2 - pos1 - 1);
                }
                else
                {
                    return "";
                }
            }


            /* LoadAllergies
             * 
             * Purpose: Loads the Allergies list from the current document
             * 
             * Allergies (48765-2)
             * 
             * CSV: UniqueId, PatientId, AllergyName, Status, StartDate, ICD9, Reaction, AllergyCategory, Note
             * 
             * XML snippet:
             * 
             *   /act 
             *      /statusCode.code                                - Status
             *      /effectiveTime/low.value (YYYYMMDD)             - StartDate
             *          /entryRelationship
             *             /value[@code="420134006”].displayName    - Reaction
             *             /observation
             *                /participant
             *                   /participantRole
             *                       /playingEntity
             *                           /name                      - DisplayName
             */
            private void LoadAllergies()
            {
                Allergies = new List<Allergy>();  
                XmlNode allergiesNode = SelectSingleNodeParent(xmlDoc, "/ccd:ClinicalDocument/ccd:component/ccd:structuredBody/ccd:component/ccd:section/ccd:code[@code='48765-2']", nsmgr);
                if (allergiesNode != null)
                {
                    string sectionTitle = NodeAttributeStringValue(allergiesNode, nsmgr, "./ccd:title", "[InnerText]");
                    if (sectionTitle == ALLERGIES_SECTION_TITLE)
                    {
                        XmlNodeList entryNodes = allergiesNode.SelectNodes("./ccd:entry", nsmgr);
                        foreach (XmlNode entryNode in entryNodes)
                        {
                            Allergy allergy = new Allergy();

                            allergy.AllergyName = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:act/ccd:entryRelationship/ccd:observation/ccd:participant/ccd:participantRole/ccd:playingEntity/ccd:name", "[InnerText]");
                            allergy.Status = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:act/ccd:statusCode", "code");
                            allergy.StartDate = NodeAttributeDateTimeValue(entryNode, nsmgr, "./ccd:act/ccd:effectiveTime/ccd:low", "value");
                            allergy.ICD9 = "";
                            allergy.Reaction = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:act/ccd:entryRelationship/ccd:observation/ccd:value[@code='420134006']", "displayName");
                            allergy.AllergyCategory = "";
                            allergy.Note = "";

                            if (allergy.Status == "active")  // Only add allergies that are active
                            {
                                Allergies.Add(allergy);  // Add the allergy
                            }
                        }
                    }
                }
            }  // LoadAllergies

            /* LoadMedications
             * 
             * Purpose: Loads the Medications list from the current document
             * 
             * Medications (10160-0)
             * 
             * CSV:  UniqueId, PatientId, DrugName, Status, MedId, StartDate, SIG, CareProviderName, MedicationStrength,
             *       MedicationStrengthUnit, Refills, DispenseAmount, DispenseUnit, DoseForm, Route, DispenseAsWritten,
             *       DispensableDrugName, Frequency
             *       
             * XML snippet:
             * 
             *      /text/table/tbody/tr/td/content[@ID='" + SigNum + "'].value     - SIG  (see SigNum below)
             *      /substanceAdministration              
             *          /statusCode.code                                            - Status
             *          /effectiveTime/low.value (YYYYMMDD)                         - StartDate
             *          /effectiveTime/high.value (YYYMMDD)                         - EndDate
             *          /routeCode.displayName                                      - Route
             *          /administrationUnitCode.displayName                         - DoseForm
             *          /text/reference.value                                       - SigNum  (used for lookup of other field)
             *          /entryRelationship.type=“REFR”
             *             /supply
             *                /statusCode.code                                      - (DO NOT USE - Always "active")
             *                /quantity.value                                       - DispenseAmount
             *                /quantity.unit                                        - DispenseUnit
             *                /product
             *                   /manufacturedProduct
             *                     /manufacturedMaterial
             *                         /code
             *                            /translation[@codeSystem='2.16.840.1.113883.6.69'].displayName    - DispensableDrugName
             *                            (parse for dosage, units )                                        - MedicationStrength
             *                                                                                              - MedicationStrengthUnit
             *                       /name                                                                  - DrugName
             *                /author/assignedAuthor/assignedPerson/name.InnerText                          - CareProviderName
             */
            private void LoadMedications()
            {
                Medications = new List<Medication>(); 
                XmlNode medicationsNode = SelectSingleNodeParent(xmlDoc, "/ccd:ClinicalDocument/ccd:component/ccd:structuredBody/ccd:component/ccd:section/ccd:code[@code='10160-0']", nsmgr);
                if (medicationsNode != null)
                {
                    string sectionTitle = NodeAttributeStringValue(medicationsNode, nsmgr, "./ccd:title", "[InnerText]");
                    if (sectionTitle == MEDICATIONS_SECTION_TITLE)
                    {

                        XmlNodeList entryNodes = medicationsNode.SelectNodes("./ccd:entry", nsmgr);
                        foreach (XmlNode entryNode in entryNodes)
                        {

                            Medication medication = new Medication();

                            medication.DrugName = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:entryRelationship/ccd:supply/ccd:product/ccd:manufacturedProduct/ccd:manufacturedMaterial/ccd:name", "[InnerText]");
                            medication.Status = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:statusCode", "code");
                            medication.MedId = "";
                            medication.StartDate = NodeAttributeDateTimeValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:effectiveTime/ccd:low", "value");
                            medication.EndDate = NodeAttributeDateTimeValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:effectiveTime/ccd:high", "value");
                            medication.SIGNum = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:text/ccd:reference", "value");
                            medication.SIG = NodeAttributeStringValue(medicationsNode, nsmgr, "./ccd:text/ccd:table/ccd:tbody/ccd:tr/ccd:td/ccd:content[@ID='" + SigNumToString(medication.SIGNum) + "']", "[InnerText]");
                            medication.CareProviderName = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:entryRelationship/ccd:supply/ccd:author/ccd:assignedAuthor/ccd:assignedPerson/ccd:name", "[InnerText]");
                            medication.Refills = "";
                            medication.DispenseAmount = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:entryRelationship/ccd:supply/ccd:quantity", "value");
                            medication.DispenseUnit = RemoveBraces(NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:entryRelationship/ccd:supply/ccd:quantity", "unit"));
                            medication.DoseForm = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:administrationUnitCode", "displayName");
                            medication.Route = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:routeCode", "displayName");
                            medication.DispenseAsWritten = "";
                            medication.DispensableDrugName = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:substanceAdministration/ccd:entryRelationship/ccd:supply/ccd:product/ccd:manufacturedProduct/ccd:manufacturedMaterial/ccd:code/ccd:translation[@codeSystem='2.16.840.1.113883.6.69']", "displayName");
                            medication.MedicationStrength = GetMedicationStrength(medication.DispensableDrugName);
                            medication.MedicationStrengthUnit = GetMedicationStrengthUnit(medication.DispensableDrugName);
                            medication.Frequency = "";

                            // Status should be based on whether or not EndDate supplied
                            if (medication.EndDate == DateTime.MinValue)
                            {
                                medication.Status = "active";  // End date not supplied so active
                            }
                            else
                            {
                                medication.Status = "completed";  // End date supplied (not equal to min value) so completed
                            }

                            // If active then include
                            if (medication.Status == "active")
                            {
                                Medications.Add(medication);  // Add the medication
                            }
                        }
                    }
                }
            } // LoadMedications

            /* LoadProblems
             * 
             * Purpose:  Loads probles from current document.
             * 
             * Problems (11450-4) aka problems-pmh  - See patient w/ last name Wineka
             * 
             * CSV: UniqueId, PatientId, ProblemName, Status, StartDate, ICD9, Comment
             * 
             * XML snippet:
             * 
             * /act
             *    /statusCode.code                                          - Status
             *    /effectiveTime/low.value (YYYYMMDD)                       - StartDate
			 *	  /effectiveTime/high.value (YYYYMMDD)
     		 *	  /entryRelationship
  			 *		   /sequenceNumber.value (listing order, may be of use)
             *         /observation
			 *		      /value.displayName                                - ProblemName
			 *		      /value.code (SNOMED CT code may map to ICD.9)
             */
            private void LoadProblems()
            {
                Problems = new List<Problem>();
                XmlNode problemsNode = SelectSingleNodeParent(xmlDoc, "/ccd:ClinicalDocument/ccd:component/ccd:structuredBody/ccd:component/ccd:section/ccd:code[@code='11450-4']", nsmgr);
                if (problemsNode != null)
                {
                    string sectionTitle = NodeAttributeStringValue(problemsNode, nsmgr, "./ccd:title", "[InnerText]");
                    if (sectionTitle == PROBLEMS_SECTION_TITLE)
                    {
                        XmlNodeList entryNodes = problemsNode.SelectNodes("./ccd:entry", nsmgr);
                        foreach (XmlNode entryNode in entryNodes)
                        {
                            Problem problem = new Problem();

                            problem.ProblemName = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:act/ccd:entryRelationship/ccd:observation/ccd:value", "displayName");
                            problem.StartDate = NodeAttributeDateTimeValue(entryNode, nsmgr, "./ccd:act/ccd:effectiveTime/ccd:low", "value");
                            problem.Status = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:act/ccd:statusCode", "code");
                            problem.ICD9 = "";
                            problem.Comment = "";

                            // Only load problem if not empty
                            if (!problem.isEmpty())
                            {
                                Problems.Add(problem);  // Add the problem
                            }
                        }
                    }
                }
            } // LoadProblems

            /*
             * LoadProcedures
             * 
             * Purpose:  Loads procedures from current document file.
             * 
             * Procedures (47519-4)  - aka surgeries
             * 
             * CSV: UniqueId, PatientId, SurgeryName, StartDate, ICD9, Comment
             * 
             * XML snippet:
             *
             *	 /procedure
             *	 	/effectiveTime/low.value = YYYYMMDD of      - StartDate
             *		/code.code = CPT Code (map to ICD.9?)
             *		/code.displayName                           - SurgeryName
             *
             */
            private void LoadProcedures()
            {
                Procedures = new List<Procedure>();
                XmlNode proceduresNode = SelectSingleNodeParent(xmlDoc, "/ccd:ClinicalDocument/ccd:component/ccd:structuredBody/ccd:component/ccd:section/ccd:code[@code='47519-4']", nsmgr);
                if (proceduresNode != null)
                {
                    string sectionTitle = NodeAttributeStringValue(proceduresNode, nsmgr, "./ccd:title", "[InnerText]");
                    if (sectionTitle == PROCEDURES_SECTION_TITLE)
                    {
                        XmlNodeList entryNodes = proceduresNode.SelectNodes("./ccd:entry", nsmgr);
                        foreach (XmlNode entryNode in entryNodes)
                        {
                            Procedure procedure = new Procedure();

                            procedure.SurgeryName = NodeAttributeStringValue(entryNode, nsmgr, "./ccd:procedure/ccd:code", "displayName");
                            procedure.StartDate = NodeAttributeDateTimeValue(entryNode, nsmgr, "./ccd:procedure/ccd:effectiveTime/ccd:low", "value");
                            procedure.ICD9 = "";
                            procedure.Comment = "";

                            Procedures.Add(procedure);  // Add the procedure
                        }
                    }
                }
            } // LoadProblems

            /* LoadVitals
             * 
             * Purpose:  Loads vital records for current document.
             * 
             *                 
             * Vital Signs (8716-3)
             *
             * CSV:  UniqueId, PatientId, Date, WeightInGrams, HeightInCm, TemperatureInCelsius, BP-Sys,
             *       BP-Dia, Pulse, RespRate, O2
             *       
             * XML snippet:
             * 	
			 *		/organizer
			 *			/component
			 *				/observation
             *                  /effectiveTime.value        - StartDate
             *                  /value.value                - value of vital sign (see code.code at same level for vital sign type)
			 *				  	/code[@code=“3141-9”]       - Weight
			 *				  	/code[@code=“8302-2”]       - Height
			 *				  	/code[@code=“8310-5”]       - Temp  (Not found)
			 *				  	/code[@code=“8480-6”]       - BP-Sys
			 *				  	/code[@code=“8462-4”]       - BP-Dia
			 *				  	/code[@code=“8867-4”]       - Pulse (Not found)
  			 *					/code[@code=“9279-1”]       - RespRate (Not found)
             */
            private void LoadVitals()
            {
                VitalSigns = new List<VitalsReading>();
                XmlNode vitalsNode = SelectSingleNodeParent(xmlDoc, "/ccd:ClinicalDocument/ccd:component/ccd:structuredBody/ccd:component/ccd:section/ccd:code[@code='8716-3']", nsmgr);
                if (vitalsNode != null)
                {
                    string sectionTitle = NodeAttributeStringValue(vitalsNode, nsmgr, "./ccd:title", "[InnerText]");
                    if (sectionTitle == VITALS_SECTION_TITLE)
                    {
                        XmlNodeList entryNodes = vitalsNode.SelectNodes("./ccd:entry", nsmgr);
                        foreach (XmlNode entryNode in entryNodes)
                        {
                            VitalsReading vitalsReading = new VitalsReading();

                            vitalsReading.StartDate = NodeAttributeDateTimeValue(entryNode, nsmgr, "./ccd:organizer/ccd:effectiveTime", "value");
                            vitalsReading.WeightInLbs = NodeAttributeFloatValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='3141-9']"), nsmgr, "./ccd:value", "value");
                            vitalsReading.HeightInInches = NodeAttributeFloatValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='8302-2']"), nsmgr, "./ccd:value", "value");
                            vitalsReading.TemperatureInFahrenheit = NodeAttributeFloatValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='8310-5']"), nsmgr, "./ccd:value", "value");
                            vitalsReading.BPSys = NodeAttributeIntValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='8480-6']"), nsmgr, "./ccd:value", "value");
                            vitalsReading.BPDia = NodeAttributeIntValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='8462-4']"), nsmgr, "./ccd:value", "value");
                            vitalsReading.Pulse = NodeAttributeIntValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='8867-4']"), nsmgr, "./ccd:value", "value");
                            vitalsReading.RespRate = NodeAttributeIntValue(NodeSelectSingleNodeParent(entryNode, nsmgr, "./ccd:organizer/ccd:component/ccd:observation/ccd:code[@code='9279-1']"), nsmgr, "./ccd:value", "value");

                            VitalSigns.Add(vitalsReading);  // Add the vitals reading
                        }
                    }
                }
            } // LoadVitals


            // SelectSingleNodeParent
            //
            // Purpose:  Much like XmlDoc.SelectSingleNode this function finds a document node given a path
            //           and then backs up one level to return the parent. If node specified by path not 
            //           found or parent is null then null is returned.
            private static XmlNode SelectSingleNodeParent(XmlDocument xmlDoc, string xpath, XmlNamespaceManager nsmgr)
            {
                XmlNode node = xmlDoc.SelectSingleNode(xpath, nsmgr);
                if (node != null)
                {
                    return node.ParentNode;
                }
                else
                {
                    return null;
                }

            }

            // NodeSelectSingleNodeParent
            //
            // Purpose:  Given an XmlNode, namspace and subNodePath returns the found node else returns null
            private static XmlNode NodeSelectSingleNodeParent(XmlNode xmlNode, XmlNamespaceManager nsmgr, String subNodePath)
            {
                XmlNode xmlSubNode = xmlNode.SelectSingleNode(subNodePath, nsmgr);
                if (xmlSubNode != null)
                {
                    return xmlSubNode.ParentNode;
                }
                else
                    return null;
            }

            // NodeAttributeStringValue
            //
            // Purpose:  Given an XmlNode, a namespace, a path to subnode and an attribute of said subnode this
            //           member function returns the string representation of said attribute.  Also supports an
            //           attribute of "[InnerText]" to get the InnerText of the node in lieu of an attribute's
            //           value.
            private static String NodeAttributeStringValue(XmlNode xmlNode, XmlNamespaceManager nsmgr, String subNodePath, String attribute)
            {
                if (xmlNode == null) return "";
                XmlNode xmlSubNode = xmlNode.SelectSingleNode(subNodePath, nsmgr);
                if (xmlSubNode != null)
                {
                    return AttributeStringValue(xmlSubNode, attribute);
                }
                else
                    return "";
            }

            // NodeAttributeFloatValue
            //
            // Purpose:  Given an XmlNode, a namespace, a path to subnode and an attribute of said subnode this
            //           member function returns the float representation of said attribute.  Also supports an
            //           attribute of "[InnerText]" to get the InnerText of the node in lieu of an attribute's
            //           value.
            private static float NodeAttributeFloatValue(XmlNode xmlNode, XmlNamespaceManager nsmgr, String subNodePath, String attribute)
            {
                if (xmlNode == null) return 0.0F;
                XmlNode xmlSubNode = xmlNode.SelectSingleNode(subNodePath, nsmgr);
                if (xmlSubNode != null)
                {
                    float f;
                    try
                    {
                        f = float.Parse(AttributeStringValue(xmlSubNode, attribute));
                    }
                    catch
                    {
                        f=0.0F;
                    }
                    return f; //float.Parse(AttributeStringValue(xmlSubNode, attribute));
                }
                else
                    return 0.0F;
            }

            // NodeAttributeIntValue
            //
            // Purpose:  Given an XmlNode, a namespace, a path to subnode and an attribute of said subnode this
            //           member function returns the integer representation of said attribute.  Also supports an
            //           attribute of "[InnerText]" to get the InnerText of the node in lieu of an attribute's
            //           value.
            private static int NodeAttributeIntValue(XmlNode xmlNode, XmlNamespaceManager nsmgr, String subNodePath, String attribute)
            {
                if (xmlNode == null) return 0;
                XmlNode xmlSubNode = xmlNode.SelectSingleNode(subNodePath, nsmgr);
                if (xmlSubNode != null)
                {
                    int i;
                    try
                    {
                        i = int.Parse(AttributeStringValue(xmlSubNode, attribute));
                    }
                    catch
                    {
                        i = 0;
                    }
                    return i;
                }
                else
                    return 0;
            }

            // NodeAttributeDateTimeValue
            //
            // Purpose:  Given an XmlNode, a namespace, a path to subnode and an attribute of said subnode this
            //           member function returns the string representation of said attribute.  Also supports an
            //           attribute of "[InnerText]" to get the InnerText of the node in lieu of an attribute's
            //           value.
            private static DateTime NodeAttributeDateTimeValue(XmlNode xmlNode, XmlNamespaceManager nsmgr, String subNodePath, String attribute)
            {
                XmlNode xmlSubNode = xmlNode.SelectSingleNode(subNodePath, nsmgr);
                if (xmlSubNode != null)
                {
                    DateTime dt;
                    try
                    {
                        dt = (DateTime)AttributeDateTimeValue(xmlSubNode, attribute);
                    }
                    catch
                    {
                        dt = DateTime.MinValue;
                    }
                    return dt;
                }
                else
                    return DateTime.MinValue;
            }

            // AttributeStringValue
            //
            // Purpose: Given an XmlNode and an attribute name returns the string value
            //          of the attribute.  If the attribute does not exist then the empty 
            //          string is returned.  Also supports attribute of "[InnerText]" keyword 
            //          in which case the inner text of the element is returned.
            private static String AttributeStringValue(XmlNode element, String attribute)
            {
                if (element != null)
                {
                    if (attribute == "[InnerText]")
                    {
                        return element.InnerText;
                    }
                    else
                    {
                        if (element.Attributes[attribute] != null)
                        {
                            return element.Attributes[attribute].Value.ToString();
                        }
                        else
                        {
                            return "";
                        }
                    }
                }
                else
                {
                    return "";
                }
            }

            // AttributeDateTimeValue
            //
            // Purpose: Given an XmlNode and an attribute name returns the DateTime value
            //          of the attribute.  If the attribute does not exist then DateTime.MinValue
            //          is returned.  Also supports attribute of "[InnerText]" keyword, in 
            //          which case the inner text of the element converted to DateTime is returned.
            private static DateTime AttributeDateTimeValue(XmlNode element, String attribute)
            {
                if (element != null)
                {
                    if (attribute == "[InnerText]")
                    {
                        return DateTime.ParseExact(element.InnerText.Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        if (element.Attributes[attribute] != null)
                        {
                            return DateTime.ParseExact(element.Attributes[attribute].Value.ToString().Substring(0, 8), "yyyyMMdd", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            return DateTime.MinValue;
                        }
                    }
                }
                else
                {
                    return DateTime.MinValue;
                }
            }

            // SigNumToString
            //
            // Purpose: Given a string formated like "#sig-0" returns the SIG number as a string
            //          sans the pound sign (#).
            private static string SigNumToString(String sigNum)
            {
                if (sigNum == null || sigNum == "" || sigNum.Length < 5 || sigNum.Substring(0, 5) != "#sig-")
                {
                    return "";
                }
                else
                {
                    return (sigNum.Substring(1));

                }
            }

            // RemoveBraces
            //
            // Purpose:  Removes leading and trailing curly braces if present else returns supplied string
            private String RemoveBraces(String str)
            {
                int strLen = str.Length;
                if (strLen >= 2 && str.Substring(0,1)=="{" && str.Substring(strLen-1, 1)=="}")
                {
                    return str.Substring(1, strLen - 2);
                }
                else
                {
                    return str;
                }
            }

            // GetNumWords
            //
            // Purpose:
            private static int GetNumWords(String str)
            {
                String[] split = str.Split(' '); // Use empty space between word(s) as split character
                return split.Length;
            }

            // GetNthWord
            //
            // Purpose:  Given a string and a word index (1 is 1st), returns the word at that index.
            //           If there are fewer words in the string than the index specified then
            //           an empty string is returned.
            private static String GetNthWord(String str, int index)
            {
                String[] split = str.Split(' '); // Use empty space between word(s) as split character
                if (split != null && split.Length >= index)
                {
                    return (split[index - 1]);
                }
                else
                {
                    return "";
                }
            }

            // isStringNumeric
            //
            // Purpose returns true if string represents a numeric value
            private static bool isStringNumeric(String str)
            {
                int intVar;
                long longVar;
                float floatVar;

                return( int.TryParse(str, out intVar)  ||
                        long.TryParse(str, out longVar) ||
                        float.TryParse(str, out floatVar) );
            }

            // isStringValidUnit
            //
            // Purpose:  Returns true if the supplied word is a valid, recognized dispensable unit
            //           else returns false.
            private static bool isStringValidUnit(String word)
            {
                word = word.ToUpper(); // uppercase
                return (word == "MG" || word == "MEQ" || word == "UNT" ||
                         (word.Length >= 3 && word.Substring(0, 3) == "MG/") ||
                         (word.Length >= 4 && word.Substring(0, 4) == "MEQ/") ||
                         (word.Length >= 4 && word.Substring(0, 4) == "UNT/"));
            }

            // findNextNumericWord
            //
            // Purpose:  Returns the index (1-based) of the next occurence of a numeric string in the word, else returns 0.
            private static int findNextNumericWord(String str, int start)
            {
                int retVal = 0;

                for (int ndx = start; ndx <= GetNumWords(str) && retVal==0; ndx++)
                {
                    if (isStringNumeric(GetNthWord(str, ndx)))
                    {
                        retVal = ndx;
                    }
                }
                return retVal;
            }

            // findNextValidUnit
            //
            // Purpose:  Returns the index (1-based) of the next occurence of a valid unit string in the word, else returns 0.
            private static int findNextValidUnitWord(String str, int start)
            {
                int retVal = 0;

                for (int ndx = start; ndx <= GetNumWords(str) && retVal == 0; ndx++)
                {
                    if (isStringValidUnit(GetNthWord(str, ndx)))
                    {
                        retVal = ndx;
                    }
                }
                return retVal;
            }

            // GetMedicationStrength
            //
            // Purpose:  Given a dispensable drug name which include strength and units return the
            //           medication strength.  Format is:   Name Strength Units miscellaneous
            private static String GetMedicationStrength(String dispensableDrugName)
            {
                String strReturn = "";

                // Get the first word that is a valid unit and return the string prior...
                int ndx = findNextValidUnitWord(dispensableDrugName, 1);
                if (ndx > 0)
                {
                    if (ndx > 1)
                    {
                        strReturn = GetNthWord(dispensableDrugName, ndx - 1);
                    }
                }

                return strReturn;
            }

            // GetMedicationStrengthUnit
            //
            // Purpose:  Given a dispensable drug name which include strength and units return the
            //           medication strength units.   Format is:   Name Strength Units miscellaneous
            static String GetMedicationStrengthUnit(String dispensableDrugName)
            {
                String strReturn = "";

                // Get the first word that is a valid unit and return it...
                int ndx = findNextValidUnitWord(dispensableDrugName, 1);
                if (ndx > 0)
                {
                    strReturn = GetNthWord(dispensableDrugName, ndx);
                }

                return strReturn;
            }

        }; // CCDFile class
 

        /*** CCDTranslate class ***/
        public class CCDTranslate
        {
            public bool isInitialized = false;
            public bool isError = false;
            public String errorMessage = "Success";
            public String SourceDir;
            public String TargetDir;
            public String ErrorDir;
            public String FileSpec;
            public int AllergyId = 1;
            public int MedicationId = 1;
            public int ProblemId = 1;
            public int ProcedureId = 1;
            public int VitalsId = 1;
            public int FileNum = 1;
            public int FilesTranslated = 0;
            public int FilesInError = 0;
            public long elapsedTimeMs = 0L;
            public List<String> Files;    
            public bool bAppendFiles;

            // Constructor
            public CCDTranslate(string sourceDir, string targetDir, string errorDir, string fileSpec, bool bAppend) 
            {
                isInitialized = true;
                SourceDir = sourceDir;
                TargetDir = targetDir;
                ErrorDir = errorDir;
                FileSpec = fileSpec;
                bAppendFiles = bAppend;

                // Check if source, target and error directories exist 
                if (!Directory.Exists(SourceDir))
                {
                    errorMessage = "ERROR: Source directory " + SourceDir + " does not exist.";
                    isError = true;
                }
                else if (!Directory.Exists(TargetDir))
                {
                    errorMessage = "ERROR: Target directory " + TargetDir + " does not exist.";
                    isError = true;
                }
                else if (!Directory.Exists(ErrorDir))
                {
                    errorMessage = "ERROR: Error directory " + ErrorDir + " does not exist.";
                    isError = true;
                }

                // Get list of files to translate and save as Files list
                if (!isError)
                {
                    string[] files = Directory.GetFiles(SourceDir, FileSpec);
                    Files = new List<String>(files);
                    if (Files.Count == 0)
                    {
                        errorMessage = "ERROR: No files matching filespec of " + FileSpec + " found in " + sourceDir + " directory.";
                        isError = true;
                    }
                }
            }  // constructor

            // Member Functions

            // AppentTextToFile - Appends specified text to a file
            private static void AppendTextToFile(string fileSpec, string text)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileSpec, true))
                {
                    file.WriteLine(text);
                }
            }

            // GetNextId - Returns the next ID to be used after reading the last id in the file specified.
            private int GetNextId(String filepath)
            {
                if (File.Exists(filepath))
                {
                    int number;
                    string lastLine = File.ReadLines(filepath).Last();
                    string[] words = lastLine.Split(',');
                    if (words.Length >= 1 && int.TryParse(words[0], out number))
                        return int.Parse(words[0]);
                    else
                        return 1;
                }
                else
                {
                    return 1;  // File does not exist, return 1
                }
            }

            // PrepareOutputFile - preps an output file for batch translation
            private bool PrepareOutputFile(String filepath, String csvHeader, ref int id)
            {
                /*  Case 1: Not Appending, Not Exists - Add
                 *  Case 2: Not Appending, Exists - Delete then Add
                 *  Case 3: Appending, Not Exists - Add
                 *  Case 4: Appending, Exists - Do Nothing */

                // Case 2: If not appending (new) and file exists then remove it then write out CSV header
                if (!bAppendFiles && File.Exists(filepath))
                {
                    File.Delete(filepath);
                    AppendTextToFile(filepath, csvHeader);
                }
                // Case 1 and 3
                else if (!File.Exists(filepath))
                {
                    AppendTextToFile(filepath, csvHeader);
                }
                // Case 4 - don't delete, don't write out CSV, just update id base on last ID in the existing file
                else
                {
                    id = GetNextId(filepath);
                }

                // Check if file now exists 
                if (File.Exists(filepath))
                {
                    return false;  // Not an error
                }
                else 
                {
                    isError = true;
                    errorMessage = "ERROR: Could not write to file " + filepath + ".";
                    return true;
                }

            }

            // Translate - loops through all found files translating them
            public bool Translate()
            {
                // Declare and initialize stop watch to time execution
                var watch = System.Diagnostics.Stopwatch.StartNew();

                // Prepare the output files
                isError = PrepareOutputFile(ALLERGIES_FILE, ALLERGIES_CSV_HEADER, ref AllergyId);
                if (!isError) isError = PrepareOutputFile(MEDICATIONS_FILE, MEDICATIONS_CSV_HEADER, ref MedicationId);
                if (!isError) isError = PrepareOutputFile(PROBLEMS_FILE, PROBLEMS_CSV_HEADER, ref ProblemId);
                if (!isError) isError = PrepareOutputFile(PROCEDURES_FILE, PROCEDURES_CSV_HEADER, ref ProcedureId);
                if (!isError) isError = PrepareOutputFile(VITALS_FILE, VITALS_CSV_HEADER, ref VitalsId);
                if (!isError) isError = PrepareOutputFile(REPORT_FILE, REPORT_CSV_HEADER, ref FileNum);

                // Translate each file found
                foreach (String file in Files)
                {
                    CCDFile ccdFile = new CCDFile(file);
                    if (!ccdFile.isError)
                    {
                        // Write out allergy entries for current file
                        foreach (Allergy allergy in ccdFile.Allergies)
                        {
                            if (!allergy.isEmpty())
                            {
                                AppendTextToFile(ALLERGIES_FILE, allergy.toCSVString((AllergyId++).ToString(), ccdFile.PatientId, ccdFile.FileInfo));

                            }
                        }

                        // Write out medication entries for current file
                        foreach (Medication medication in ccdFile.Medications)
                        {
                            if (!medication.isEmpty())
                            {
                                AppendTextToFile(MEDICATIONS_FILE, medication.toCSVString((MedicationId++).ToString(), ccdFile.PatientId, ccdFile.FileInfo));

                            }
                        }

                        // Write out problems entries for current file
                        foreach (Problem problem in ccdFile.Problems)
                        {
                            if (!problem.isEmpty())
                            {
                                AppendTextToFile(PROBLEMS_FILE, problem.toCSVString((ProblemId++).ToString(), ccdFile.PatientId, ccdFile.FileInfo));

                            }
                        }

                        // Write out procedures entries for current file
                        foreach (Procedure procedure in ccdFile.Procedures)
                        {
                            if (!procedure.isEmpty())
                            {
                                AppendTextToFile(PROCEDURES_FILE, procedure.toCSVString((ProcedureId++).ToString(), ccdFile.PatientId, ccdFile.FileInfo));

                            }
                        }

                        // Write out vitals entries for current file
                        foreach (VitalsReading vitalsReading in ccdFile.VitalSigns)
                        {
                            if (!vitalsReading.isEmpty())
                            {
                                AppendTextToFile(VITALS_FILE, vitalsReading.toCSVString((VitalsId++).ToString(), ccdFile.PatientId, ccdFile.FileInfo));

                            }
                        }

                        // Move the file to processed
                        File.Move(file, TargetDir + "\\" + ccdFile.FileName);

                        // Output report line
                        AppendTextToFile(REPORT_FILE, FileNum++ + "," + ccdFile.FileName + ",\"" + ccdFile.errorMessage + "\"," + ccdFile.Allergies.Count + "," + ccdFile.Medications.Count + "," + ccdFile.Problems.Count + "," + ccdFile.Procedures.Count + "," + ccdFile.VitalSigns.Count );

                    }
                    else  // Error reading file
                    {
                        // Move the file to error directory
                        File.Move(file, ErrorDir + "\\" + ccdFile.FileName);

                        // Output report line
                        AppendTextToFile(REPORT_FILE, FileNum++ + "," + ccdFile.FileName + ",\"" + ccdFile.errorMessage + "\",0,0,0,0");

                        FilesInError++; // Increment number of files having errors in translation
                    }

                    FilesTranslated++;  // Increment number of files translated

                    // Output to std out number of files translated
                    Console.Write("\rFiles processed: {0}", FilesTranslated);

                }  // For each file

                // Return based on whether or not errors were encountered
                if (FilesInError > 0)
                {
                    errorMessage = "ERROR: " + FilesInError + " could not be translated. See report.csv for errors.";
                }

                // Stop the watch get elapsed MS
                watch.Stop();
                elapsedTimeMs = watch.ElapsedMilliseconds;

                // Return true of an error or there are files processed in error, or false
                return (isError || FilesInError > 0);
            }
        }; // CCDTranslate
        
        //*************************************************************************************
        //*** MAIN MISCELLANEOUS FUNCTIONS                                                  ***
        //*************************************************************************************

        // DateTimeToDateString
        //
        // Purpose:  Returns a DateTime as a string, including only the Date portion of the
        //           DateTime.  If the DateTime supplied is equal to DateTime.MinValue then
        //           and empty string is returned.
        static String DateTimeToDateString(DateTime datetime)
        {
            if (datetime != DateTime.MinValue)
            {
                return datetime.ToShortDateString();
            }
            else
            {
                return "";
            }
        }

        // QuoteCSVString
        //
        // Purpose:  If string includes a comma then incapsulate it with double quotes else return as-is.
        static String QuoteCSVString(String str)
        {
            if (str.Contains(','))
                return "\"" + str + "\"";
            else
                return str;
        }

        // FormatCSVDate
        //
        // Purpose:  If supplied DateTime is equal to MinValue then return "" else return date portion only
        //           as a string.
        static String FormatCSVDate(DateTime dt)
        {
            if (dt == DateTime.MinValue)
                return "";
            else
                return dt.ToShortDateString();
        }

        // CCDFileTest
        //
        // Purpose: Used to unit test the CCDFile class which is at the heart of reading and parsing a CCDFile.
        public static void CCDFileTest(string filepath)
        {
            CCDFile ccdFile = new CCDFile(filepath);
            if (!ccdFile.isError)
            {
                Console.WriteLine(ccdFile.FileName);
                Console.WriteLine("  Patient: " + ccdFile.PatientName);
                Console.WriteLine("  Allergies.Count: " + ccdFile.Allergies.Count);
                Console.WriteLine("  Medications.Count: " + ccdFile.Medications.Count);
                Console.WriteLine("  Problems.Count: " + ccdFile.Problems.Count);
                Console.WriteLine("  Procedures.Count: " + ccdFile.Procedures.Count);
                Console.WriteLine("  VitalSigns.Count: " + ccdFile.VitalSigns.Count);
            }
        }

        // SyntaxHelp
        //
        // Purpose:  Display syntax help
        public static void SyntaxHelp()
        {
            Console.Error.WriteLine("   ___   ___   _____  ___       _       ");
            Console.Error.WriteLine("  / __\\ / __\\ /   \\ \\/ / | __ _| |_ ___ ");
            Console.Error.WriteLine(" / /   / /   / /\\ /\\  /| |/ _` | __/ _ \\ ");
            Console.Error.WriteLine("/ /___/ /___/ /_// /  \\| | (_| | ||  __/ ");
            Console.Error.WriteLine("\\___/ \\____/___,' /_/\\_\\_|\\__,_|\\__\\___|\n");
            //Console.Error.WriteLine("\n" + moduleName + "\n");
            Console.Error.WriteLine("Translates Continuing Care Documents (CCD) in XML format to Comma-Separated-\nValue (CSV) files suitable for batch import.\n");
            Console.Error.WriteLine(moduleName + " [/s=sourcedir] [/t=targetdir] [/e=errordir] [/f=filespec] [/b]\n");
            Console.Error.WriteLine("Where:\n");
            Console.Error.WriteLine("  /s\tSource directory from which to find files to translate. Default is\t\t'pending' directory of current working directory.\n");
            Console.Error.WriteLine("  /t\tTarget directory to move processed files within. Default is 'processed'\t\tdirectory of current working directory.\n");
            Console.Error.WriteLine("  /e\tError directory to move files that cannot be processed to. Default is\t\t'error' directory of current working directory.\n");
            Console.Error.WriteLine("  /f\tFile specification used to match files to transform(e.g. *.xml).\t\tDefault is ccd*.xml.\n");
            Console.Error.WriteLine("  /a\tAppend translated files to existing import files if they exist or start\t\tnew ones. Default is to start new files.\n");
            Console.Error.WriteLine("\t   (c) Copyright TJRTech, Inc 2016 - All rights reserved.");
        }

        // GetArgValue
        //
        // Purpose: Given an argument of format /x=value, parse the argument and return value. 
        //          If not of the correct format return false.
        public static bool GetArgValue(string arg, ref string value)
        {
            bool bError = false;
            if (arg.Length>=4 && arg.Substring(2, 1) == "=")
            {
                value = arg.Substring(3);
            } 
            else
            {
                bError = true;
            }
            return bError;
        }

        // ParseArgs
        //
        // Purpose: Parse program arguments into program variables passed in.  Returns true if a syntax error, else false.
        public static bool ParseArgs(string[] args, ref string sourceDir, ref string targetDir, ref string errorDir, ref string fileSpec, ref bool bAppend)
        {
            bool bError = false;
            for (int x = 0; x < args.Length && bError==false; x++)
            {
                // All args start with / and must be at least two chars long (most are 4 or more)
                if (args[x].Length<2 || args[x].Substring(0, 1) != "/")
                {
                    bError = true;
                }

                // Safe to get first character after escape and then process
                string sParam = args[x].Substring(1, 1);
                switch (sParam)
                {
                    case "s":
                        bError = GetArgValue(args[x], ref sourceDir);
                        break;
                    case "t":
                        bError = GetArgValue(args[x], ref targetDir);
                        break;
                    case "e":
                        bError = GetArgValue(args[x], ref errorDir);
                        break;
                    case "f":
                        bError = GetArgValue(args[x], ref fileSpec);
                        break;
                    case "a":
                        if (args[x].Length == 2)
                        {
                            bAppend = true;
                        }
                        else
                        {
                            bError = true; // Unexpected characters after /b
                        }
                        break;
                    default:
                        bError = true;  // Unknown parameter
                        break;
                }
            }
            return bError;
        }

        // Main
        //
        // Purpose:  Main function of the program. Parses arguments and prepares and executes translation object.
        static int Main(string[] args)
        {
            string sourceDir = SOURCE_DIR;   // Source directory containing CCD files. Default is "source" in current directory.
            string targetDir = TARGET_DIR;   // Target directory processed CCD files will be moved to. Default is "processed" in current directory.
            string errorDir = ERROR_DIR;     // Error directory failing CCD files will be moved to. Default is "error" in current directory.
            string fileSpec = FILE_SPEC;     // The default file specification to process.
            bool bAppend = false;            // Append CCD file information to existing import files if they exist or start new ones
            CCDTranslate ccdTranslate;       // CCD file translator
 
            // Check if requesting syntax help
            if (args.Length >= 1 && (args[0] == "-?" || args[0] == "/?"))
            {
                SyntaxHelp();
                return 1;
            }

            // Parse the arguments
            if (ParseArgs(args, ref sourceDir, ref targetDir, ref errorDir, ref fileSpec, ref bAppend)) 
            {
                Console.Error.WriteLine("\nERROR:  Incorrect syntax.  Try: " + moduleName + " /?");
                return 1;  // Syntax error
            }

            // Output execution plan
            Console.WriteLine("\nExecuting CCDXlate with the following options:");
            Console.WriteLine("   sourceDir=" + sourceDir);
            Console.WriteLine("   targetDir=" + targetDir);
            Console.WriteLine("   errorDir=" + errorDir);
            Console.WriteLine("   fileSpecr=" + fileSpec);
            Console.WriteLine("   bAppend=" + bAppend);

            // Declare translation object and check for error
            ccdTranslate = new CCDTranslate(sourceDir, targetDir, errorDir, fileSpec, bAppend);
            if (!ccdTranslate.isError)  // If no error then translate all files
            {
                ccdTranslate.Translate();  // This method does all the translation, writing number of files processed repeatedly to stdout

                // Output timing and other results
                Console.WriteLine("\nCompleted!\n");
                Console.WriteLine("Results:");
                Console.WriteLine("   # Total Files: " + ccdTranslate.FilesTranslated);
                Console.WriteLine("   # Successful Files: " + (ccdTranslate.FilesTranslated - ccdTranslate.FilesInError));
                Console.WriteLine("   # Failed Files: " +  ccdTranslate.FilesInError);
                Console.WriteLine("   Execution Seconds: " + ccdTranslate.elapsedTimeMs / 1000.0);
            }
            else
            {
                Console.Error.WriteLine("\n\n" + ccdTranslate.errorMessage);
                return 1;  // Translation error
            }
            return 0;
        }  // main

    }   // class
}   // namespace
