using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CSSPLabSheetParserDLL.Services;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPFCFormWriterDLL.Services
{
    public partial class CSSPFCFormWriter
    {
        #region Variables
        #endregion Variables

        #region Properties
        public LanguageEnum LanguageRequest { get; set; }
        public CSSPLabSheetParser csspLabSheetParser { get; set; }
        public LabSheetA1Sheet labSheetA1Sheet { get; set; }
        public string FullLabSheet { get; set; }
        #endregion Properties

        #region Constructors
        public CSSPFCFormWriter(LanguageEnum LanguageRequest, string FullLabSheet)
        {
            this.LanguageRequest = LanguageRequest;
            //CultureInfo cultureInfo = new CultureInfo(LanguageRequest + "-CA");
            //Thread.CurrentThread.CurrentCulture = cultureInfo;
            //Thread.CurrentThread.CurrentUICulture = cultureInfo;

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-CA");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-CA");

            csspLabSheetParser = new CSSPLabSheetParser();
            this.FullLabSheet = FullLabSheet;
        }
        #endregion Constructors

        #region Functions public
        public string CreateFCForm(string FileName)
        {
            labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullLabSheet);
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.Error))
                return labSheetA1Sheet.Error;

            ReorderLabSheetA1Sheet();
            if (!string.IsNullOrWhiteSpace(labSheetA1Sheet.Error))
                return labSheetA1Sheet.Error;

            return CreatePackage(FileName);
        }
        #endregion Functions public

        public string ReorderLabSheetA1Sheet()
        {
            LabSheetA1Measurement LabSheetA1MeasurementDailyDuplicate = labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.SampleType == CSSPEnumsDLL.Enums.SampleTypeEnum.DailyDuplicate).FirstOrDefault();
            LabSheetA1Measurement LabSheetA1MeasurementIntertechDuplicate = labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.SampleType == CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechDuplicate).FirstOrDefault();
            LabSheetA1Measurement LabSheetA1MeasurementIntertechRead = labSheetA1Sheet.LabSheetA1MeasurementList.Where(c => c.SampleType == CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechRead).FirstOrDefault();

            labSheetA1Sheet.LabSheetA1MeasurementList = (from c in labSheetA1Sheet.LabSheetA1MeasurementList
                                                         where c.SampleType != CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechDuplicate
                                                         && c.SampleType != CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechRead
                                                         && c.SampleType != CSSPEnumsDLL.Enums.SampleTypeEnum.DailyDuplicate
                                                         select c).ToList();

            if (LabSheetA1MeasurementIntertechDuplicate != null)
            {
                // Insert empty Line
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetEmptyLine());

                // Insert modified IntertechDuplicate
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetModifiedIntertechDuplicate(LabSheetA1MeasurementIntertechDuplicate));

                // Insert modified IntertechDuplicate RLog
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetRLogLineOfIntertechDuplicate());

                // Insert modified IntertechDuplicate Pricision Criteria
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetPrecisionCriteriaLineOfIntertechDuplicate());
            }

            if (LabSheetA1MeasurementIntertechRead != null)
            {
                // Insert empty Line
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetEmptyLine());

                // Insert modified IntertechRead
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetModifiedIntertechRead(LabSheetA1MeasurementIntertechRead));

                // Insert modified IntertechRead
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(GetAcceptableLineOfIntertechRead());
            }

            if (LabSheetA1MeasurementDailyDuplicate != null)
            {
                labSheetA1Sheet.LabSheetA1MeasurementList.Add(LabSheetA1MeasurementDailyDuplicate);
            }

            return "";
        }

        public LabSheetA1Measurement GetEmptyLine()
        {
            LabSheetA1Measurement labSheetA1MeasurementEmptyLine = new LabSheetA1Measurement()
            {
                Site = null,
                Time = null,
                MPN = null,
                Tube10 = null,
                Tube1_0 = null,
                Tube0_1 = null,
                Salinity = null,
                Temperature = null,
                ProcessedBy = null,
                SampleType = CSSPEnumsDLL.Enums.SampleTypeEnum.Error,
                SiteComment = null,
                TVItemID = 0,
            };

            return labSheetA1MeasurementEmptyLine;
        }
        public LabSheetA1Measurement GetModifiedIntertechDuplicate(LabSheetA1Measurement LabSheetA1MeasurementIntertechDuplicate)
        {
            LabSheetA1Measurement labSheetA1MeasurementModifiedIntertechDuplicate = new LabSheetA1Measurement()
            {
                Site = LabSheetA1MeasurementIntertechDuplicate.Site,
                Time = null,
                MPN = LabSheetA1MeasurementIntertechDuplicate.MPN,
                Tube10 = LabSheetA1MeasurementIntertechDuplicate.Tube10,
                Tube1_0 = LabSheetA1MeasurementIntertechDuplicate.Tube1_0,
                Tube0_1 = LabSheetA1MeasurementIntertechDuplicate.Tube0_1,
                Salinity = null,
                Temperature = null,
                ProcessedBy = LabSheetA1MeasurementIntertechDuplicate.ProcessedBy,
                SampleType = CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechDuplicate,
                SiteComment = null,
                TVItemID = -1, // when SampleType = IntertechDuplicate and TVItemID = -1 then replace time text with "Int. D" 
            };

            return labSheetA1MeasurementModifiedIntertechDuplicate;
        }
        public LabSheetA1Measurement GetRLogLineOfIntertechDuplicate()
        {
            if (labSheetA1Sheet.IntertechDuplicateRLog == "Not calculated")
            {
                labSheetA1Sheet.IntertechDuplicateRLog = "0.0";
            }
            LabSheetA1Measurement labSheetA1MeasurementModifiedIntertechDuplicate = new LabSheetA1Measurement()
            {
                Site = null,
                Time = null,
                MPN = (int)(float.Parse(labSheetA1Sheet.IntertechDuplicateRLog) * 1000000),
                Tube10 = null,
                Tube1_0 = null,
                Tube0_1 = null,
                Salinity = null,
                Temperature = null,
                ProcessedBy = null,
                SampleType = CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechDuplicate,
                SiteComment = null,
                TVItemID = -2, // when SampleType = IntertechDuplicate and TVItemID = -2 then replace time text with "RLog" and devide MPN by 1000000
            };

            return labSheetA1MeasurementModifiedIntertechDuplicate;
        }
        public LabSheetA1Measurement GetPrecisionCriteriaLineOfIntertechDuplicate()
        {
            LabSheetA1Measurement labSheetA1MeasurementModifiedIntertechDuplicate = new LabSheetA1Measurement()
            {
                Site = null,
                Time = null,
                MPN = (int)(float.Parse(labSheetA1Sheet.IntertechDuplicatePrecisionCriteria) * 1000000),
                Tube10 = null,
                Tube1_0 = null,
                Tube0_1 = null,
                Salinity = null,
                Temperature = null,
                ProcessedBy = (labSheetA1Sheet.IntertechDuplicateAcceptableOrUnacceptable == "Acceptable" ? "OK" : "Not OK"),
                SampleType = CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechDuplicate,
                SiteComment = null,
                TVItemID = -3, // when SampleType = IntertechDuplicate and TVItemID = -3 then replace time text with "PCrit" and devide MPN by 1000000
            };

            return labSheetA1MeasurementModifiedIntertechDuplicate;
        }
        public LabSheetA1Measurement GetModifiedIntertechRead(LabSheetA1Measurement LabSheetA1MeasurementIntertechRead)
        {
            LabSheetA1Measurement labSheetA1MeasurementModifiedIntertechDuplicate = new LabSheetA1Measurement()
            {
                Site = LabSheetA1MeasurementIntertechRead.Site,
                Time = null,
                MPN = LabSheetA1MeasurementIntertechRead.MPN,
                Tube10 = LabSheetA1MeasurementIntertechRead.Tube10,
                Tube1_0 = LabSheetA1MeasurementIntertechRead.Tube1_0,
                Tube0_1 = LabSheetA1MeasurementIntertechRead.Tube0_1,
                Salinity = null,
                Temperature = null,
                ProcessedBy = LabSheetA1MeasurementIntertechRead.ProcessedBy,
                SampleType = CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechRead,
                SiteComment = null,
                TVItemID = -4, // when SampleType = IntertechRead and TVItemID = -4 then replace time text with "Int. R" 
            };

            return labSheetA1MeasurementModifiedIntertechDuplicate;
        }
        public LabSheetA1Measurement GetAcceptableLineOfIntertechRead()
        {
            LabSheetA1Measurement labSheetA1MeasurementModifiedIntertechRead = new LabSheetA1Measurement()
            {
                Site = null,
                Time = null,
                MPN = null,
                Tube10 = null,
                Tube1_0 = null,
                Tube0_1 = null,
                Salinity = null,
                Temperature = null,
                ProcessedBy = (labSheetA1Sheet.IntertechReadAcceptableOrUnacceptable == "Acceptable" ? "OK" : "Not OK"),
                SampleType = CSSPEnumsDLL.Enums.SampleTypeEnum.IntertechRead,
                SiteComment = null,
                TVItemID = -5, // when SampleType = IntertechRead and TVItemID = -3 then nothing to replace
            };

            return labSheetA1MeasurementModifiedIntertechRead;
        }
    }
}
