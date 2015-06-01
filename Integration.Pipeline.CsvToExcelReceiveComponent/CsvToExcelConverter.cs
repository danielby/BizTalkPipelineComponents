using System;
using System.Resources;
using System.Drawing;
using System.Collections;
using System.Reflection;
using System.ComponentModel;
using System.Text;
using System.IO;
using Microsoft.BizTalk.Message.Interop;
using Microsoft.BizTalk.Component.Interop;
using GemBox.Spreadsheet;

namespace Integration.Pipeline.CsvToExcelReceiveComponent
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [ComponentCategory(CategoryTypes.CATID_Any)]
    [ComponentCategory(CategoryTypes.CATID_Validate)]
    [System.Runtime.InteropServices.Guid("48BEC85A-20EE-40ad-BFD0-319B59A0DDBC")]
    public class CsvToExcelConverter : BaseCustomTypeDescriptor, IBaseComponent, Microsoft.BizTalk.Component.Interop.IComponent, Microsoft.BizTalk.Component.Interop.IPersistPropertyBag, IComponentUI
    {
        static ResourceManager resManager = new ResourceManager("Integration.Pipeline.CsvToExcelReceiveComponent.CsvToExcelConverter", Assembly.GetExecutingAssembly());

        public CsvToExcelConverter() : base(resManager)
        {
            SpreadsheetInfo.SetLicense("<LicenseKey>");
        }

        private string separator;

        [CsvToExcelConverterPropertyName("PropSeparator"), CsvToExcelConverterPropertyDescription("DescrSeparator")]
        public string Separator
        {
            get { return separator; }
            set { separator = value; }
        }

        #region IComponent

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            var bodyPart = pInMsg.BodyPart;

            if (bodyPart != null)
            {
                var originalStream = bodyPart.GetOriginalDataStream();
                if (originalStream != null)
                {
                    ExcelFile excelFile = new ExcelFile();
                    excelFile.LoadCsv(originalStream, separator.ToCharArray()[0]);
                    var excelStream = new MemoryStream();
                    excelFile.SaveXlsx(excelStream);
                    excelStream.Seek(0, SeekOrigin.Begin);

                    bodyPart.Data = excelStream;
                    pContext.ResourceTracker.AddResource(excelStream);

                    if (pInMsg.Context != null)
                    {
                        var receivedFilename = pInMsg.Context.Read("ReceivedFileName", "http://schemas.microsoft.com/BizTalk/2003/file-properties");
                        pInMsg.Context.Write("ReceivedFileName", "http://schemas.microsoft.com/BizTalk/2003/file-properties", receivedFilename + ".xlsx");
                    }

                }
            }

            return pInMsg;     
        }

        #endregion

        #region IBaseComponent

        public string Description
        {
            get { return "Converts a csv format to an excel format."; }
        }

        public string Name
        {
            get { return "CsvToExcelConverter"; }
        }

        public string Version
        {
            get { return "1.0.0.0"; }
        }

        #endregion

        #region IPersistPropertyBag

        public void GetClassID(out Guid classID)
        {
            classID = new System.Guid("48BEC85A-20EE-40ad-BFD0-319B59A0DDBC");
        }

        public void InitNew()
        {
        }

        public void Load(IPropertyBag propertyBag, int errorLog)
        {
            string val = (string)ReadPropertyBag(propertyBag, "Separator");
            if (val != null) separator = val;
        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            object val = (object)separator;
            WritePropertyBag(propertyBag, "Separator", val);
        }

        private static object ReadPropertyBag(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, string propName)
        {
            object val = null;
            try
            {
                pb.Read(propName, out val, 0);
            }

            catch (ArgumentException)
            {
                return val;
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
            return val;
        }

        private static void WritePropertyBag(Microsoft.BizTalk.Component.Interop.IPropertyBag pb, string propName, object val)
        {
            try
            {
                pb.Write(propName, ref val);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }

        #endregion

        #region IComponentUI

        [Browsable(false)]
        public IntPtr Icon
        {
            get { return new System.IntPtr(); }
        }

        public IEnumerator Validate(object projectSystem)
        {
            IEnumerator enumerator = null;
            ArrayList strList = new ArrayList();

            // Validate prepend data property
            if ((separator != null) &&
            (Separator.Length > 1))
            {
                strList.Add(resManager.GetString("ErrorSeparatorTooLong"));
            }

            if (strList.Count > 0)
            {
                enumerator = strList.GetEnumerator();
            }

            return enumerator;
        }

        #endregion
    }
}
