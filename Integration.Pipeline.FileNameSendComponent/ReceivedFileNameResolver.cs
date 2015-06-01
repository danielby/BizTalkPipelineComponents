using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.BizTalk.Message.Interop;
using Microsoft.BizTalk.Component.Interop;
using System.Resources;
using System.Drawing;
using System.Reflection;
using System.ComponentModel;
using System.Collections;

namespace Integration.Pipeline.FileNameSendComponent
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [ComponentCategory(CategoryTypes.CATID_Any)]
    [ComponentCategory(CategoryTypes.CATID_Validate)]
    [System.Runtime.InteropServices.Guid("48BEC85A-20EE-40ad-BFD0-319B59A0DDBC")]
    public class ReceivedFileNameResolver : BaseCustomTypeDescriptor, IBaseComponent, Microsoft.BizTalk.Component.Interop.IComponent, Microsoft.BizTalk.Component.Interop.IPersistPropertyBag, IComponentUI
    {
        static ResourceManager resManager = new ResourceManager("Integration.Pipeline.FileNameSendComponent", Assembly.GetExecutingAssembly());

        private string staticFileName;
        [ReceivedFileNameResolverPropertyName("PropStatic"), ReceivedFileNameResolverPropertyDescription("DescrStatic")]
        public string StaticFileName
        {
            get { return staticFileName; }
            set { staticFileName = value; }
        }

        private string promotedProperyName;
        [ReceivedFileNameResolverPropertyName("PropPromoted"), ReceivedFileNameResolverPropertyDescription("DescrPromoted")]
        public string PromotedPropertyName
        {
            get { return promotedProperyName; }
            set { promotedProperyName = value; }
        }

        private string promotedProperyNamespace;
        [ReceivedFileNameResolverPropertyName("PropPromotedNamespace"), ReceivedFileNameResolverPropertyDescription("DescrPromotedNamespace")]
        public string PromotedProperyNamespace
        {
            get { return promotedProperyNamespace; }
            set { promotedProperyNamespace = value; }
        }

        public ReceivedFileNameResolver() : base(resManager) {}

        #region IComponent

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            if (pInMsg.Context != null)
            {
                if (!string.IsNullOrEmpty(staticFileName))
                {
                    pInMsg.Context.Write("ReceivedFileName", "http://schemas.microsoft.com/BizTalk/2003/file-properties", staticFileName);
                } else if (!string.IsNullOrEmpty(promotedProperyName))
                {
                    object obj = pInMsg.Context.Read(promotedProperyName, promotedProperyNamespace);
                    pInMsg.Context.Write("ReceivedFileName", "http://schemas.microsoft.com/BizTalk/2003/file-properties", obj);
                }    
            }

            return pInMsg;
        }

        #endregion

        #region IBaseComponent

        public string Description
        {
            get { return "Set FILE.ReceivedFileName statically or from a promoted property"; }
        }

        public string Name
        {
            get { return "ReceivedFileNameResolver"; }
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
            string val = (string)ReadPropertyBag(propertyBag, "StaticFileName");
            if (val != null) staticFileName = val;

            val = (string)ReadPropertyBag(propertyBag, "PromotedPropertyName");
            if (val != null) promotedProperyName = val;

            val = (string)ReadPropertyBag(propertyBag, "PromotedPropertyNamespace");
            if (val != null) promotedProperyNamespace = val;
        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            object val = (object)staticFileName;
            WritePropertyBag(propertyBag, "StaticFileName", val);

            val = (object)promotedProperyName;
            WritePropertyBag(propertyBag, "PromotedPropertyName", val);

            val = (object)promotedProperyNamespace;
            WritePropertyBag(propertyBag, "PromotedPropertyNamespace", val);
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

        public System.Collections.IEnumerator Validate(object projectSystem)
        {
            ArrayList strList = new ArrayList();

            // Validate prepend data property
            if (string.IsNullOrEmpty(staticFileName) && string.IsNullOrEmpty(promotedProperyName))
            {
                strList.Add(resManager.GetString("EmptyConfiguration"));
            }


            // Validate prepend data property
            if (!string.IsNullOrEmpty(staticFileName) && !string.IsNullOrEmpty(promotedProperyName))
            {
                strList.Add(resManager.GetString("TooManyConfigurations"));
            }

            return (strList.Count > 0 ? strList.GetEnumerator() : null);
        }

        #endregion
    }
}
