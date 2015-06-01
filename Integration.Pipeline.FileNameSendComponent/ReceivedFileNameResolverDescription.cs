using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace Integration.Pipeline.FileNameSendComponent
{
    [AttributeUsage(AttributeTargets.All, Inherited = true, AllowMultiple = false)]
    public sealed class ReceivedFileNameResolverPropertyNameAttribute : Attribute
    {
        private string propertyName;

        public ReceivedFileNameResolverPropertyNameAttribute(string propertyName)
        {
            this.propertyName = propertyName;
        }

        public string PropertyName
        {
            get
            {
                return propertyName;
            }
        }
    }

    [AttributeUsage(AttributeTargets.All, Inherited = true, AllowMultiple = false)]
    public sealed class ReceivedFileNameResolverPropertyDescriptionAttribute : Attribute
    {
        private string propertyDescription;

        public ReceivedFileNameResolverPropertyDescriptionAttribute(string propertyDescription)
        {
            this.propertyDescription = propertyDescription;
        }

        public string PropertyDescription
        {
            get
            {
                return propertyDescription;
            }
        }
    }


    #region CsvToExcelConverterPropertyDescriptor

    public class ReceivedFileNameResolverDescriptor : PropertyDescriptor
    {
        private PropertyDescriptor descriptor;
        private ResourceManager resManager;

        public ResourceManager ResourceManager
        {
            get
            {
                return resManager;
            }
        }

        public ReceivedFileNameResolverDescriptor(PropertyDescriptor descriptor, ResourceManager resourceManager) : base(descriptor)
        {
            this.descriptor = descriptor;
            this.resManager = resourceManager;
        }

        public override AttributeCollection Attributes
        {
            get
            {
                return descriptor.Attributes;
            }
        }

        public override object GetEditor(Type editorBaseType)
        {
            return descriptor.GetEditor(editorBaseType);
        }

        public override string Category
        {
            get
            {
                return descriptor.Category;
            }
        }

        public override Type ComponentType
        {
            get
            {
                return descriptor.ComponentType;
            }
        }

        public override TypeConverter Converter
        {
            get
            {
                return descriptor.Converter;
            }
        }

        public override string Description
        {
            get
            {
                AttributeCollection attributes = descriptor.Attributes;

                ReceivedFileNameResolverPropertyDescriptionAttribute descriptionAttribute =
                    (ReceivedFileNameResolverPropertyDescriptionAttribute)attributes[typeof(ReceivedFileNameResolverPropertyDescriptionAttribute)];

                if (descriptionAttribute == null)
                    return descriptor.Description;

                string strId = descriptionAttribute.PropertyDescription;
                if (resManager == null)
                    return strId;

                return resManager.GetString(strId);
            }
        }

        public override bool DesignTimeOnly
        {
            get
            {
                return descriptor.DesignTimeOnly;
            }
        }

        public override bool IsBrowsable
        {
            get
            {
                return descriptor.IsBrowsable;
            }
        }

        public override bool IsLocalizable
        {
            get
            {
                return descriptor.IsLocalizable;
            }
        }

        public override bool IsReadOnly
        {
            get
            {
                return descriptor.IsReadOnly;
            }
        }

        public override Type PropertyType
        {
            get
            {
                return descriptor.PropertyType;
            }
        }

        public override bool ShouldSerializeValue(object o)
        {
            return descriptor.ShouldSerializeValue(o);
        }

        public override void AddValueChanged(object o, EventHandler handler)
        {
            descriptor.AddValueChanged(o, handler);
        }

        public override string DisplayName
        {
            get
            {
                AttributeCollection attributes = descriptor.Attributes;

                ReceivedFileNameResolverPropertyNameAttribute nameAttribute =
                    (ReceivedFileNameResolverPropertyNameAttribute)attributes[typeof(ReceivedFileNameResolverPropertyNameAttribute)];

                if (nameAttribute == null)
                    return descriptor.DisplayName;

                string strId = nameAttribute.PropertyName;
                if (resManager == null)
                    return strId;

                return resManager.GetString(strId);
            }
        }

        public override string Name
        {
            get
            {
                return descriptor.Name;
            }
        }

        public override Object GetValue(object o)
        {
            return descriptor.GetValue(o);
        }

        public override void ResetValue(object o)
        {
            descriptor.ResetValue(o);
        }

        public override bool CanResetValue(object o)
        {
            return descriptor.CanResetValue(o);
        }

        public override void SetValue(object obj1, object obj2)
        {
            descriptor.SetValue(obj1, obj2);
        }
    }

    #endregion

    #region BaseCustomTypeDescriptor
    
    public class BaseCustomTypeDescriptor : ICustomTypeDescriptor
    {
        private ResourceManager resourceManager;

        public BaseCustomTypeDescriptor(ResourceManager resourceManager)
        {
            this.resourceManager = resourceManager;
        }

        public AttributeCollection GetAttributes()
        {
            return new AttributeCollection(null);
        }

        public virtual string GetClassName()
        {
            return null;
        }

        public virtual string GetComponentName()
        {
            return null;
        }

        public TypeConverter GetConverter()
        {
            return null;
        }

        public EventDescriptor GetDefaultEvent()
        {
            return null;
        }

        public PropertyDescriptor GetDefaultProperty()
        {
            return null;
        }

        public object GetEditor(Type editorBaseType)
        {
            return null;
        }

        public EventDescriptorCollection GetEvents()
        {
            return EventDescriptorCollection.Empty;
        }

        public EventDescriptorCollection GetEvents(Attribute[] filter)
        {
            return EventDescriptorCollection.Empty;
        }

        public virtual PropertyDescriptorCollection GetProperties()
        {
            PropertyDescriptorCollection srcProperties = TypeDescriptor.GetProperties(this.GetType());
            ReceivedFileNameResolverDescriptor[] bteProperties = new ReceivedFileNameResolverDescriptor[srcProperties.Count];

            int i = 0;
            foreach (PropertyDescriptor srcDescriptor in srcProperties)
            {
                ReceivedFileNameResolverDescriptor destDescriptor = new ReceivedFileNameResolverDescriptor(srcDescriptor, resourceManager);
                AttributeCollection attributes = srcDescriptor.Attributes;
                bteProperties[i++] = destDescriptor;
            }
            PropertyDescriptorCollection destProperties = new PropertyDescriptorCollection(bteProperties);
            return destProperties;
        }

        public virtual PropertyDescriptorCollection GetProperties(Attribute[] filter)
        {
            return this.GetProperties();
        }

        public object GetPropertyOwner(PropertyDescriptor pd)
        {
            return this;
        }
    }

    #endregion
}
