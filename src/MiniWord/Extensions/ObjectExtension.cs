using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;

namespace MiniSoftware.Extensions;

internal static class ObjectExtension
{
    internal static Dictionary<string, object> ToDictionary(this object value)
    {
        if (value == null)
        {
            return new Dictionary<string, object>();
        }
        else if (value is Dictionary<string, object> dicStr)
        {
            return dicStr;
        }
        else if (value is ExpandoObject)
        {
            return new Dictionary<string, object>(value as ExpandoObject);
        }

        if (IsStrongTypeEnumerable(value))
        {
            throw new Exception("The parameter cannot be a collection type");
        }

        Dictionary<string, object> result = new Dictionary<string, object>();
        PropertyDescriptorCollection props = TypeDescriptor.GetProperties(value);
        foreach (PropertyDescriptor prop in props)
        {
            object propValue = prop.GetValue(value);

            if (IsStrongTypeEnumerable(propValue))
            {
                bool isValueList = false;
                List<Dictionary<string, object>> sx = new List<Dictionary<string, object>>();
                foreach (object val1item in (IEnumerable)propValue)
                {
                    if (val1item == null)
                    {
                        sx.Add(new Dictionary<string, object>());
                        continue;
                    }
                    if (val1item is Dictionary<string, object> dicStr)
                    {
                        sx.Add(dicStr);
                        continue;
                    }
                    if (val1item is ExpandoObject)
                    {
                        sx.Add(new Dictionary<string, object>(val1item as ExpandoObject));
                        continue;
                    }

                    // When any value is a primitive type, then add list as-is
                    if (val1item is string || val1item.GetType().IsValueType)
                    {
                        isValueList = true;
                        result.Add(prop.Name, propValue);
                        break;
                    }

                    PropertyDescriptorCollection props2 = TypeDescriptor.GetProperties(val1item);
                    Dictionary<string, object> result2 = new Dictionary<string, object>();
                    foreach (PropertyDescriptor prop2 in props2)
                    {
                        object val2 = prop2.GetValue(val1item);
                        result2.Add(prop2.Name, val2);
                    }
                    sx.Add(result2);
                }
                if (!isValueList)
                {
                    result.Add(prop.Name, sx);
                }
            }
            else
            {
                result.Add(prop.Name, propValue);
            }
        }
        return result;
    }

    internal static bool IsStrongTypeEnumerable(this object obj)
    {
        return obj is IEnumerable
            && obj is not string
            && obj is not char[]
            && obj is not string[]
            && obj is not IList<IMiniWordComponentList>;
    }
}