/*
 * Copyright (c) 2017 Deal Stream sàrl. All rights reserved
 */
using System.IO;
using System.Reflection;
using System.Resources;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Helper class to get an embedded resources.
    /// </summary>
    public static class ResourceHelper
    {
        public static string GetString(string resourceName)
        {
            return GetString(typeof(ResourceHelper).GetTypeInfo().Assembly, resourceName);
        }

        public static string GetString(Assembly assembly, string resourceName)
        {
            using (var stream = GetStream(assembly, resourceName))
            {
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        public static Stream GetStream(string resourceName)
        {
            return GetStream(typeof(ResourceHelper).GetTypeInfo().Assembly, resourceName);
        }

        public static Stream GetStream(Assembly assembly, string resourceName)
        {
            var stream = assembly.GetManifestResourceStream(assembly.GetName().Name + "." + resourceName);
            if (stream == null)
                throw new MissingManifestResourceException($"Requested resource `{resourceName}` was not found in the assembly `{assembly}`.");

            return stream;
        }
    }
}