using System;
using System.Collections.Generic;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Test Case with support to generic Test methods.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
    public class GenericTestCaseAttribute : TestCaseAttribute, ITestBuilder
    {
        // Code source from https://stackoverflow.com/questions/2364929/nunit-testcase-with-generics

        private readonly Type type;

        public GenericTestCaseAttribute(Type type, params object[] arguments) : base(arguments)
        {
            this.type = type;
        }

        IEnumerable<TestMethod> ITestBuilder.BuildFrom(IMethodInfo method, Test suite)
        {
            if (method.IsGenericMethodDefinition && type != null)
            {
                var gm = method.MakeGenericMethod(type);
                return BuildFrom(gm, suite);
            }
            return BuildFrom(method, suite);
        }
    }
}