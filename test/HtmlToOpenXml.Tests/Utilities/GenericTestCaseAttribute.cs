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

        private readonly Type _type;

        public GenericTestCaseAttribute(Type type, params object[] arguments) : base(arguments)
        {
            this._type = type;
        }

        IEnumerable<TestMethod> ITestBuilder.BuildFrom(IMethodInfo method, Test suite)
        {
            if (method.IsGenericMethodDefinition && _type != null)
            {
                var gm = method.MakeGenericMethod(_type);
                return BuildFrom(gm, suite);
            }
            return BuildFrom(method, suite);
        }
    }
}