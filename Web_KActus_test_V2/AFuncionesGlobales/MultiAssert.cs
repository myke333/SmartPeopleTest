using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Web_Kactus_Test_V2
{
    public static class MultiAssert
    {
        public static void Aggregate(params Action[] actions)
        {
            var exceptions = new List<AssertFailedException>();

            foreach (var action in actions)
            {
                try
                {
                    action();
                }
                catch (AssertFailedException ex)
                {
                    exceptions.Add(ex);
                }
            }

            var assertionTexts =
                exceptions.Select(assertFailedException => assertFailedException.Message);
            if (0 != assertionTexts.Count())
            {
                throw new
                    AssertFailedException(
                    assertionTexts.Aggregate(
                        (aggregatedMessage, next) => aggregatedMessage + Environment.NewLine + next));
            }
        }
    }
}
