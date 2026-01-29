using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace VH_DaglichtPlugin
{
    public class WarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
            foreach (FailureMessageAccessor f in failures)
            {
                FailureDefinitionId id = f.GetFailureDefinitionId();
                
                // Suppress "Line is slightly off axis" warning
                if (id == BuiltInFailures.InaccurateFailures.InaccurateLine)
                {
                    failuresAccessor.DeleteWarning(f);
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
