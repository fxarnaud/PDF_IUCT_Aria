using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;

namespace PDF_IUCT
{
    public class PreliminaryTests
    {
        public PreliminaryTests()
        {           
        }

        public static void Test(ScriptContext context)
        {
            if (context == null)
            {
                throw new ApplicationException("Merci de charger un patient et un plan");
            }
            if (context.PlanSetup == null)
            {
                throw new ApplicationException("Merci de charger un plan");
            }
            if (context.PlanSetup.PlanIntent.Equals("VERIFICATION"))
            {
                throw new ApplicationException("Merci de charger un plan qui ne soit pas un plan de vérification");
            }
            if (!context.PlanSetup.IsDoseValid)
            {
                throw new ApplicationException("Merci de charger un plan avec une dose");
            }

        }
    }
}
