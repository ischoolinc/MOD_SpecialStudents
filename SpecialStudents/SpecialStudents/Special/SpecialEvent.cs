using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpecialStudents
{
    public static class SpecialEvent
    {
        public static void RaiseSpecialChanged()
        {
            if (SpecialChanged != null)
                SpecialChanged(null, EventArgs.Empty);
        }

        public static event EventHandler SpecialChanged;
    }
}
