using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolFunctionsStub
{
    /// <summary>
    /// Specify how the user wants an action with regard to the range he's selected
    /// </summary>
    public enum EditingZone
    {
        [StringValue("Not specified")]
        NONE = 0,
        [StringValue("Last")]
        LAST = 1,
        [StringValue("Current Up")]
        CURRENT_UP = 2,
        [StringValue("Current Down")]
        CURRENT_DOWN = 3
    }
}
