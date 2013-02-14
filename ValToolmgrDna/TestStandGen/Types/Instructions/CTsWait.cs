using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen.Types.Instructions
{
    /// <summary>
    /// Class that handles Label type.
    /// </summary>
    class CTsWait : CTsGenericInstr
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="duration">to wait (in seconds)</param>
        public CTsWait(int duration)
        {
            this.Text = "Pause during "+duration+"s";
        }

        /// <summary>
        /// Static-like name of sequence
        /// </summary>
        public override string InstructionName
        {
            get { return "NI_Wait"; }
            protected set { }
        }
    }
}
