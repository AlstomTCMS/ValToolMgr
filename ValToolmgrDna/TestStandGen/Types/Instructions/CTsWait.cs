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
        public string Value;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="duration">to wait (in milliseconds)</param>
        public CTsWait(int duration)
        {
            Value = System.Convert.ToString((float)duration/1000).ToString().Replace(',', '.');
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
