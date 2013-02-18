using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen.Types.Instructions
{
    /// <summary>
    /// Class that handles Label type.
    /// </summary>
    class CTsLabel : CTsGenericInstr
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="text">Text to be displayed</param>
        public CTsLabel(string text)
        {
            this.Text = text;
        }

        /// <summary>
        /// Static-like name of sequence
        /// </summary>
        public override string InstructionName
        {
            get { return "Label"; }
            protected set { }
        }
    }
}
