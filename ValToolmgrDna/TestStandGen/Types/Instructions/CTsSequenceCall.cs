using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen.Types.Instructions
{
    /// <summary>
    /// Class that handles SequenceCall type.
    /// </summary>
    class CTsSequenceCall : CTsGenericInstr
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="identifier">Identifier of sequence</param>
        /// <param name="text">Text to be displayed</param>
        public CTsSequenceCall(string identifier, string text)
        {
            this.Identifier = identifier;
            this.Text = text;
        }

        public override string InstrTsName
        {
            get { return "SequenceCall"; }
            protected set { }
        }

        /// <summary>
        /// Identifier of sequence to call
        /// </summary>
        public string Identifier;
    }
}
