NI_Wait(SeqIdx, InstrIdx, Instruction) ::= <<
[SF.Seq[$SeqIdx$].Setup[$InstrIdx$].TS]
Id = "ID#:$Instruction.GUID$"
NoResult = True

[SF.Seq[$SeqIdx$].Setup[$InstrIdx$]]
SeqCallStepGroupIdx = -1
TimeExpr = "$Instruction.Value$"

[DEF, SF.Seq[$SeqIdx$].Setup[$InstrIdx$]]
%NAME = "Wait $Instruction.Value$ sec."

>>