NI_Wait(SeqIdx, InstrIdx, Instruction) ::= <<
[SF.Seq[$SeqIdx$].Main[$InstrIdx$].TS]
Id = "ID#:$Instruction.GUID$"
NoResult = True

[SF.Seq[$SeqIdx$].Main[$InstrIdx$]]
SeqCallStepGroupIdx = -1
TimeExpr = "$Instruction.Value$"

[DEF, SF.Seq[$SeqIdx$].Main[$InstrIdx$]]
%NAME = "Wait $Instruction.Value$ seconds"

>>