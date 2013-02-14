SequenceCall(SeqIdx, InstrIdx, Instruction) ::= <<
[SF.Seq[$SeqIdx$].Main[$InstrIdx$].TS]
Id = "ID#:$Instruction.GUID$"

[DEF, SF.Seq[$SeqIdx$].Main[$InstrIdx$].TS]
SData = "TYPE, SeqCallStepAdditions"

[SF.Seq[$SeqIdx$].Main[$InstrIdx$].TS.SData]
SeqName = "$Instruction.Identifier$"
UseCurFile = True

[DEF, SF.Seq[$SeqIdx$].Main[$InstrIdx$]]
%NAME = "$Instruction.Text$"

>>