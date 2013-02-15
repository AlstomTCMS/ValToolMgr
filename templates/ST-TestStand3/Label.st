Label(SeqIdx, InstrIdx, Instruction) ::= <<

[SF.Seq[$SeqIdx$].Main[$InstrIdx$].TS]
Id = "ID#:$Instruction.GUID$"

[DEF, SF.Seq[$SeqIdx$].Main[$InstrIdx$]]
%NAME = "$Instruction.Text$"

>>