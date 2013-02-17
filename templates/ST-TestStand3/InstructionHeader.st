InstructionHeader(SeqIdx, InstrIdx, Instruction) ::= <<
[SF.Seq[$SeqIdx$].Main[$InstrIdx$].TS]
Id = "ID#:$Instruction.GUID$"
$if(Instruction.ForceFailed)$Mode = "Skipped"
$elseif(Instruction.Skipped)$Mode = "Fail"
$elseif(Instruction.ForcePassed)$Mode = "Passed"
$endif$

>>