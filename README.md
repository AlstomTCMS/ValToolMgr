ValToolMgr
==========

Abstract
----------
This tool is intended to generate TestStand sequences using some Excel macros.

Demonstration sample
----------
State of art of what is able to handle this script is presented on directory <test/UnitTest 2013 00.xlsx>. This file is intended to describe all currently available features, at least once.
Remember to use the one from a released version (see #Version history below).

Version history
==========

ValToolMgr_0.3
----------
Source code : https://github.com/AlstomTCMS/ValToolMgr/tree/ValToolMgr_0.3
List of processed issues : https://github.com/AlstomTCMS/ValToolMgr/issues?milestone=4&state=closed

Main points :
 * #38 : VBA is not used anymore, it is replaced by C# which offers much more possibilities
 * #52 : Logging feature is enabled on some parts of the project. A file called log-file.txt is created where are stored DLL's. It is possible to read this file using Chainsaw, which is delivered in following directory : <tools/Chainsaw>
 * #53, #16, #20, #32 : New Teststand steps are available. It should be sufficient to begin using validation sequences.
 * #50 : Tool is able to generate all selected sheets together.
 
Main limitations :
 * #62 : No possibility to unforce arrays.
 * #47 : No possibility to generate a Multiple unit/Multiple equipment sequence file.
 * #63 : Generate .SEQ where is saved XLS file, with same name as the XLS file.
 * #64 : Outdated template for test sheets.