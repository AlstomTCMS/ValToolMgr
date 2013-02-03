Attribute VB_Name = "TestModule"
Option Explicit

Private Function testStep(Idx As Integer) As CStep
    Set testStep = New CStep
    
    testStep.title = "Step_" & Idx
    
    Dim numero As Integer
    numero = 1 'Numéro de départ (correspond ici au n° de ligne et au n° de numérotation)

    While numero <= 12 'TANT QUE la variable numero est <= 12, la boucle est répétée
       Dim o_action As CInstruction
        Set o_action = New CInstruction
    
        o_action.category = A_FORCE
        o_action.Data = "toto " & numero
        testStep.AddInstruction o_action
        numero = numero + 1
    Wend
End Function

Private Function testTest(Idx As Integer) As CTest
    Set testTest = New CTest
    
    testTest.title = "Test_" & Idx
    Dim numero As Integer
    numero = 1 'Numéro de départ (correspond ici au n° de ligne et au n° de numérotation)

    While numero <= 12 'TANT QUE la variable numero est <= 12, la boucle est répétée
       Dim o_step As CInstruction
        Set o_step = New CInstruction

        testTest.AddStep testStep(numero)
        numero = numero + 1
    Wend
End Function

Private Function TestContainer() As CTestContainer
    Set TestContainer = New CTestContainer

    TestContainer.title = "TestContainer"
    Dim numero As Integer
    numero = 1 'Numéro de départ (correspond ici au n° de ligne et au n° de numérotation)

    While numero <= 2 'TANT QUE la variable numero est <= 12, la boucle est répétée
       Dim o_test As CTest
        Set o_test = New CTest
        TestContainer.AddTest testTest(numero)
        numero = numero + 1
    Wend

End Function

Private Sub testScenarioTs()
    Dim cContainer As CTestContainer
    Set cContainer = TestContainer()
    
    Dim genTs As GeneratorTs401
    Set genTs = New GeneratorTs401
    
    Call genTs.writeScenario("C:\\macros_alstom\\test\\testGen.seq", cContainer)
End Sub
