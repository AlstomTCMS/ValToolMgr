Attribute VB_Name = "GenCodeTestStand_Access"
Option Explicit

Dim strNomEqtTemp, sNomVarTemp As String


Public Sub GenCodeMatriceBC(RanginSeq As Integer)
'RanginSeq est le rang de la séquence d'init
Dim strContexte As String
Dim reset_fin_etape As Boolean

    ' Création de la séquence de la matrice en setup
    genLabel sequence, RanginSeq, "-------------------------------------------------", False, StepGroup_Setup: RanginSeq = RanginSeq + 1
    genLabel sequence, RanginSeq, "Init de la matrice du BC", True, StepGroup_Setup: RanginSeq = RanginSeq + 1
    genLabel sequence, RanginSeq, "-------------------------------------------------", False, StepGroup_Setup: RanginSeq = RanginSeq + 1
    reset_fin_etape = False
    
    genMatriceInit sequence, RanginSeq, StepGroup_Setup: RanginSeq = RanginSeq + 1
    genMatriceConnect sequence, RanginSeq, "", StepGroup_Setup: RanginSeq = RanginSeq + 1
    
RanginSeq = rg_EtapeDansSequence
End Sub

Public Function GenVariableFip_DDU(strNomVariableFip As String, ByVal sDirectionFrom As String, ByVal sDirectionTo As String, ByVal strNumVehicle As String, ByVal sModuleMmap As String, ByVal bUnForce As Boolean, ByVal bLecture As Boolean, ByVal bPresentRP As Boolean, Optional strNumEtape As String, Optional sText As String) As Integer
' intError -> Numéro d'erreur (0=OK, 1=erreur indice, 2=erreur interne, 4=Pas de module MMAP)
'Variable provenant du DDU
Dim intError, iIndice, iIndiceVar As Integer

intError = 0 'La génération s'est bien passé

    If Mid(sDirectionTo, 1, 3) = "DDU" Then
        sText = strNomVariableFip & " en interne dans le DDU"
        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        intError = 2
        GoTo error
    Else
        If Right(strNomVariableFip, 1) = ")" Then
            If (Mid(sDirectionFrom, 4, 2)) = "" Then
                sText = "Pas d'indice entre le DDU et la var " & strNomVariableFip
                If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
                intError = 1
                GoTo error
            Else
                iIndice = Mid(sDirectionFrom, 4, 2)
            End If
            If (Mid(Right(strNomVariableFip, 3), 1, 2)) <> CStr(iIndice) Then
                sText = "Erreur d'indiçage entre la var(i) et la DDU(i) " & strNomVariableFip
                If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
                intError = 3
                GoTo error
            Else
                iIndiceVar = Mid(Right(strNomVariableFip, 3), 1, 2)
            End If
            strNomVariableFip = Mid(strNomVariableFip, 1, Len(strNomVariableFip) - 4)
        Else
            sText = "Pas d'indice entre le DDU et la var " & strNomVariableFip
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
            intError = 1
            GoTo error
        End If
        
        If strNomVariableFip <> "B_CONS1" And strNomVariableFip <> "B_PRINCI" And _
        Mid(strNomVariableFip, 1, 6) <> "B_BOOL" And Mid(strNomVariableFip, 1, 4) <> "N_MO" And _
        strNomVariableFip <> "N_OCTRES1" Then
          '  strNomEqtTemp = "L" & strNumVehicle & "_C" & iIndice
            strNomEqtTemp = "@\\" & sModuleMmap
            sNomVarTemp = strNomEqtTemp & "XM_" & strNomVariableFip
        Else
           ' strNomEqtTemp = "L" & strNumVehicle & "_C" & iIndice
            strNomEqtTemp = "@\\" & sModuleMmap & "\\"
            sNomVarTemp = strNomEqtTemp & "MX_" & strNomVariableFip
        End If
    End If
error:
GenVariableFip_DDU = intError
End Function


Public Function GenVariableFip_ACU(strNomVariableFip As String, ByVal sDirectionFrom As String, ByVal sDirectionTo As String, ByVal strNumVehicle As String, ByVal sModuleMmap As String, ByVal bUnForce As Boolean, ByVal bLecture As Boolean, ByVal bPresentRP As Boolean, Optional strNumEtape As String, Optional sText As String) As Integer
' intError -> Numéro d'erreur (0=OK, 1=erreur indice, 2=erreur interne)
'Variable provenant de l'ACU
Dim intError, iIndice, iIndiceVar As Integer

intError = 0 'La génération s'est bien passé

    If Mid(sDirectionTo, 1, 3) = "ACU" Then
        sText = strNomVariableFip & " en interne dans le ACU"
        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        intError = 2
        GoTo error
    Else
        If (Mid(sDirectionFrom, 4, 1)) = "" Then
            sText = "Pas d'indice entre le ACU et la var " & strNomVariableFip
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
            intError = 3
            GoTo error
        Else
            iIndice = Mid(sDirectionFrom, 4, 1)
        End If
        If (Mid(Right(strNomVariableFip, 2), 1, 1)) <> CStr(iIndice) Then   'on vérifie que l'indice de la var et l'indice de l'équipement sont corrects
            sText = "Erreur d'indiçage entre la var(i) et l'ACU(i) " & strNomVariableFip
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
            intError = 1
            GoTo error
        Else
            iIndiceVar = Mid(Right(strNomVariableFip, 2), 1, 1)
        End If
        
        strNomVariableFip = Mid(strNomVariableFip, 1, Len(strNomVariableFip) - 3)   'on supprime l'indiçage de la variable
        If Mid(strNomVariableFip, 1, 5) <> "B_CVT" And _
        Mid(strNomVariableFip, 3, 8) <> "CRACRT" Then
            strNomEqtTemp = "L" & strNumVehicle & "_AU" & iIndice
            sNomVarTemp = strNomEqtTemp & "XM_" & strNomVariableFip
        Else
            strNomEqtTemp = "L" & strNumVehicle & "_AU" & iIndice
            sNomVarTemp = strNomEqtTemp & "_" & strNomVariableFip
        End If
        If bPresentRP = False Then
            sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
            strNomEqtTemp = "@\\" & sModuleMmap
            If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Then
                If bUnForce = False Then
                    If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                    If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
                Else
                    sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
                End If
            ElseIf sModuleMmap = "" Then
                sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                intError = 4
                GoTo error
            End If
        End If
    End If
error:
GenVariableFip_ACU = intError
End Function


Public Function GenVariableFip_BCU(strNomVariableFip As String, ByVal sDirectionFrom As String, ByVal sDirectionTo As String, ByVal strNumVehicle As String, ByVal sModuleMmap As String, ByVal bUnForce As Boolean, ByVal bLecture As Boolean, ByVal bPresentRP As Boolean, Optional strNumEtape As String, Optional sText As String) As Integer
' intError -> Numéro d'erreur (0=OK, 1=erreur indice, 2=erreur interne)
'Variable provenant du BCU
Dim intError, iIndice, iIndiceVar As Integer

intError = 0 'La génération s'est bien passé
    
    If Mid(sDirectionTo, 1, 3) = "BCU" Then
        sText = strNomVariableFip & " en interne dans le BCU"
        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        intError = 2
        GoTo error
    Else
        If Mid(strNomVariableFip, 3, 5) <> "BCRES" And _
        strNomVariableFip <> "B_DMDRE" And _
        strNomVariableFip <> "B_VTL00" And _
        strNomVariableFip <> "B_CG" And _
        strNomVariableFip <> "B_MDGBCU" And _
        strNomVariableFip <> "B_PREGBCU" And _
        strNomVariableFip <> "B_ISBCU" And _
        strNomVariableFip <> "N_PREGBCU" And _
        strNomVariableFip <> "N_PCGVB" And _
        strNomVariableFip <> "N_PCPVB" Then
            strNomEqtTemp = "L" & strNumVehicle & "_BC1"
            sNomVarTemp = strNomEqtTemp & "XM_" & strNomVariableFip
        Else
            strNomEqtTemp = "L" & strNumVehicle & "_BC1"
            sNomVarTemp = strNomEqtTemp & "MX_" & strNomVariableFip
        End If
        If bPresentRP = False Then
            sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
            strNomEqtTemp = "@\\" & sModuleMmap
            If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Then
                If bUnForce = False Then
                    If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                    If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
                Else
                    sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
                End If
            ElseIf sModuleMmap = "" Then
                sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                intError = 4
                GoTo error
            End If
        End If
    End If
error:
GenVariableFip_BCU = intError
End Function

Public Function GenVariableFip_RIO_LVV(strNomVariableFip As String, ByVal sDirectionFrom As String, ByVal sDirectionTo As String, ByVal strNumVehicle As String, ByVal sModuleMmap As String, ByVal bUnForce As Boolean, ByVal bLecture As Boolean, ByVal bPresentRP As Boolean, Optional strNumEtape As String, Optional sText As String)
' intError -> Numéro d'erreur (0=OK, 1=erreur indice, 2=erreur interne)
'Variable provenant du RIOM ou LVV
Dim intError, iIndice, iIndiceVar As Integer

intError = 0 'La génération s'est bien passé
    
    If Mid(sDirectionFrom, 1, 3) = "RIO" And Mid(strNomVariableFip, 2, 3) = "_BT" Then
        If (Mid(sDirectionFrom, 5, 1)) = "" Then
            sText = "Pas d'indice entre le RIOM et la var " & strNomVariableFip
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
            intError = 1
            GoTo error
        Else
            iIndice = Mid(sDirectionFrom, 5, 1)
        End If
        
        If Mid(strNomVariableFip, 1, 4) <> "E_BT" And _
        Mid(strNomVariableFip, 1, 4) <> "S_BT" Then
            'Cas non traité, car la plupart des variables sont en réserve
        Else
            strNomEqtTemp = "L" & strNumVehicle & "_R" & iIndice
            If Mid(strNomVariableFip, 1, 1) = "E" Then
                strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
                sNomVarTemp = strNomEqtTemp & "EL_" & strNomVariableFip
            ElseIf Mid(strNomVariableFip, 1, 1) = "S" Then
                strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
                sNomVarTemp = strNomEqtTemp & "SL_" & strNomVariableFip
            End If
        End If
        
        If (bPresentRP = False) And (sNomVarTemp <> "") Then
            sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
            strNomEqtTemp = "@\\" & sModuleMmap
            If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Then
                If bUnForce = False Then
                    If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                    If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
                Else
                    sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
                End If
            ElseIf sModuleMmap = "" Then
                sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                intError = 4
                GoTo error
            End If
        End If
    ElseIf Mid(sDirectionFrom, 1, 3) = "RIO" And Mid(strNomVariableFip, 1, 4) = "E_CA" Then
    'Cas particulier pour les capteurs analogiques
        If strNomVariableFip = "E_CA_1MPCOTF1" Or strNomVariableFip = "E_CA_2MPCOTF1" Then
            iIndice = Mid(strNomVariableFip, 6, 1)
            strNomVariableFip = Left(strNomVariableFip, Len(strNomVariableFip) - 1)
            strNomEqtTemp = "L" & strNumVehicle & "_R" & iIndice
            If Mid(strNomVariableFip, 1, 1) = "E" Then
                strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
                sNomVarTemp = strNomEqtTemp & "EA_" & strNomVariableFip
            ElseIf Mid(strNomVariableFip, 1, 1) = "S" Then
                strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
                sNomVarTemp = strNomEqtTemp & "SA_" & strNomVariableFip
            End If
        End If
        
        If (bPresentRP = False) And (sNomVarTemp <> "") Then
            sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
            strNomEqtTemp = "@\\" & sModuleMmap
            If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Then
                If bUnForce = False Then
                    If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                    If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
                Else
                    sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
                End If
            ElseIf sModuleMmap = "" Then
                sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                intError = 4
                GoTo error
            End If
        End If
    
    ElseIf Mid(sDirectionFrom, 1, 3) = "LVV" Then
    'Dans le cas du LVV, toutes les variables sont automatiquement dans le serveur RP
        If Mid(strNomVariableFip, 1, 4) <> "E_BT" And _
        Mid(strNomVariableFip, 1, 4) <> "S_BT" Then
            'Cas non traité, car la plupart des variables sont en réserve
        Else
            If Mid(strNomVariableFip, 1, 5) <> "E_BT_" And Mid(strNomVariableFip, 1, 5) <> "S_BT_" Then
                iIndice = Mid(strNomVariableFip, 5, 1)
                strNomEqtTemp = "L" & strNumVehicle & "_R" & iIndice
            Else
                strNomEqtTemp = "L" & strNumVehicle & "_R1"
            End If
            If Mid(strNomVariableFip, 1, 1) = "E" Then
                strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
                sNomVarTemp = strNomEqtTemp & "EL_" & strNomVariableFip
            ElseIf Mid(strNomVariableFip, 1, 1) = "S" Then
                strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
                sNomVarTemp = strNomEqtTemp & "SL_" & strNomVariableFip
            End If
        End If
    Else
        sText = strNomVariableFip & " en interne dans le RIOM"
        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        intError = 2
        GoTo error
    End If
error:
GenVariableFip_RIO_LVV = intError
End Function

Public Function GenVariableFip_TCU(strNomVariableFip As String, ByVal sDirectionFrom As String, ByVal sDirectionTo As String, ByVal strNumVehicle As String, ByVal sModuleMmap As String, ByVal bUnForce As Boolean, ByVal bLecture As Boolean, ByVal bPresentRP As Boolean, Optional strNumEtape As String, Optional sText As String) As Integer
' intError -> Numéro d'erreur (0=OK, 1=erreur indice, 2=erreur interne)
'Variable provenant du TCU
Dim intError, iIndice, iIndiceVar As Integer

intError = 0 'La génération s'est bien passé

    If Mid(sDirectionTo, 1, 3) = "TCU" Then
    'Vérifie que la variable n'est pas en interne
        sText = strNomVariableFip & " en interne dans le TCU"
        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        intError = 2
        GoTo error
    Else
        If (Mid(sDirectionFrom, 4, 1)) = "" Then
            sText = "Pas d'indice entre le TCU et la var " & strNomVariableFip
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
            intError = 3
            GoTo error
        Else
            iIndice = Mid(sDirectionFrom, 4, 1) 'l'indice est bien renseigné sur l'equipement
        End If
        If (Mid(Right(strNomVariableFip, 2), 1, 1)) <> CStr(iIndice) Then   'on vérifie que l'indice de la var et l'indice de l'équipement sont corrects
            sText = "Erreur d'indiçage entre la var(i) et le TCU(i) " & strNomVariableFip
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
            intError = 1
            GoTo error
        Else
            iIndiceVar = Mid(Right(strNomVariableFip, 2), 1, 1) 'Test de cohérence des indices
        End If
        
        strNomVariableFip = Mid(strNomVariableFip, 1, Len(strNomVariableFip) - 3)
        If Mid(strNomVariableFip, 1, 7) <> "B_AMCVS" And _
        Mid(strNomVariableFip, 3, 8) <> "CRTCRA" And _
        Mid(strNomVariableFip, 1, 5) <> "N_FVT" And _
        Mid(strNomVariableFip, 1, 7) <> "B_CVTES" And _
        Mid(strNomVariableFip, 1, 6) <> "B_INUL" And _
        Mid(strNomVariableFip, 1, 6) <> "N_FAUX" Then
            strNomEqtTemp = "L" & strNumVehicle & "_VA" & iIndice
            sNomVarTemp = strNomEqtTemp & "XM_" & strNomVariableFip
        Else
            strNomEqtTemp = "L" & strNumVehicle & "_VA" & iIndice
            sNomVarTemp = strNomEqtTemp & "_" & strNomVariableFip
        End If
        
        If bPresentRP = False Then
            sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
            strNomEqtTemp = "@\\" & sModuleMmap
            If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Then
                If bUnForce = False Then
                    If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                    If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
                Else
                    sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
                End If
            ElseIf sModuleMmap = "" Then
                sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                intError = 4
                GoTo error
            End If
        End If
    End If
error:
GenVariableFip_TCU = intError
End Function

Public Function GenVariableFip_MPU(strNomVariableFip As String, ByVal sDirectionFrom As String, ByVal sDirectionTo As String, ByVal strNumVehicle As String, ByVal sModuleMmap As String, ByVal bUnForce As Boolean, ByVal bLecture As Boolean, ByVal bPresentRP As Boolean, Optional strNumEtape As String, Optional sText As String) As Integer
' intError -> Numéro d'erreur (0=OK, 1=erreur indice, 2=erreur interne)
'Variable provenant du TCU
Dim intError, iIndice, iIndiceVar As Integer

intError = 0 'La génération s'est bien passé

If Left(strNomVariableFip, 5) = "E_BTR" Or Left(strNomVariableFip, 5) = "S_BTR" Then
'Cas 1 : ModuleMmap=ISaGRAF et la variable est du type E_BT ou S_BT
    If (sModuleMmap = "ISaGRAF\\") And (bPresentRP = False) Then
        sNomVarTemp = strNomVariableFip
        strNomEqtTemp = "@\\" & sModuleMmap
        bRedondance = False 'Annulation de la redondance seulement dans ce cas
        GoTo SORTIE
    End If
'Cas 2 : la variable est du type E_BT ou S_BT
    iIndice = 1
    strNomEqtTemp = "L" & strNumVehicle & "_R" & iIndice   'V1_R1
    If Mid(strNomVariableFip, 1, 1) = "E" Then
        strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
        sNomVarTemp = strNomEqtTemp & "EL_" & strNomVariableFip  'V1_R1EL_
    ElseIf Mid(strNomVariableFip, 7, 5) = "VE1GB" Or Mid(strNomVariableFip, 7, 5) = "VESA1" Or Mid(strNomVariableFip, 7, 5) = "VESA2" Then
        strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
        sNomVarTemp = strNomEqtTemp & "SS_" & strNomVariableFip  'V1_R1SS_
    ElseIf Mid(strNomVariableFip, 1, 1) = "S" Then
        strNomVariableFip = Mid(strNomVariableFip, 7, Len(strNomVariableFip) - 6)
        sNomVarTemp = strNomEqtTemp & "SL_" & strNomVariableFip  'V1_R1SL_
    End If
'Cas 3 : la variable est du type E_BT et instanciée
Else
    If Left(strNomVariableFip, 2) = "E_" And Right(strNomVariableFip, 1) = ")" Then
    End If
    If sDirectionTo <> "" Then
        Select Case Mid(sDirectionTo, 1, 3)
            Case "TCU"
                sDirectionTo = "VA"
            Case "ACU"
                sDirectionTo = "AU"
            Case "BCU"
                sDirectionTo = "BC"
            Case "RIO"
                sDirectionTo = "R"
            Case "DDU"
                If Right(strNomVariableFip, 1) = ")" Then
                    sDirectionTo = "C"
                Else
                    sDirectionTo = "V"
                End If
            Case "MPU"
                'variable interne au MPU
                bPresentRP = False
                If bPresentRP = False And Right(strNomVariableFip, 1) <> ")" Then
                'La variable n'est pas présente sur le serveur R/P
                    strNomEqtTemp = "@\\" & sModuleMmap
                    If sModuleMmap = "es_960\\" Then
                        If bUnForce = False Then
                            If bLecture = False Then strNomVariableFip = "SIGNAL_VMX_" & strNomVariableFip & "$SignalInhiber"
                            If bLecture = True Then strNomVariableFip = "SIGNAL_VMX_" & strNomVariableFip & "$SignalValue"
                        Else
                            strNomVariableFip = "SIGNAL_VMX_" & strNomVariableFip & "$SignalAutoriser"
                        End If
                    ElseIf sModuleMmap = "" Then
                        sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                        intError = 4
                        GoTo error
                    End If
                    sNomVarTemp = strNomVariableFip
                ElseIf bPresentRP = False And Right(strNomVariableFip, 1) = ")" Then
                    iIndiceVar = Mid(Right(strNomVariableFip, 2), 1, 1)
                    strNomEqtTemp = "@\\" & sModuleMmap
                    If sModuleMmap = "es_960\\" Then
                        If bUnForce = False Then
                            If bLecture = False Then strNomVariableFip = "SIGNAL_VMX_" & strNomVariableFip & "$SignalInhiber"
                            If bLecture = True Then strNomVariableFip = "SIGNAL_VMX_" & strNomVariableFip & "$SignalValue"
                        Else
                            strNomVariableFip = "SIGNAL_VMX_" & strNomVariableFip & "$SignalAutoriser"
                        End If
                    ElseIf sModuleMmap = "" Then
                        sText = "Pas de module MMAP associé à la variable " & strNomVariableFip
                        intError = 4
                        GoTo error
                    End If
                    sNomVarTemp = strNomVariableFip
                ElseIf bPresentRP = True And Right(strNomVariableFip, 1) <> ")" Then
                    MsgBox "Variable " & strNomVariableFip & " non correct" & vbCr & "Présent R/P=" & bPresentRP & Right(strNomVariableFip, 1), vbOKOnly, "Erreur"
                
                ElseIf bPresentRP = True And Right(strNomVariableFip, 1) = ")" Then
                    iIndiceVar = Mid(Right(strNomVariableFip, 2), 1, 1)
                    MsgBox "Variable " & strNomVariableFip & " non correct" & vbCr & "Présent R/P=" & bPresentRP & Right(strNomVariableFip, 1), vbOKOnly, "Erreur"
                  
                End If
                GoTo SORTIE
            Case Else
                'Autres cas où l'EQT est non utilisé ou mal renseigné
                sText = "Erreur, pas de nom d'EQT valide pour : " & strNomVariableFip
                If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
                'GoTo error
        End Select
        If Right(strNomVariableFip, 1) = ")" Then
            'Cas d'une variable avec des instanciations ex: B_AAISCO(4)
            iIndiceVar = Mid(Right(strNomVariableFip, 2), 1, 1)  '4
            strNomVariableFip = Mid(strNomVariableFip, 1, Len(strNomVariableFip) - 3) 'B_AAISCO
            strNomEqtTemp = "L" & strNumVehicle & "_" & sDirectionTo & iIndiceVar 'V1_VA4
        Else
            strNomEqtTemp = "L" & strNumVehicle & "_" & sDirectionTo  'V1_BC
        End If
        
        If sDirectionTo = "C" Then
            sNomVarTemp = strNomEqtTemp & "MX_" & strNomVariableFip
        Else
            sNomVarTemp = "L" & strNumVehicle & "_VMX_" & strNomVariableFip 'V1_VMX_
        End If
        
        If bPresentRP = False Then
            sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
            'If Right(sNomVarTemp, 2) = "V1" Or Right(sNomVarTemp, 2) = "V2" Or Right(sNomVarTemp, 2) = "V3" Or Right(sNomVarTemp, 2) = "V4" Then sNomVarTemp = Mid(sNomVarTemp, 1, Len(sNomVarTemp) - 2)
            If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Or sModuleMmap = "isa_inte\\" Then
                If bUnForce = False Then
                    If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                    If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
                Else
                    sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
                End If
            ElseIf sModuleMmap = "ISaGRAF\\" And Mid(sNomVarTemp, 1, 4) = "VMX_" Then
                sNomVarTemp = Mid(sNomVarTemp, 5, Len(sNomVarTemp) - 4)
            End If
            strNomEqtTemp = "@\\" & sModuleMmap
        ElseIf bPresentRP = True Then
            If Right(sNomVarTemp, 2) = "V1" Or Right(sNomVarTemp, 2) = "V2" Or Right(sNomVarTemp, 2) = "V3" Or Right(sNomVarTemp, 2) = "V4" Then sNomVarTemp = Mid(sNomVarTemp, 1, Len(sNomVarTemp) - 2)
            strNomEqtTemp = "L" & strNumVehicle & "_" & sMPU_DIFFUSE
        End If
    Else
        'Variable sans équipement de destination
        sText = "Erreur, pas d'EQT de sortie valide pour : " & strNomVariableFip
        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        intError = 5
        GoTo error
    End If
End If

    If (bPresentRP = False) And (Left(strNomEqtTemp, 1) <> "@") Then
        sNomVarTemp = Mid(sNomVarTemp, 4, Len(sNomVarTemp) - 3)
        strNomEqtTemp = "@\\" & sModuleMmap
        If sModuleMmap = "es_960\\" Or sModuleMmap = "public\\" Then
            If bUnForce = False Then
                If bLecture = False Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalInhiber"
                If bLecture = True Then sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalValue"
            Else
                sNomVarTemp = "SIGNAL_" & sNomVarTemp & "$SignalAutoriser"
            End If
        ElseIf sModuleMmap = "" Then
            sText = "Erreur, la variable : " & strNomVariableFip & " n'a pas de module Mmap associé"
            intError = 4
            GoTo error
        End If
    ElseIf bPresentRP = True Then
        'Aucun changement
    End If


SORTIE:

error:
GenVariableFip_MPU = intError
End Function


Public Sub PasteInitMmap(sVehicule As String, ByVal RanginSeq As Integer, ByVal iDest As Integer) 'StepGroup_Main)

    If CInt(sVehicule) > 1 Then
        genAppelSequenceMMAP_INIT sequence, sVehicule, RanginSeq, iDest 'Creation MMAP dans vehicule >= 2
    Else
        genCreateObjectMMAP sequence, sVehicule, RanginSeq, iDest 'Creation MMAP dans le Vehicule 1
    End If
End Sub


Public Function GenCodeConfigTrain(ByVal NomTest As String)

Dim bUm1, bUm2, bUmd As Boolean
Dim iNBVehicle As Integer
Dim sInstance As String
Dim sTitreConfig As String
Dim rst As Recordset

'If Left(NomTest, 1) = "B" Then
'    sTitreConfig = "AT"     'Prima par défaut
'End If
'If Left(NomTest, 1) = "C" Then 'Détection du type de PR Maroc
'    sTitreConfig = "MA"
'End If
'If Left(NomTest, 1) = "D" Then
'    sTitreConfig = "AX"     'Prima par défaut
'End If


If PR_Origine <> PR_Destination Then   'Inversion activé
    If PR_Origine = 1 And PR_Destination = 2 Then sTitreConfig = "MA"
    If PR_Origine = 1 And PR_Destination = 3 Then sTitreConfig = "AX"
'    If PR_Origine = 1 And PR_Destination = 4 Then
    
    If PR_Origine = 2 And PR_Destination = 1 Then sTitreConfig = "AT"
    If PR_Origine = 2 And PR_Destination = 3 Then sTitreConfig = "AX"
'    If PR_Origine = 2 And PR_Destination = 4 Then
    
    If PR_Origine = 3 And PR_Destination = 2 Then sTitreConfig = "MA"
    If PR_Origine = 3 And PR_Destination = 1 Then sTitreConfig = "AT"
'    If PR_Origine = 3 And PR_Destination = 4 Then
    
'    If PR_Origine = 4 And PR_Destination = 1 Then
'    If PR_Origine = 4 And PR_Destination = 2 Then
'    If PR_Origine = 4 And PR_Destination = 3 Then
    
'    If sTitreConfig = "MA" Then
'        sTitreConfig = "AT"
'    Else
'        sTitreConfig = "MA"
'    End If
Else    'Pas d'inversion    (origine==destination)
    If PR_Origine = 1 Then sTitreConfig = "EP"
 '   If PR_Origine = 2 Then sTitreConfig = "MA"
 '   If PR_Origine = 3 Then sTitreConfig = "AX"
End If
sTitreConfig = ""

ParametersSequence sTitreConfig


End Function

Public Sub GenCodeCheckModeleAuto(RanginSeq As Integer, sCodeMis As String, iDest As Integer, iNbVehicleMax As Integer)
Dim rst As Recordset

Set rst = db_ihm.OpenRecordset("SELECT ID_Mission, Mission_Code FROM T_Mission WHERE Mission_Code=" & Chr(34) & sCodeMis & Chr(34) & ";")
If Not (rst.EOF) Then
    GenMission RanginSeq, rst.Fields("ID_Mission"), StepGroup_Setup, iNbVehicleMax, "E"
Else
    iContinue = 4
End If
End Sub

Public Sub GenMission(RanginSeq As Integer, Id_mission As Integer, iDest As Integer, Optional iNbVehicleMax As Integer, Optional sDirectionMission As String)

Dim arMission(30, 1) As String  'Code X Indice
Dim sAdress_FIP, sCodeMis As String
Dim iMax As Integer
Dim CodeMission, TitreMission As String  'Code & Titre Mission du PR
Dim nTempoo As Long
Dim i As Integer

nTempoo = 3000
sAdress_FIP = ArrTable_Config_Rezo(1)

''''''temporaire
'If sDirectionMission = "S" Then
''Eteindre le MPU et se déconnecter des taches CB
'    genAppelSéquenceAutreFichier sequence, 0, "Batterie Hors service", "MainSequence", "C:\Teststand\MPU_OFF.seq", False, False, ""
'    Exit Sub
'End If

Set db_ihm = CurrentDb

'Set rs_mission = db_ihm.OpenRecordset("SELECT * FROM T_Mission WHERE ID_Mission=" & Id_mission & ";")
'If Not (rs_mission.EOF) Then
'    CodeMission = rs_mission.Fields("Mission_Code")
'    TitreMission = rs_mission.Fields("Mission_Designation")
'Else
'    iContinue = 1
'    Set db_ihm = Nothing
'    Exit Sub
'End If

If Len(CodeMission) = 0 Then
    iContinue = 1
    Exit Sub
End If

Set rs_etapemission = db_ihm.OpenRecordset("SELECT T_EtapeMission.*, * FROM (T_Mission INNER JOIN T_Mission_EtapeMission ON T_Mission.ID_Mission = T_Mission_EtapeMission.ID_Mission) INNER JOIN T_EtapeMission ON T_Mission_EtapeMission.ID_EtapeMission = T_EtapeMission.ID_EtapeMission WHERE T_Mission.ID_Mission=" & Id_mission & ";", dbOpenDynaset)
While rs_etapemission.EOF = False
    If sDirectionMission = Mid(rs_etapemission.Fields("EtapeMission_Code"), 2, 1) Then
        arMission(i, 0) = rs_etapemission.Fields("EtapeMission_Code")
        arMission(i, 1) = rs_etapemission.Fields("EtapeMission_Indice")
        i = i + 1
    End If
    rs_etapemission.MoveNext
Wend
Set rs_etapemission = Nothing

iMax = i
iNbVehicleMax = 1
If CodeMission = "M" Then 'Cas particulier pour checker en fonction du nombre de véhicule
    'Le mode Mission "M" est toujours en entrée
    If sDirectionMission = Mid(arMission(0, 0), 2, 1) Then
        For i = 1 To iNbVehicleMax
            GenCodeMission arMission(0, 0), arMission(0, 1), CStr(CodeMission), nTempoo, iDest, sAdress_FIP, RanginSeq, i
        Next
    End If
Else
    For i = 0 To iMax - 1
        Set rs_missionetapemission = db_ihm.OpenRecordset("SELECT EtapeMission_Tempo,EtapeMission_Code FROM T_EtapeMission WHERE EtapeMission_Code=" & Chr(34) & arMission(i, 0) & Chr(34) & ";") ' " AND EtapeMission_Indice=" & Chr(34) & arMission(i, 1) & Chr(34) & ";")
        If IsNull(rs_missionetapemission.Fields("EtapeMission_Tempo")) = True Then
            nTempoo = 1000
        Else
            nTempoo = rs_missionetapemission.Fields("EtapeMission_Tempo").Value
        End If
        GenCodeMission arMission(i, 0), arMission(i, 1), CStr(CodeMission), nTempoo, iDest, sAdress_FIP, RanginSeq
        rs_missionetapemission.Close
        Set rs_missionetapemission = Nothing
    Next
End If

Set db_ihm = Nothing
End Sub

Public Sub genUnForceAllVariables(ByVal iDest As Integer)
'Remplissage de la page CleanUp du mode de Sortie
Dim iCounter, iLigne As Integer
    
    iLigne = 0
    
    For iCounter = 1 To 500
        If ForceVar.VarModel(iCounter) <> "" Then
            If ForceVar.CounterVar(iCounter) > 0 Then
                If ForceVar.Instance(iCounter) <> "" Then
                    genUnForceModel sequence, iLigne, ForceVar.Instance(iCounter), 1, ForceVar.VarModel(iCounter), "", True, iDest: iLigne = iLigne + 1
                    genWaitTS sequence, iLigne, 1, iDest: iLigne = iLigne + 1
                Else
                    genUnForceNNModel sequence, iLigne, 1, ForceVar.VarModel(iCounter), "", True, iDest: iLigne = iLigne + 1
                    genWaitTS sequence, iLigne, 1, iDest: iLigne = iLigne + 1
                End If
            End If
        Else
            Exit For
        End If
    Next
    
End Sub


Public Sub GenCodeMission(ETP_CODE_M As String, ETP_INDICE_M As String, ModeMission As String, nTempo As Long, iDest As Integer, ByVal strAddress_Fip As String, Optional RanginSeq As Integer = 0, Optional iNumVehMission As Integer = 0)
'Les Modes Mission sont dans l'ordre du style A1, A2, A3, l'appel ne se fait que sur la lettre
  '  On Error GoTo GenCodeMission_Err
Dim rstRappel As Recordset
Dim rst_F_Act, rst_F_Verif As Recordset 'tables correspondant aux variables du FIP
Dim rst_M_Act, rst_M_Verif As Recordset 'tables correspondant aux variables du MODEL
Dim rst_F_EQT, rst_M_EQT As Recordset   'tables pointant sur l'equipement
Dim rst_Plage As Recordset              'table pointant sur la plage de la variable
Dim bds As Database
Dim sInstance, sVariable, sVariableTempo, sValeur, str_EQT, strTolerance As String
Dim sNomPop, sTitrePop, sMessagePop, sText, sTemp As String
Dim sVehicule, sTempoVehicule, strPlage, sModuleMmap As String
Dim reset_fin_etape, boolPresentRP, boolAccesNickname As Boolean
Dim iNumConsist, iCounter, rg, iHandleMmap, iPositionCherche As Integer
Dim bMPU1, bMPU2 As Boolean             'Lecture/Ecriture autorisée
Dim strContexte As String
Dim bEnvir As Boolean    'Variable permettant de savoir si les fonctions CB sont pour l'env ou l'emb
   
rg = RanginSeq
Set bds = CurrentDb
'----------------------------------------
'----------------------------------------
' Ouvre les tables correspondant aux étapes (Action pour le MVB et le Model)
'    Set rst_F_Act = bds.OpenRecordset("SELECT T_Variable.Variable_Code, T_EtapeMission_VariableAction.MPU1, T_EtapeMission_VariableAction.MPU2, T_EtapeMission_VariableAction.Valeur, Tref_PlageMMAP.PlageMMAP_Module, * FROM ((T_EtapeMission INNER JOIN T_EtapeMission_VariableAction ON T_EtapeMission.ID_EtapeMission = T_EtapeMission_VariableAction.ID_Etape) INNER JOIN T_Variable ON T_EtapeMission_VariableAction.ID_Variable = T_Variable.ID_Variable) INNER JOIN Tref_PlageMMAP ON T_EtapeMission_VariableAction.ID_PlageMMAP = Tref_PlageMMAP.ID_PlageMMAP WHERE (((T_EtapeMission.EtapeMission_Code)='" & ETP_CODE_M & "') AND ((T_EtapeMission.EtapeMission_Indice)='" & ETP_INDICE_M & "'));", dbOpenDynaset)
Set rst_M_Act = bds.OpenRecordset("SELECT T_VariableCB.VariableCB_Code, T_VariableCB.VariableCB_Doublon, T_EtapeMission_VariableCBAction.Valeur, Tref_EquipementCB.EquipementCB_Code, * FROM T_VariableCB INNER JOIN ((T_EtapeMission INNER JOIN T_EtapeMission_VariableCBAction ON T_EtapeMission.ID_EtapeMission = T_EtapeMission_VariableCBAction.ID_Etape) INNER JOIN Tref_EquipementCB ON T_EtapeMission_VariableCBAction.ID_EquipementCB = Tref_EquipementCB.ID_EquipementCB) ON T_VariableCB.ID_VariableCB = T_EtapeMission_VariableCBAction.ID_Variable WHERE (((T_EtapeMission.EtapeMission_Code)='" & ETP_CODE_M & "') AND ((T_EtapeMission.EtapeMission_Indice)='" & ETP_INDICE_M & "'));", dbOpenDynaset)
'----------------------------------------

If rst_M_Act.EOF = True Then
    GoTo GenCodeMission_ExitPopup
End If

iCounter = 1

If rg < 3 Then
    ' Création de l'étape libellé dans le test Mission
    bEnvir = True
    genLabel sequence, rg, "-------------------------------------------------", False, iDest: rg = rg + 1
    If Mid(ETP_CODE_M, 2, 1) = "E" Then
        genLabel sequence, rg, "Mode Mission : " & ModeMission & " (Début)", True, iDest: rg = rg + 1
    ElseIf Mid(ETP_CODE_M, 2, 1) = "S" Then
        genLabel sequence, rg, "Mode Mission : " & ModeMission & " (Fin)", True, iDest: rg = rg + 1
    End If
    genLabel sequence, rg, "-------------------------------------------------", False, iDest: rg = rg + 1
    strContexte = "Etape Mission : " & ETP_CODE_M & " ed. " & ETP_INDICE_M
    
    If Mid(ETP_CODE_M, 2, 1) = "E" Then
        bEnvir = True
        'Lecture de la version de l'appli CB
        genLectureNNModel sequence, rg, 1, sVariable_Version_CB, "FileGlobals.strVersion", False, "", True, iDest: rg = rg + 1
        genLabel sequence, rg, "Version de l'Environnement CB : ", True, iDest, "FileGlobals.strVersion": rg = rg + 1
        genLabel sequence, rg, "Version de la génération : ", True, iDest, Chr(34) & VERSION_BDD & Chr(34): rg = rg + 1
        'Lecture de la version de Teststand
        genVersionTS sequence, rg, "FileGlobals.strVersion", "", iDest: rg = rg + 1
        genLabel sequence, rg, "Version de TestStand : ", True, iDest, "FileGlobals.strVersion": rg = rg + 1
    ElseIf Mid(ETP_CODE_M, 2, 1) = "S" Then
        bEnvir = True
        'A la sortie de la séquence, on libère le handle IHM (appui touches)
        rg = InsertLabel(sequence, iDest, rg, "Var. Entrées = " & intNBVar_E & " | " & "Var. Sorties = " & intNBVar_S)
        If TestStandIHMInited = True Then genCleanUpIHM
    End If
End If
reset_fin_etape = False

    '---------------------------
    '------ ACTION MODEL -------
While Not rst_M_Act.EOF
    sTempoVehicule = 1 'CStr(rst_M_Act.Fields("Vehicule").Value)
    bEnvir = True
    If sTempoVehicule <= 6 Then iNumConsist = 3
    If sTempoVehicule <= 4 Then iNumConsist = 2
    If sTempoVehicule <= 2 Then iNumConsist = 1
    sVariable = rst_M_Act.Fields("VariableCB_Code").Value
    sVariableTempo = sVariable
    
    Set rst_Plage = bds.OpenRecordset("SELECT Tref_PlageVariable.PlageVariable_Code, * FROM Tref_PlageVariable INNER JOIN T_VariableCB ON Tref_PlageVariable.ID_PlageVariable = T_VariableCB.VariableCB_Plage WHERE T_VariableCB.VariableCB_Code='" & sVariable & "';", dbOpenDynaset)
    strPlage = rst_Plage.Fields("PlageVariable_Code").Value
    boolAccesNickname = rst_Plage.Fields("VariableCB_Doublon").Value
    str_EQT = rst_M_Act.Fields("EquipementCB_Code").Value
    sValeur = CStr(rst_M_Act.Fields("Valeur").Value)
            
    If sVariable = "RAZ_ALL_ISLT" Then
        strContexte = " Section=" & sTempoVehicule & " ; Equipement=Raz_All_Islt ; "
        genAppelSéquenceAutreFichier sequence, rg, "Efface les Isolements", "MainSequence", "C:\Teststand\RAZ_ALL_ISLT_MPU.seq", False, False, ""
        rg = rg + 1
    Else
        If (boolAccesNickname = False) Then 'Traitement cas Mnémonique du LV_VEHICLE
            If iNumVehMission = 0 Then
                sInstance = GenCheminVariable(sTempoVehicule, str_EQT, True)
            Else
                sInstance = GenCheminVariable(CStr(iNumVehMission), str_EQT, True)
            End If
            If (sValeur = LCase("U")) Or (sValeur = UCase("u")) Then
                genUnForceModel sequence, rg, sInstance, sTempoVehicule, sVariable, CStr(iCounter), True, StepGroup_Main: rg = rg + 1
            Else
                If strPlage = "Booléen" Then
                    If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureModel sequence, rg, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter), True, iDest: rg = rg + 1
                    Else
                        genForceModel sequence, rg, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter), True, iDest: rg = rg + 1
                    End If
                Else
                    If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureModel sequence, rg, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter), True, iDest: rg = rg + 1
                    Else
                        genForceModel sequence, rg, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter), True, iDest: rg = rg + 1
                    End If
                End If
            End If
        Else
            sVariable = GenVariableModel(sVariable, bRedondance)
            sInstance = ""
                
            If (sValeur = LCase("U")) Or (sValeur = UCase("u")) Then
                genUnForceNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, CStr(iCounter), True, iDest: rg = rg + 1
                If bRedondance = True Then 'Redondance si la var est du type E_BT1
                    bRedondance = False
                    iCounter = iCounter + 1
                    sVariable = "R2" & Mid(sVariable, 3, Len(sVariable) - 2)
                    genUnForceNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, CStr(iCounter), True, iDest: rg = rg + 1
                End If
            Else
                If strPlage = "Booléen" Then
                    If VarToWriteOrForce(sVariableTempo) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter), True, iDest: rg = rg + 1
                    Else
                        genForceNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter), True, iDest: rg = rg + 1
                    End If
                Else
                    If VarToWriteOrForce(sVariableTempo) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter), True, iDest: rg = rg + 1
                    Else
                        genForceNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter), True, iDest: rg = rg + 1
                    End If
                End If
                If bRedondance = True Then 'Redondance si la var est du type E_BT1
                    bRedondance = False
                    iCounter = iCounter + 1
                    sVariable = "R2" & Mid(sVariable, 3, Len(sVariable) - 2)
                    If strPlage = "Booléen" Then
                        If VarToWriteOrForce(sVariableTempo) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                            genEcritureNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter), True, iDest: rg = rg + 1
                        Else
                            genForceNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter), True, iDest: rg = rg + 1
                        End If
                    Else
                        If VarToWriteOrForce(sVariableTempo) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                            genEcritureNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter), True, iDest: rg = rg + 1
                        Else
                            genForceNNModel sequence, rg, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter), True, iDest: rg = rg + 1
                        End If
                    End If
                End If
            End If
        End If
    End If
    If sVariable = "Z_MES_BA" And sValeur = "2" Then
        genAppelSéquenceAutreFichier sequence, rg, "Connection aux taches", "MainSequence", "C:\Teststand\MPU_ON.seq", False, False, "": rg = rg + 1
    End If

    
    
    iCounter = iCounter + 1
    rst_M_Act.MoveNext
Wend
'---------------------------
'------ TEMPO --------------
genWaitTS sequence, rg, nTempo / 1000, iDest: rg = rg + 1
'---------------------------


RanginSeq = rg
        
GenCodeMission_ExitPopup:
    '----------------------------------------
'    rst_F_Act.Close
rst_M_Act.Close
    '------------------------------------------------
Exit Sub

GenCodeMission_Err:
    ' Affiche les informations sur l'erreur.
MsgBox "Numéro d'erreur " & Err.Number & ": " & Err.Description
    ' Reprend l'exécution à l'instruction suivant la ligne où l'erreur s'est produite.
Resume Next
GenCodeMission_Variable_Err:
MsgBox "La variable " & rst_F_Verif.Fields("Variable_Code").Value & " n'a pu etre correctement généré," & vbCr & "La génération a été arrêté !"
End Sub

Public Function VarToWriteOrForce(ByVal strVariable As String) As Boolean
'Fonction qui détermine si la variable Model doit etre écrite ou forcé
'Fonction spécifique car synoptique sur les boutons de ControlBuild
' VarToWriteOrForce=True si Ecriture
' VarToWriteOrForce=False si Forcage
Dim bResultat As Boolean
    
    If bResultat = True Then bResultat = False
    If Mid(strVariable, 1, 3) = "BP_" Then bResultat = True     'Bouton poussoir
    If Mid(strVariable, 1, 4) = "BPL_" Then bResultat = True    'Bouton poussoir lumineux
    If Mid(strVariable, 1, 5) = "BPL1_" Then bResultat = True   'Bouton poussoir lumineux 1
    If Mid(strVariable, 1, 5) = "BPL2_" Then bResultat = True   'Bouton poussoir lumineux 2
    If Mid(strVariable, 1, 4) = "BP1_" Then bResultat = True    'Bouton poussoir 1
    If Mid(strVariable, 1, 4) = "BP2_" Then bResultat = True    'Bouton poussoir 2
    If Mid(strVariable, 1, 2) = "P_" Then bResultat = True      'Pédale
    If Mid(strVariable, 1, 2) = "Z_" Then bResultat = True      'Commutateur
    If Mid(strVariable, 1, 3) = "ZP_" Then bResultat = True     'Commutateur de Process
    If Mid(strVariable, 1, 2) = "K_" Then bResultat = True
    If Mid(strVariable, 1, 3) = "RB_" Then bResultat = True     'Robinet
    If Mid(strVariable, 1, 4) = "VV1_" Then bResultat = True
    If Mid(strVariable, 1, 3) = "Q1_" Then bResultat = True
    If Mid(strVariable, 1, 3) = "Q2_" Then bResultat = True
    If Mid(strVariable, 1, 3) = "DF_" Then bResultat = True     'Défaut

VarToWriteOrForce = bResultat
End Function

Public Function GenVariableModel(ByVal strVarModel As String, ByVal bRedondance As Boolean) As String
Dim strTemp As String
Dim strVariable, sDirection, sNumero As String

    'Mémorise le nom de la variable passée
    strVariable = strVarModel
    
    'Isole les 5 premiers caractères de la variable
    strTemp = Left(strVariable, 5)
    'Si ces 5 premières lettres correspondent à une égalité ci-dessous alors
    If strTemp = "E_BT1" Or strTemp = "S_BT1" Or strTemp = "E_BT2" Or strTemp = "S_BT2" Then
        'on isole la lettre permettant de déterminer si elle est en Entrée ou en Sortie
        'on isole son chiffre
        sDirection = Left(strVariable, 1)
        sNumero = Mid(strVariable, 5, 1)
        'Reconstruction de la variable comme : R1EL_nomVariable
        strVariable = "R" & sNumero & sDirection & "L_" & Mid(strVarModel, 7, Len(strVarModel) - 6)
    ElseIf Left(strVariable, 10) = "E_BT_AL24_" Then    'Cas particulier pour E_BT_AL24_1 ou 2
        sNumero = Right(strVariable, 1)
        strVariable = "R" & sNumero & "EL_AL24_" & sNumero
    End If
    'Renvoi de la variable comme résultat de la fonction
    GenVariableModel = strVariable
End Function

Public Function GenCheminVariable(ByVal strNumVehicle As String, ByVal strEqt_Code As String, bEnv As Boolean) As String
'Si bEnv=true Alors var Environnemment ; Si bEnv=False Alors autre
Dim iCounter As Integer
Dim sTemp, sTempEnv As String
Dim sNomTempoEqt, stempequip As String
Dim iLevel As Integer   'Niveau dans le chemin de l'instance
Dim rstChemin As Recordset
    
Set db_ihm = CurrentDb

If bEnv = True Then 'Variable d'environnement
    If strNumVehicle = 1 Or strNumVehicle = 3 Then
        sTempEnv = "SECTION_1/"
    ElseIf strNumVehicle = 2 Or strNumVehicle = 4 Then
        sTempEnv = "SECTION_2/"
    Else
        sTempEnv = "SECTION_1/" 'Section par défaut
    End If
    Set rstChemin = db_ihm.OpenRecordset("SELECT EquipementCB_Code,EquipementCB_Prefixe,EquipementCB_AdresseNommage FROM Tref_EquipementCB WHERE EquipementCB_Code=" & Chr(34) & strEqt_Code & Chr(34))
    If Not (rstChemin.EOF) Then
        sNomTempoEqt = rstChemin.Fields("EquipementCB_Prefixe")
        sTemp = sTempEnv & sNomTempoEqt
        If rstChemin.Fields("EquipementCB_AdresseNommage") <> "" Then
            sTemp = sTemp & "/" & rstChemin.Fields("EquipementCB_AdresseNommage")
        End If
    End If
    
Else 'bEnv=false (Embarquee)
    sTemp = CheminEmbarque
    Set rstChemin = db_ihm.OpenRecordset("SELECT FBS_Fonction,FBS_Titre FROM Tref_FBS WHERE FBS_Fonction=" & Chr(34) & strEqt_Code & Chr(34))
    If Not (rstChemin.EOF) Then
        If Left(rstChemin.Fields("FBS_Fonction"), 5) = "MODES" Then
            sTemp = "Embarquee/" & rstChemin.Fields("FBS_Titre")
        ElseIf Left(rstChemin.Fields("FBS_Fonction"), 5) = "SYSTEM" Then
            sTemp = "Embarquee/Systeme/" & rstChemin.Fields("FBS_Titre")
        Else
            sTemp = sTemp & "FBS/" & rstChemin.Fields("FBS_Titre")
        End If
    End If
End If

db_ihm.Close
Set db_ihm = Nothing

GenCheminVariable = sTemp
End Function

Public Function GenCheminFolio(ByVal Instance As String, ByVal variable_Folio As String) As String
Dim Resultat As String
Dim Page As String
Dim Folio As String
    If Left(Right(variable_Folio, 5), 4) = "_CAB" Then
        Folio = "K" & Right(variable_Folio, 9)    'K01_CAB1
        Folio = Format(Folio, ">")          'Force le format majuscule
        Page = Left(Folio, 3)               'K01
        Resultat = Instance & "/" & Page & "/" & Folio
        
    ElseIf Left(Right(variable_Folio, 6), 4) = "_ACU" Then
        Folio = "K" & Right(variable_Folio, 10)    'r01A0_ACU11
        Folio = Format(Folio, ">")          'Force le format majuscule
        Page = Left(Folio, 3)               'b01
        Resultat = Instance & "/" & Page & "/" & Folio
    
    ElseIf Left(Right(variable_Folio, 4), 1) = "B" Then
        Folio = "K" & Right(variable_Folio, 3)    'K01
        Folio = Format(Folio, ">")          'Force le format majuscule => K01
        Page = Left(Folio, 3)               'K01
        Resultat = Instance & "/" & Page & "/" & Folio
    Else
        Resultat = "ERROR Folio"
    End If

GenCheminFolio = Resultat
End Function

Public Function GenVariableFolio(ByVal Variable As String, ByVal Instance As String) As String
Dim Resultat As String

If Left(Right(Instance, 4), 3) = "CAB" Then
    Resultat = Left(Variable, (Len(Variable) - 10))
ElseIf Left(Right(Instance, 5), 3) = "ACU" Then
    Resultat = Left(Variable, (Len(Variable) - 11))
ElseIf Right(Variable, 3) = Right(Instance, 3) Then
    Resultat = Mid(Variable, 1, Len(Variable) - 5)
Else
    Resultat = "ERROR variable"
End If

GenVariableFolio = Resultat
End Function

Public Function GenVariableFip_TEMPO(ByVal strNomVariableFip As String, ByVal strNumVehicle As String, Optional strNumEtape As String, Optional sText As String) As Integer

Dim rst, rst2 As Recordset
Dim VersionSTR As String
Dim sPositionTempo As String
Dim intError As Integer

    intError = 0
    sPositionTempo = ""
    VersionSTR = Form_GenerationCode.LM_STR.Value
    
'    Set rst = CurrentDb.OpenRecordset("SELECT LISTE_TEMPO_DEF.CODE_TEMP, LISTE_CODE_TEMPO_DEF.STR, LISTE_CODE_TEMPO_DEF.CODE_VAR_TEMP,LISTE_CODE_TEMPO_DEF.TYPE_TEMP_DEF,LISTE_CODE_TEMPO_DEF.CODE_POS_TEMPO FROM  LISTE_TEMPO_DEF INNER JOIN LISTE_CODE_TEMPO_DEF ON LISTE_TEMPO_DEF.CODE_VAR_TEMP = LISTE_CODE_TEMPO_DEF.CODE_VAR_TEMP WHERE LISTE_TEMPO_DEF.CODE_TEMP=" & Chr(34) & strNomVariableFip & Chr(34) & "and LISTE_CODE_TEMPO_DEF.STR=" & Chr(34) & VersionSTR & Chr(34))
'    If rst.EOF Then
'        intError = 1
'        sText = "La variable " & strNomVariableFip & " n'existe pas dans la table des correspondances LISTE_TEMPO_DEF"
'        If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
'    Else
'        rst.MoveFirst
'        While Not rst.EOF
'            If rst.Fields("TYPE_TEMP_DEF") = "TEMPORISE" Then
'                sPositionTempo = rst.Fields("CODE_POS_TEMPO")
'            End If
'            rst.MoveNext
'        Wend
        If sPositionTempo <> "" Then
            sNomVarTemp = "t_ini"   'n_ini
            strNomEqtTemp = "@\\" & "ramsvg_aff\\tab_surv_var[" & sPositionTempo & "].comptage_def."
        Else
            intError = 2
            sText = "La temporisation " & strNomVariableFip & " n'a pas de position valide"
            If strNumEtape <> "" Then sText = sText & " pour l'étape " & strNumEtape
        End If
'    End If
GenVariableFip_TEMPO = intError
End Function

Public Sub GenCodeEtape(bds As Database, NumEtape As Integer, Id_etape As Long, ETP_CODE As String, ETP_INDICE As String, nTempo As Single, Test_Designation As String, SKIP As Boolean, SKIP_TEST As Boolean, nTabConfigTrain() As Integer, ByVal strAddress_Fip As String, ByVal strAddress_Ihm As String, ByVal ID_TEST As Long, ByVal Conf_PR As String, ByVal Conf_banc As String)
  
  '  On Error GoTo GenCodeEtape_Err
  
Dim rstAct1, rstAct2, rstVerif1, rstVerif2, rstRappel As Recordset
Dim rst_F_Act, rst_F_Verif, rst_M_Act, rst_M_Verif As Recordset
Dim rst_F_EQT, rst_M_EQT As Recordset   'tables pointant sur l'equipement
Dim rst_Plage As Recordset              'table pointant sur la plage de la variable
Dim rst_Etape As Recordset
Dim sInstance, sVariable, sVariableTempo, sValeur, str_EQT, sVariablePrec, strTolerance As String
Dim sNomPop, sTitrePop, sMessagePop As String
Dim sVehicule, sTempoVehicule, sSection, strPlage, sTemp, sTexteError As String
Dim reset_fin_etape, boolPresentRP, boolAccesNickname, boolSTR As Boolean
Dim iNumConsist, iCounter, iPositionCherche, iHandleMmap As Integer
Dim VersionVariable As Variant
Dim sFichierSequence, sModuleMmap As String
Dim bMPU1, bMPU2 As Boolean             'Lecture/Ecriture autorisée
Dim intExistePrecond As Integer         'Une précondition existe dans le pas
Dim strEquipement As String             'Nom de l'Equipement tempo
Dim strNomEcran As String               'Nom de l'écran à vérifier
Dim bEnvir As Boolean
Dim passage As Integer

'Stocke les données de chaque Etape
Set rst_Etape = bds.OpenRecordset("SELECT T_Etape.Etape_Designation,T_Etape.Etape_ActionLoco,T_Etape.Etape_Commentaire, * FROM T_Etape WHERE T_Etape.ID_Etape=" & Id_etape & ";", dbOpenDynaset)

' Création de l'étape dans le test
Dim strContexte As String

If rg_EtapeDansSequence = 0 Then
'Ecrit sur le rapport la dégnation du Test et les exigences en haut de page
    bEnvir = True
    Set rs_Tempo = bds.OpenRecordset("SELECT T_Test_Exigences.ID_Test, Tref_Exigences.Exigences_Titre, Tref_Exigences.Exigences_Sil FROM Tref_Exigences INNER JOIN T_Test_Exigences ON Tref_Exigences.ID_Exigences = T_Test_Exigences.ID_Exigences WHERE T_Test_Exigences.ID_Test=" & ID_TEST & ";")
    genLabel sequence, rg_EtapeDansSequence, "-------------------------------------------------", True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    genLabel sequence, rg_EtapeDansSequence, Test_Designation, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    Do Until rs_Tempo.EOF
        genLabel sequence, rg_EtapeDansSequence, "Exigence : " & rs_Tempo.Fields("Exigences_Titre") & " - SIL = " & rs_Tempo.Fields("Exigences_Sil"), True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        rs_Tempo.MoveNext
    Loop
    genLabel sequence, rg_EtapeDansSequence, "-------------------------------------------------", True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    If SKIP_TEST = True Then
        genLabel sequence, rg_EtapeDansSequence, "-- Test désactivé pour la Non Reg --", True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        GoTo GenCodeEtape_Exit
    End If
    Set rs_Tempo = Nothing
End If

If SKIP_TEST = True Then
    GoTo GenCodeEtape_Exit
End If

bEnvir = True
'     Endroit où va se trouver la précondition de l'étape
genLabel sequence, rg_EtapeDansSequence, "-------------------------------------------------", False, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
genLabel sequence, rg_EtapeDansSequence, NumEtape & "_Etape : " & ETP_CODE, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
genLabel sequence, rg_EtapeDansSequence, "-------------------------------------------------", False, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
reset_fin_etape = False

If SKIP = True Then
'L'étape est désactivée --> Affichage d'un message
    genLabel sequence, rg_EtapeDansSequence, "-- Etape désactivée --", True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    GoTo GenCodeEtape_Exit
End If

' Ouvre les tables correspondant aux étapes (Action et verif pour l'Embarqué et l'Environnement)
Set rst_F_Act = bds.OpenRecordset("SELECT T_Variable.ID_Variable, T_Variable.Variable_Code, T_Etape_VariableAction.MPU1, T_Etape_VariableAction.MPU2, T_Etape_VariableAction.Valeur, T_Variable.Variable_Plage, T_Variable.Variable_Bloc, T_Variable.Variable_Chemin,T_Variable.Variable_STR,T_Etape_VariableAction.Ordre, * FROM T_Variable INNER JOIN (T_Etape INNER JOIN T_Etape_VariableAction ON T_Etape.ID_Etape = T_Etape_VariableAction.ID_Etape) ON T_Variable.ID_Variable = T_Etape_VariableAction.ID_Variable WHERE (((T_Etape.ID_Etape)=" & Id_etape & ")) ORDER BY T_Etape_VariableAction.Ordre;", dbOpenDynaset)
Set rst_M_Act = bds.OpenRecordset("SELECT T_VariableCB.VariableCB_Code, T_VariableCB.VariableCB_Doublon, T_Etape_VariableCBAction.Valeur, Tref_EquipementCB.EquipementCB_Code,T_Etape_VariableCBAction.Ordre, * FROM T_VariableCB INNER JOIN ((T_Etape INNER JOIN T_Etape_VariableCBAction ON T_Etape.ID_Etape = T_Etape_VariableCBAction.ID_Etape) INNER JOIN Tref_EquipementCB ON T_Etape_VariableCBAction.ID_EquipementCB = Tref_EquipementCB.ID_EquipementCB) ON T_VariableCB.ID_VariableCB = T_Etape_VariableCBAction.ID_VariableCB WHERE (((T_Etape.ID_Etape)=" & Id_etape & ")) ORDER BY T_Etape_VariableCBAction.Ordre;", dbOpenDynaset)
Set rst_F_Verif = bds.OpenRecordset("SELECT T_Variable.ID_Variable, T_Variable.Variable_Code, T_Etape_VariableCheck.MPU1, T_Etape_VariableCheck.MPU2, T_Etape_VariableCheck.Valeur, T_Variable.Variable_Plage, T_Variable.Variable_Bloc, T_Variable.Variable_Chemin,T_Variable.Variable_STR,T_Etape_VariableCheck.Ordre, * FROM T_Variable INNER JOIN (T_Etape INNER JOIN T_Etape_VariableCheck ON T_Etape.ID_Etape = T_Etape_VariableCheck.ID_Etape) ON T_Variable.ID_Variable = T_Etape_VariableCheck.ID_Variable WHERE (((T_Etape.ID_Etape)=" & Id_etape & ")) ORDER BY T_Etape_VariableCheck.Ordre;", dbOpenDynaset)
Set rst_M_Verif = bds.OpenRecordset("SELECT T_VariableCB.VariableCB_Code, T_Etape_VariableCBCheck.Valeur, Tref_EquipementCB.EquipementCB_Code,T_Etape_VariableCBCheck.Ordre, * FROM Tref_EquipementCB INNER JOIN (T_VariableCB INNER JOIN (T_Etape INNER JOIN T_Etape_VariableCBCheck ON T_Etape.ID_Etape = T_Etape_VariableCBCheck.ID_Etape) ON T_VariableCB.ID_VariableCB = T_Etape_VariableCBCheck.ID_VariableCB) ON Tref_EquipementCB.ID_EquipementCB = T_Etape_VariableCBCheck.ID_EquipementCB WHERE (((T_Etape.ID_Etape)=" & Id_etape & ")) ORDER BY T_Etape_VariableCBCheck.Ordre;", dbOpenDynaset)

'------------------------------------------------------
'------ ACTION MVB ou local ---------
'Gestion des MPU inconnu !!! Pas de redondance
iCounter = 1
While Not rst_F_Act.EOF
    str_PreCond = "" 'Vide le texte de précondition car variable global
    bEnvir = False
    intExistePrecond = 0 ' ----------------------------Recherche_Precond(ETP_CODE, ETP_INDICE, True, True)
    If intExistePrecond > 0 Then
        Ajout_Precond intExistePrecond, True, True, StepGroup_Main
    End If
    sVariable = rst_F_Act.Fields("Variable_Code").Value
    sValeur = CStr(rst_F_Act.Fields("Valeur").Value)
    sValeur = RemplacePointVirgule(sValeur)
    VersionVariable = rst_F_Act.Fields("Variable_STR").Value
    str_EQT = rst_F_Act.Fields("Variable_Bloc").Value 'nom du bloc
    strPlage = rst_F_Act.Fields("Variable_Plage").Value
    
    passage = 0
    '-------------------
    sSection = CStr(rst_F_Act.Fields("Vehicule").Value)
    sTempoVehicule = Def_Section(sSection, Conf_PR, Conf_banc, passage)
    '-------------------
    strContexte = " Section=" & sTempoVehicule & " ; Fonction=" & str_EQT & " ; "
    sInstance = GenCheminVariable(sTempoVehicule, str_EQT, False)
    If rst_F_Act.Fields("Variable_Chemin").Value <> "" Then
        sInstance = sInstance & "/" & rst_F_Act.Fields("Variable_Chemin").Value
        strContexte = " Section=" & sTempoVehicule & " ; Fonction=" & str_EQT & "/" & rst_F_Act.Fields("Variable_Chemin").Value & " ; "
    End If
    If (sValeur = LCase("U")) Or (sValeur = UCase("u")) Then
        genUnForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, CStr(iCounter) & strContexte, False, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    Else
       If strPlage = 6 Then
            genForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter) & strContexte, False, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
       Else
           If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
               genEcritureModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter) & strContexte, False, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
           Else
               genForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter) & strContexte, False, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
           End If
       End If
    End If
    iCounter = iCounter + 1
    rst_F_Act.MoveNext
Wend
'------------------------------------------------------
'------ ACTION MODEL -------
While Not rst_M_Act.EOF
    str_PreCond = "" 'Vide le texte de précondition car variable global
    intExistePrecond = 0 ' ----------------------- Recherche_Precond(ETP_CODE, ETP_INDICE, False, True)
    bEnvir = True
    If intExistePrecond > 0 Then
        Ajout_Precond intExistePrecond, False, True, StepGroup_Main
    End If
    
    passage = 0
    '-------------------
    sSection = CStr(rst_M_Act.Fields("Vehicule").Value)
    sTempoVehicule = Def_Section(sSection, Conf_PR, Conf_banc, passage)
    '-------------------
    str_EQT = rst_M_Act.Fields("EquipementCB_Code").Value
    sVariable = rst_M_Act.Fields("VariableCB_Code").Value
    sValeur = CStr(rst_M_Act.Fields("Valeur").Value)
    
    If (sVariable = "POPUP") Then
        sNomPop = "Pause"
        sTitrePop = "Pause"
        sMessagePop = rst_Etape.Fields("Etape_ActionLoco").Value
        strContexte = " Pause"
        genPopup sequence, rg_EtapeDansSequence, CStr(iCounter) & strContexte, sTitrePop, sMessagePop, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    ElseIf Left(sVariable, 2) = "TF" Then
        'PC_IHM
        sMessagePop = "Action sur une touche " & sVariable
        genTF sequence, rg_EtapeDansSequence, sValeur, sVariable, sMessagePop, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    ElseIf Left(sVariable, 3) = "CHK" Then
        sTexteError = "Il est impossible de checker un écran en action " & ETP_CODE & " pour " & sVariable
        ArrTable_ReportError(iCounterError) = sTexteError: iCounterError = iCounterError + 1
    ElseIf ((Left(sVariable, 3) = "CC_") Or (Left(sVariable, 4) = "CC1_") Or (Left(sVariable, 4) = "CC2_") Or (Left(sVariable, 9) = "DJ_SI_CPR")) Then ''And (Surveillance_HT = True) Then
         'Accès au CC via les instances, les 4 derniers caractères définissent le folio
        str_EQT = rst_M_Act.Fields("EquipementCB_Code").Value
        sInstance = GenCheminVariable(sTempoVehicule, str_EQT, True)
        sInstance = GenCheminFolio(sInstance, sVariable)
        sVariableTempo = GenVariableFolio(sVariable, sInstance)
        sValeur = CStr(rst_M_Act.Fields("Valeur").Value)
        sValeur = RemplacePointVirgule(sValeur)
        strContexte = " Section=" & sTempoVehicule & " ; Equipement=" & str_EQT & " ; "
        If (sValeur = LCase("U")) Or (sValeur = UCase("u")) Then
            genUnForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariableTempo, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        Else
            genForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariableTempo, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        End If
    ElseIf Left(sVariable, 3) = "ZP_" And Right(sVariable, 4) = "_OFF" And ((Left(str_EQT, 3) = "ACU") Or (Left(str_EQT, 3) = "TCU") Or (Left(str_EQT, 3) = "BCU") Or (Left(str_EQT, 4) = "RIOM") Or (Left(str_EQT, 3) = "MPU")) Then
    'Gestion du cas particulier où l'on arrête ou démarre un équipement sur le réseau
        str_EQT = rst_M_Act.Fields("EquipementCB_Code").Value
        sValeur = CStr(rst_M_Act.Fields("Valeur").Value)
 '------------------------------------
        If Left(str_EQT, 3) = "MPU" Then
            strEquipement = Right(str_EQT, 1)
            If sValeur = "1" Then
            'Arret de l'équipement MPU
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "MPU OFF", "MainSequence", "C:\Teststand\Stop MPUx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            Else
            'Démarrage de l'équipement MPU
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "MPU ON", "MainSequence", "C:\Teststand\Start MPUx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            End If
 '------------------------------
        ElseIf Left(str_EQT, 4) = "RIOM" Then
            strEquipement = Right(str_EQT, 2)
            If sValeur = "1" Then
            'Arret de l'équipement RIOM sur le réseau
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "RIOM OFF", "MainSequence", "C:\Teststand\Stop RIOMxx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            Else
            'Démarrage de l'équipement RIOM sur le réseau
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "RIOM ON", "MainSequence", "C:\Teststand\Start RIOMxx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            End If
 '------------------------------
        ElseIf Left(str_EQT, 3) = "TCU" Then
            strEquipement = Right(str_EQT, 2)
            If sValeur = "1" Then
            'Arret de l'équipement RIOM sur le réseau
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "TCU OFF", "MainSequence", "C:\Teststand\Stop TCUx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            Else
            'Démarrage de l'équipement RIOM sur le réseau
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "TCU ON", "MainSequence", "C:\Teststand\Start TCUx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            End If
 '------------------------------
        ElseIf Left(str_EQT, 3) = "ACU" Then
            strEquipement = Right(str_EQT, 2)
            If sValeur = "1" Then
            'Arret de l'équipement RIOM sur le réseau
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "ACU OFF", "MainSequence", "C:\Teststand\Stop ACUx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            Else
            'Démarrage de l'équipement RIOM sur le réseau
                    genAppelSéquenceAutreFichier sequence, rg_EtapeDansSequence, "ACU ON", "MainSequence", "C:\Teststand\Start ACUx.seq", False, False, "", strEquipement, sTempoVehicule
                    rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            End If
        End If
'-----------------------------
    Else
    
        '' bug pour la saisie de la variable dans la table, NE jamais effacer
        If sVariable = "ZP_Ucat" Then sVariable = "ZP_UCAT"
        If sVariable = "R_TST_Ucat" Then sVariable = "R_TST_UCAT"
        If sVariable = "N_TST_Fcat" Then sVariable = "N_TST_FCAT"
        
        
        Set rst_Plage = bds.OpenRecordset("SELECT Tref_PlageVariable.PlageVariable_Code, * FROM Tref_PlageVariable INNER JOIN T_VariableCB ON Tref_PlageVariable.ID_PlageVariable = T_VariableCB.VariableCB_Plage WHERE T_VariableCB.VariableCB_Code='" & sVariable & "';", dbOpenDynaset)
        strPlage = rst_Plage.Fields("PlageVariable_Code").Value
        boolAccesNickname = rst_Plage.Fields("VariableCB_Doublon").Value
        str_EQT = rst_M_Act.Fields("EquipementCB_Code").Value
        sValeur = CStr(rst_M_Act.Fields("Valeur").Value)
        sValeur = RemplacePointVirgule(sValeur)
        strContexte = " Section=" & sTempoVehicule & " ; Equipement=" & str_EQT & " ; "
        
        If (boolAccesNickname = False) Then 'Traitement cas Mnémonique du LV_VEHICLE
            sInstance = GenCheminVariable(sTempoVehicule, str_EQT, True)
            
            '****** Modif du 28/05/2009 ** Cohérence avec le filtrage des ACU
            If Left(str_EQT, 3) = "ACU" Then
                sVariable = VarACU(sVariable, str_EQT)
            End If
            '******
            
            If (sValeur = LCase("U")) Or (sValeur = UCase("u")) Then
                genUnForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            Else
                If strPlage = "Booléen" Then
                    If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    Else
                        genForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    End If
                Else
                    If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    Else
                        genForceModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    End If
                End If
            End If
        Else
            sVariable = GenVariableModel(sVariable, bRedondance)
            
            sInstance = ""
            
            If (sValeur = LCase("U")) Or (sValeur = UCase("u")) Then
                genUnForceNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sVariable, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                If bRedondance = True Then 'Redondance si la var est du type E_BT1
                    bRedondance = False
                    iCounter = iCounter + 1
                    sVariable = "R2" & Mid(sVariable, 3, Len(sVariable) - 2)
                    genUnForceNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                End If
            Else
                If strPlage = "Booléen" Then
                    If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    Else
                        genForceNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    End If
                Else
                    If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                        genEcritureNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    Else
                        genForceNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    End If
                End If
                If bRedondance = True Then 'Redondance si la var est du type E_BT1
                    bRedondance = False
                    iCounter = iCounter + 1
                    sVariable = "R2" & Mid(sVariable, 3, Len(sVariable) - 2)
                    If strPlage = "Booléen" Then
                        If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                            genEcritureNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                        Else
                            genForceNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                        End If
                    Else
                        If VarToWriteOrForce(sVariable) Then   'Détermine si la variable doit etre Forcée ou Ecrite
                            genEcritureNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                        Else
                            genForceNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                        End If
                    End If
                End If
            End If
        End If
        
    End If
        
    iCounter = iCounter + 1
    rst_M_Act.MoveNext
Wend
CompteVar_ES False, iCounter - 1
'------------------------------------------------------
'------ TEMPO --------------
    genWaitTS sequence, rg_EtapeDansSequence, nTempo / 1000, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
'------------------------------------------------------
'------ VERIF MVB ou local ----------
iCounter = 1
While Not rst_F_Verif.EOF
    bEnvir = False
    sVariable = rst_F_Verif.Fields("Variable_Code").Value
    strTolerance = ArrTable_VarPlage(5, Trouve_Tolerance(sVariable))
    passage = 0
    '-------------------
    sSection = CStr(rst_F_Verif.Fields("Vehicule").Value)
    sTempoVehicule = Def_Section(sSection, Conf_PR, Conf_banc, passage)
    '-------------------
 '   If CStr(rst_F_Verif.Fields("Vehicule").Value) = "2" Then
 '   '' Cas loco 2 pour l'UM
 '   Else
    
        If Left(sVariable, 2) = "TF" Then
            sTexteError = "Il est impossible de checker l'appui sur une touche " & ETP_CODE & " pour " & sVariable
            ArrTable_ReportError(iCounterError) = sTexteError: iCounterError = iCounterError + 1
        ElseIf (sVariable = "POPUP") And (iCounter = 1) Then
            sTexteError = "Il est impossible d'afficher un popup en vérif " & ETP_CODE & " pour " & sVariable
            ArrTable_ReportError(iCounterError) = sTexteError: iCounterError = iCounterError + 1
        Else
            VersionVariable = rst_F_Verif.Fields("Variable_STR").Value
            str_EQT = rst_F_Verif("Variable_Bloc").Value   'Nom de la fonction
            
            'Détermine si la variable fait partie de la STR à générer
            If Calcul_VarDansSTR = True Then
            boolSTR = VariableDansSTR(VersionVariable, GNumSTRaGenerer, ETP_CODE, sVariable, rst_F_Verif.Fields("T_Variable.ID_Variable").Value)
            End If
            DoEvents
            
            strPlage = rst_F_Verif.Fields("Variable_Plage").Value
            sValeur = CStr(rst_F_Verif.Fields("Valeur").Value)
            sValeur = RemplacePointVirgule(sValeur)
            strContexte = " Section=" & sTempoVehicule & " ; Fonction=" & str_EQT & " ; "
            
            sInstance = GenCheminVariable(sTempoVehicule, str_EQT, False)
            If rst_F_Verif.Fields("Variable_Chemin").Value <> "" Then
                sInstance = sInstance & "/" & rst_F_Verif.Fields("Variable_Chemin").Value  'Ajoute le chemin apres le code fonction
                strContexte = " Section=" & sTempoVehicule & " ; Fonction=" & str_EQT & "/" & rst_F_Verif.Fields("Variable_Chemin").Value & " ; "
            End If
            
            If strPlage = 6 Then
                genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter) & strContexte, False, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            ElseIf (strPlage = 21) And (Left(sVariable, 2) = "R_") Then
                genTestAnaModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, CStr(iCounter) & strContexte, False, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            Else
                genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter) & strContexte, False, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            End If
        End If
    'End If
    iCounter = iCounter + 1
    rst_F_Verif.MoveNext
Wend
'------------------------------------------------------
'------ VERIF MODEL --------
While Not rst_M_Verif.EOF
    bEnvir = True
    sVariable = rst_M_Verif.Fields("VariableCB_Code").Value
    
    passage = 0
    '-------------------
    sSection = CStr(rst_M_Verif.Fields("Vehicule").Value)
    sTempoVehicule = Def_Section(sSection, Conf_PR, Conf_banc, passage)
    '-------------------
        If Left(sVariable, 3) = "CHK" Then
            'PC_IHM : Appel d'une séquence de vérif des écran ( Fichier sur PC_IHM)
            sMessagePop = ""
            sFichierSequence = ""
            
            If Left(sVariable, 7) = "CHK_TF4" Then
                sFichierSequence = "IHM_CHK_KEY"
            ElseIf Left(sVariable, 7) = "CHK_TF2" Then
                sFichierSequence = "IHM_CHK_NAVIGATION"
            ElseIf Left(sVariable, 9) = "CHK_1000_" Then
                sFichierSequence = "IHM_CHK_ETD_MAIN"
            ElseIf Left(sVariable, 9) = "CHK_3000_" Then
                sFichierSequence = "IHM_CHK_TDD_MAIN"
            ElseIf Left(sVariable, 5) = "CHK_S" Then
                sFichierSequence = "IHM_CHK_BANDEAU_STATUT"
            ElseIf Left(sVariable, 9) = "CHK_B" Then
                sFichierSequence = "IHM_CHK_BARRE_MENU"
            ElseIf Left(sVariable, 9) = "CHK_001_" Then
                sFichierSequence = "IHM_CHK_SEL_LANGUE"
            ElseIf Left(sVariable, 8) = "CHK_100_" Then
                sFichierSequence = "IHM_CHK_SYST"
            ElseIf Left(sVariable, 8) = "CHK_101_" Then
                sFichierSequence = "IHM_CHK_PANTO"
            ElseIf Left(sVariable, 8) = "CHK_300_" Then
                sFichierSequence = "IHM_CHK_STATUT_TRAIN"
            ElseIf Left(sVariable, 8) = "CHK_301_" Then
                sFichierSequence = "IHM_CHK_STATUT_HT"
            ElseIf Left(sVariable, 8) = "CHK_ERR_" Then
                sFichierSequence = "IHM_CHK_ERREUR_ECRAN"
            ElseIf Left(sVariable, 8) = "CHK_ATT_" Then
                sFichierSequence = "IHM_CHK_ATT_COM"
            ElseIf Left(sVariable, 10) = "CHK_PERTE_" Then
                sFichierSequence = "IHM_CHK_PERTE_COM"
            ElseIf Right(sVariable, 5) = "DTPLG" Then
                sFichierSequence = "IHM_CHK_DTPLG"
            End If
                
            sMessagePop = "Vérification écran : " & sVariable
            genAppelSéquencePCIHM sequence, rg_EtapeDansSequence, sMessagePop, sVariable, sFichierSequence, "", StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        ElseIf ((Left(sVariable, 3) = "CC_") Or (Left(sVariable, 9) = "DJ_SI_CPR")) Then '' And (Surveillance_HT = True) Then
            'Accès au CC via les instances, les 4 derniers caractères définissent le folio
            str_EQT = rst_M_Verif("EquipementCB_Code").Value
            sInstance = GenCheminVariable(sTempoVehicule, str_EQT, True)
            sInstance = GenCheminFolio(sInstance, sVariable)
            sVariableTempo = GenVariableFolio(sVariable, sInstance)
            ''sVariableTempo = Mid(sVariable, 1, Len(sVariable) - 5)
            sValeur = CStr(rst_M_Verif.Fields("Valeur").Value)
            sValeur = RemplacePointVirgule(sValeur)
            strContexte = " Section=" & sTempoVehicule & " ; Equipement=" & str_EQT & " ; "
            genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariableTempo, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        ElseIf ((Left(sVariable, 8) = "Q_TH_CPR") Or (Left(sVariable, 6) = "C1_CPR") Or (Left(sVariable, 6) = "TH_CPR")) Then ''Rustine M38 (Maroc) car variable non remontée dans l'arbre
            str_EQT = rst_M_Verif.Fields("EquipementCB_Code").Value
            sValeur = CStr(rst_M_Verif.Fields("Valeur").Value)
            sValeur = RemplacePointVirgule(sValeur)
            sInstance = GenCheminVariable(sTempoVehicule, str_EQT, True)
            strContexte = " Section=" & sTempoVehicule & " ; Equipement=" & str_EQT & " ; "
            If (Left(sVariable, 8) = "Q_TH_CPR") Then
                sInstance = sInstance & "/m38"
                sVariableTempo = sVariable
                genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariableTempo, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            ElseIf (Left(sVariable, 6) = "C1_CPR") Then
                If Right(sVariable, 1) = "2" Then
                    sInstance = sInstance & "/m38/M38B"
                Else
                    sInstance = sInstance & "/m38/M38A"
                End If
                sVariableTempo = sVariable
                genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariableTempo, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            ElseIf (Left(sVariable, 6) = "TH_CPR") Then
                sInstance = sInstance & "/m38/M38C"
                sVariableTempo = sVariable
                genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariableTempo, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main, str_PreCond: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
            End If
        ElseIf rst_M_Verif("EquipementCB_Code").Value = "PGM" Then
      '''      genAppelSequencePGM sequence, rg_EtapeDansSequence, "PGM = " & Str(sVariable), "MainSequence", "C:\TestStand\PGM\" & sVariable & ".seq", False, False, "", sTempoVehicule
            genAppelSequencePGM sequence, rg_EtapeDansSequence, "PGM = " & sVariable, "MainSequence", "C:\TestStand\PGM\" & sVariable & ".seq", False, False, "", sTempoVehicule
            rg_EtapeDansSequence = rg_EtapeDansSequence + 1
        Else
            If sTempoVehicule <= 6 Then iNumConsist = 3
            If sTempoVehicule <= 4 Then iNumConsist = 2
            If sTempoVehicule <= 2 Then iNumConsist = 1
            Set rst_Plage = bds.OpenRecordset("SELECT Tref_PlageVariable.PlageVariable_Code, * FROM Tref_PlageVariable INNER JOIN T_VariableCB ON Tref_PlageVariable.ID_PlageVariable = T_VariableCB.VariableCB_Plage WHERE T_VariableCB.VariableCB_Code='" & sVariable & "';", dbOpenDynaset)
            strPlage = rst_Plage.Fields("PlageVariable_Code").Value
            boolAccesNickname = rst_Plage.Fields("VariableCB_Doublon").Value
            str_EQT = rst_M_Verif("EquipementCB_Code").Value
            sValeur = CStr(rst_M_Verif.Fields("Valeur").Value)
            sValeur = RemplacePointVirgule(sValeur)
            strContexte = " Section=" & sTempoVehicule & " ; Equipement=" & str_EQT & " ; "
            
            If (boolAccesNickname = False) Then 'Traitement cas Mnémonique du LV_VEHICLE
                sInstance = GenCheminVariable(sTempoVehicule, str_EQT, True)
                
                '****** Modif du 28/05/2009 ** Cohérence avec le filtrage des ACU
                If Left(str_EQT, 3) = "ACU" Then
                    sVariable = VarACU(sVariable, str_EQT)
                End If
                '******
                
                If strPlage = "Booléen" Then
                    genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                Else
                    genTestModel sequence, rg_EtapeDansSequence, sInstance, sTempoVehicule, sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                End If
            Else
                sVariable = GenVariableModel(sVariable, bRedondance)
                
                sInstance = ""
                
                If strPlage = "Booléen" Then
                    genTestNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                Else
                    genTestNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                End If
                If bRedondance = True Then 'Redondance si la var est du type E_BT1
                    bRedondance = False
                    iCounter = iCounter + 1
                    sVariable = "R2" & Mid(sVariable, 3, Len(sVariable) - 2)
                    If strPlage = "Booléen" Then
                        genTestNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, True, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    Else
                        genTestNNModel sequence, rg_EtapeDansSequence, sTempoVehicule, sInstance & sVariable, sValeur, False, CStr(iCounter) & strContexte, True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    End If
                End If
            End If
        End If
    iCounter = iCounter + 1
    rst_M_Verif.MoveNext
Wend
CompteVar_ES False, , iCounter - 1
'------------------------------------------------------
'Fonction activation surveillance HT
If Surveillance_HT Then
    bEnvir = True
    sInstance = GenCheminVariable(sTempoVehicule, "SURV_HT", True)
    genLabel sequence, rg_EtapeDansSequence, "-------------------", True, StepGroup_Main: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    genLectureModel sequence, rg_EtapeDansSequence, 1, sInstance, A1, "StationGlobals.ErrorHT", False, NumEtape & "_Surveillance HT", True, StepGroup_Main, False: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
    genLabel sequence, rg_EtapeDansSequence, "Surveillance HT - Code d'Erreur : ", True, StepGroup_Main, "StationGlobals.ErrorHT", True: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
End If
'------------------------------------------------------
'------------------------------------------------------
        
GenCodeEtape_ExitPopup:
    rst_F_Act.Close
    rst_F_Verif.Close
    rst_M_Act.Close
    rst_M_Verif.Close
    '------------------------------------------------
    GoTo GenCodeEtape_Exit

GenCodeEtape_Err:
    ' Affiche les informations sur l'erreur.
    MsgBox "Numéro d'erreur " & Err.Number & ": " & Err.Description
    ' Reprend l'exécution à l'instruction suivant la ligne où l'erreur s'est produite.
    Resume Next
GenCodeEtape_Exit:
End Sub


Public Sub GenCodeTest(ID_TEST As Integer, Tst_Code As String, Tst_Indice As String, ConfPR As String, ByVal ConfBanc As String)
 ' ConfPR   Configuration du programme de test
 ' ConfBanc configuration du banc
 
 '   On Error GoTo GenCodeTest_Err
    
Dim bds As Database, rst As Recordset
Dim bds2 As Database, rst2 As Recordset
Dim rst_conf As Recordset
Dim sPC_Ihm, sPC_Fip As String
Dim TestPrec, indicePrec As String  'Test et indice précédent
Dim iTempo As Single
Dim iTabVeh(6) As Integer
Dim Etape_Skip As Boolean
Dim Test_Skip As Boolean
Set bds = CurrentDb
Set bds2 = CurrentDb
Set db_ihm = CurrentDb

If Tst_Code = "" And Tst_Indice = "" Then
    MsgBox "Génération arretée, code de test inexistant !"
    iContinue = 4
    Exit Sub
End If


'------------------------
'' Set rst2 = bds2.OpenRecordset("SELECT * FROM VEHICULE WHERE ((VEHICULE.TST_CODE)= " & Chr(34) & Tst_Code & Chr(34) & " AND ((VEHICULE.TST_INDICE)= " & Chr(34) & Tst_Indice & Chr(34) & "))", dbOpenDynaset)
' --------------------------------
' Def_Section


sPC_Fip = ArrTable_Config_Rezo(1)    'IP ou nom réseau du Pc SCXI !!!! NE DOIT JAMAIS BOUGER, CAR 1 CIBLE !!!!!
sPC_Ihm = ArrTable_Config_Rezo(4)    'IP ou nom réseau du Pc IHM  !!!! NE DOIT JAMAIS BOUGER, CAR 1 CARTE D'ACQUISITION !!!!!
bRedondance = False

Gsql = ""
Gsql = "SELECT T_Test.ID_Test, T_Test.Test_Code, T_Test.Test_Indice,T_Test.Test_Skip, T_Etape.ID_Etape, T_Etape.Etape_Code, T_Etape.Etape_Indice, T_Test.Test_Designation, T_Test.Test_Commentaire, T_Etape.Etape_Tempo, T_Etape.Etape_Skip, T_Test_Etape.Ordre"
Gsql = Gsql & " FROM T_Test INNER JOIN (T_Etape INNER JOIN T_Test_Etape ON T_Etape.ID_Etape = T_Test_Etape.ID_Etape) ON T_Test.ID_Test = T_Test_Etape.ID_Test"
Gsql = Gsql & " GROUP BY T_Test.ID_Test, T_Test.Test_Code, T_Test.Test_Indice,T_Test.Test_Skip, T_Etape.ID_Etape, T_Etape.Etape_Code, T_Etape.Etape_Indice, T_Test.Test_Designation, T_Test.Test_Commentaire, T_Etape.Etape_Tempo, T_Etape.Etape_Skip, T_Test_Etape.Ordre"
Gsql = Gsql & " HAVING (((T_Test.Id_test) =" & ID_TEST & "))"
Gsql = Gsql & " ORDER BY T_Test_Etape.Ordre;"

Set rst = db_ihm.OpenRecordset(Gsql, dbOpenDynaset)

If ConfBanc = 2 Then '2 sections utilisées avec 1 env
    iTabVeh(1) = 1  ' ---------------- Section 1
    iTabVeh(2) = 1  ' ---------------- Section 2
    iTabVeh(3) = 0
    iTabVeh(4) = 0
    iTabVeh(5) = 0
    iTabVeh(6) = 1  'Env 1 utilisé
ElseIf ConfBanc = 3 Then '3 sections avec 2 env
    iTabVeh(1) = 1  ' ---------------- Section 1
    iTabVeh(2) = 1  ' ---------------- Section 2
    iTabVeh(3) = 0
    iTabVeh(4) = 1  ' ---------------- Section 4
    iTabVeh(5) = 1  'Env 2 utilisé
    iTabVeh(6) = 1  'Env 1 utilisé
ElseIf ConfBanc = 4 Then '3 sections avec 2 env
    iTabVeh(1) = 1 ' ---------------- Section 1
    iTabVeh(2) = 1 ' ---------------- Section 2
    iTabVeh(3) = 1  ' ---------------- Section 3
    iTabVeh(4) = 1  ' ---------------- Section 4
    iTabVeh(5) = 1  'Env 2 utilisé
    iTabVeh(6) = 1  'Env 1 utilisé
End If

NumEtapeDansTest = 1 'Numéro de l'étape dans le test
rg_EtapeDansSequence = 0    'Remise à 0 global du rang dans le Test
Test_Skip = False

While Not rst.EOF
    If (rst.Fields("Test_Skip").Value = True And Mode_NonReg = True) Then
        Test_Skip = True    'Skippe le pas si seulement pour la Non Reg
    End If
    
    If IsNull(rst.Fields("Etape_Tempo").Value) Then
        iTempo = 3000 ' Pas de valeur donc 3000 ms par défaut
    Else
        iTempo = rst.Fields("Etape_Tempo").Value  'La valeur présente est prioritaire
    End If
    
    Etape_Skip = rst.Fields("Etape_Skip").Value 'Skippe l'étape prioritaire
    
    rst.MoveNext
    If Not rst.EOF Then
        TestPrec = rst.Fields("Etape_Code")
        indicePrec = rst.Fields("Etape_Indice")
    End If
    rst.MovePrevious
        
    Call GenCodeEtape(bds, NumEtapeDansTest, rst.Fields("ID_Etape"), rst.Fields("Etape_Code").Value, rst.Fields("Etape_Indice").Value, iTempo, rst.Fields("Test_Designation"), Etape_Skip, Test_Skip, iTabVeh(), sPC_Fip, sPC_Ihm, ID_TEST, ConfPR, ConfBanc)
    
    rst.MoveNext
    NumEtapeDansTest = NumEtapeDansTest + 1
Wend

rst.Close
'rst2.Close
Set bds = Nothing
Set bds2 = Nothing
'------------------------------------------------
Exit Sub
GenCodeTest_Err:
' Affiche les informations sur l'erreur.
MsgBox "Numéro d'erreur " & Err.Number & ": " & Err.Description
' Reprend l'exécution à l'instruction suivant la ligne où l'erreur s'est produite.
Resume Next

End Sub

Public Sub GenCodeSeqDocument(blnGenDoc As Boolean)
    ' Si les paramètres ne sont pas vide c'est qu'on génère de façon sélective
    'On Error GoTo GenCodeSeqDocument_Err
    '------------------------------------------------
    Dim bds As Database, rst As Recordset, rstCount As Recordset, doccrt As String, i As Integer
    ' Retourne une référence à la base de données active.
    Set bds = CurrentDb
    '------------------------------------------------
    ' Ouvre l'objet Recordset des étapes par test
    If Not blnGenDoc Then
        Set rst = bds.OpenRecordset("SELECT DOCUMENT.DOC_CODE, DOCUMENT.DOC_INDICE, DOCUMENT.DOC_NUM, TEST.TST_CODE, TEST.TST_INDICE, TEST.LOC_CODE FROM (DOCUMENT INNER JOIN DOCUMENT_TEST ON (DOCUMENT.DOC_INDICE = DOCUMENT_TEST.DOC_INDICE) AND (DOCUMENT.DOC_CODE = DOCUMENT_TEST.DOC_CODE)) INNER JOIN TEST ON (DOCUMENT_TEST.TST_INDICE = TEST.TST_INDICE) AND (DOCUMENT_TEST.TST_CODE = TEST.TST_CODE) ORDER BY DOCUMENT.DOC_NUM, DOCUMENT_TEST.DOC_TST_ORDRE", dbOpenDynaset)
        Set rstCount = bds.OpenRecordset("SELECT DOCUMENT.DOC_NUM, Count(DOCUMENT_TEST.TST_CODE) AS NB_TST_CODE FROM (DOCUMENT INNER JOIN DOCUMENT_TEST ON (DOCUMENT.DOC_INDICE = DOCUMENT_TEST.DOC_INDICE) AND (DOCUMENT.DOC_CODE = DOCUMENT_TEST.DOC_CODE)) INNER JOIN TEST ON (DOCUMENT_TEST.TST_INDICE = TEST.TST_INDICE) AND (DOCUMENT_TEST.TST_CODE = TEST.TST_CODE) GROUP BY DOCUMENT.DOC_NUM ORDER BY DOCUMENT.DOC_NUM", dbOpenDynaset)
    Else
        'Set rst = bds.OpenRecordset("SELECT DOCUMENT.DOC_CODE, DOCUMENT.DOC_INDICE, DOCUMENT.DOC_NUM, TEST.TST_CODE, TEST.TST_INDICE, TEST.LOC_CODE FROM (DOCUMENT INNER JOIN DOCUMENT_TEST ON (DOCUMENT.DOC_CODE = DOCUMENT_TEST.DOC_CODE) AND (DOCUMENT.DOC_INDICE = DOCUMENT_TEST.DOC_INDICE)) INNER JOIN TEST ON (DOCUMENT_TEST.TST_CODE = TEST.TST_CODE) AND (DOCUMENT_TEST.TST_INDICE = TEST.TST_INDICE) WHERE (((DOCUMENT.DOC_CODE) = " & Chr(34) & strNomDoc & Chr(34) & ") And ((DOCUMENT.DOC_INDICE) = " & Chr(34) & strVersionDoc & Chr(34) & ")) ORDER BY DOCUMENT.DOC_NUM, DOCUMENT_TEST.DOC_TST_ORDRE", dbOpenDynaset)
        Set rst = bds.OpenRecordset("SELECT DOCUMENT.DOC_CODE, DOCUMENT.DOC_INDICE, DOCUMENT.DOC_NUM, TEST.TST_CODE, TEST.TST_INDICE, TEST.LOC_CODE FROM DOC_GEN_SEL INNER JOIN ((DOCUMENT INNER JOIN DOCUMENT_TEST ON (DOCUMENT.DOC_INDICE = DOCUMENT_TEST.DOC_INDICE) AND (DOCUMENT.DOC_CODE = DOCUMENT_TEST.DOC_CODE)) INNER JOIN TEST ON (DOCUMENT_TEST.TST_INDICE = TEST.TST_INDICE) AND (DOCUMENT_TEST.TST_CODE = TEST.TST_CODE)) ON (DOC_GEN_SEL.DOC_INDICE = DOCUMENT.DOC_INDICE) AND (DOC_GEN_SEL.DOC_CODE = DOCUMENT.DOC_CODE) ORDER BY DOCUMENT.DOC_NUM, DOCUMENT_TEST.DOC_TST_ORDRE", dbOpenDynaset)
        'Set rstCount = bds.OpenRecordset("SELECT DOCUMENT.DOC_NUM, Count(DOCUMENT_TEST.TST_CODE) AS NB_TST_CODE FROM (DOCUMENT INNER JOIN DOCUMENT_TEST ON (DOCUMENT.DOC_CODE = DOCUMENT_TEST.DOC_CODE) AND (DOCUMENT.DOC_INDICE = DOCUMENT_TEST.DOC_INDICE)) INNER JOIN TEST ON (DOCUMENT_TEST.TST_CODE = TEST.TST_CODE) AND (DOCUMENT_TEST.TST_INDICE = TEST.TST_INDICE) GROUP BY DOCUMENT.DOC_NUM, DOCUMENT.DOC_CODE, DOCUMENT.DOC_INDICE HAVING (((Document.DOC_CODE) = " & Chr(34) & strNomDoc & Chr(34) & ") And ((Document.DOC_INDICE) = " & Chr(34) & strVersionDoc & Chr(34) & ")) ORDER BY DOCUMENT.DOC_NUM", dbOpenDynaset)
        Set rstCount = bds.OpenRecordset("SELECT DOCUMENT.DOC_NUM, Count(DOCUMENT_TEST.TST_CODE) AS NB_TST_CODE FROM DOC_GEN_SEL INNER JOIN ((DOCUMENT INNER JOIN DOCUMENT_TEST ON (DOCUMENT.DOC_INDICE = DOCUMENT_TEST.DOC_INDICE) AND (DOCUMENT.DOC_CODE = DOCUMENT_TEST.DOC_CODE)) INNER JOIN TEST ON (DOCUMENT_TEST.TST_INDICE = TEST.TST_INDICE) AND (DOCUMENT_TEST.TST_CODE = TEST.TST_CODE)) ON (DOC_GEN_SEL.DOC_INDICE = DOCUMENT.DOC_INDICE) AND (DOC_GEN_SEL.DOC_CODE = DOCUMENT.DOC_CODE) GROUP BY DOCUMENT.DOC_NUM ORDER BY DOCUMENT.DOC_NUM", dbOpenDynaset)
    End If
    
    If rst.RecordCount > 0 Then
        doccrt = ""
    Else
        MsgBox "Il n'y a pas de séquenceurs de documents à générer !"
        GoTo GenCodeSeqDocument_Exit
    End If
    
    While Not rst.EOF
        If (rst.Fields("DOC_NUM").Value <> doccrt) Then
            If doccrt <> "" Then
                ' Génération de la fin de fichier
                rst.MovePrevious
'''''                Call GenCodeSeqDocumentFin(rst)
                rst.MoveNext
                rstCount.MoveNext
            End If
            
            'Création d'un nouveau fichier
            doccrt = rst.Fields("DOC_NUM").Value
            Close #1
            
            Select Case Len(CStr(doccrt))
            Case 1
                Open Left(CurrentDb.Name, Len(CurrentDb.Name) - Len("Valids.mdb")) & "ISaGRAF\SqDoc00" & doccrt & ".lsf" For Output As #1
            Case 2
                Open Left(CurrentDb.Name, Len(CurrentDb.Name) - Len("Valids.mdb")) & "ISaGRAF\SqDoc0" & doccrt & ".lsf" For Output As #1
            Case 3
                Open Left(CurrentDb.Name, Len(CurrentDb.Name) - Len("Valids.mdb")) & "ISaGRAF\SqDoc" & doccrt & ".lsf" For Output As #1
            End Select
                
            ' Génération du début de fichier
            Print #1, "#info=WSGE1EDT"
            Print #1, "@I:0,0=1;Séquencement Document " & rst.Fields("DOC_NUM").Value
            Print #1, "@O:1,0=0;"
            Print #1, "@d:1,1=0;"
            Print #1, "@T:2,0=1;"
            Print #1, "@T:2,1=4;"
            Print #1, "@S:3,0=2;"
            Print #1, "@S:3,1=4;"
            Print #1, "@T:4,0=2;"
            Print #1, "@T:4,1=5;"
            Print #1, "@S:5,0=3;Vide pour besoins de génération"
            Print #1, "@J:5,1=2;"
            Print #1, "@T:6,0=3;"
                cptLigne = 7
                cptColonne = 0
                cptEtape = 5
                cptTransition = 6
                For i = 1 To rstCount.Fields("NB_TST_CODE").Value Step 1
                    Print #1, "@S:" & cptLigne & "," & cptColonne & "=" & cptEtape & ";"
                    cptLigne = cptLigne + 1
                    cptEtape = cptEtape + 1
                    Print #1, "@T:" & cptLigne & "," & cptColonne & "=" & cptEtape & ";"
                    cptLigne = cptLigne + 1
                Next i
            Print #1, "@S:" & cptLigne & "," & cptColonne & "=" & cptEtape + 1 & ";"
            Print #1, "@T:" & cptLigne + 1 & "," & cptColonne & "=" & cptEtape + 2 & ";"
            Print #1, "@S:" & cptLigne + 2 & "," & cptColonne & "=" & cptEtape + 2 & ";"
            Print #1, "@T:" & cptLigne + 3 & "," & cptColonne & "=" & cptEtape + 3 & ";"
            Print #1, "@J:" & cptLigne + 4 & ",0=1;"
            Print #1, "#endinfo"
            Print #1, "_step _init GS1;"
            Print #1, "#sfc=GS1(0,0)"
            Print #1, "(* LE PREMIER GRAFCET DU MODE DOIT"
            Print #1, "ETRE UTILISE POUR LANCER LES GROUPES"
            Print #1, "DE TESTS, MEME SI IL N'Y EN A QU'UN SEUL *)"
            Print #1, ""
            Print #1, "(* Les groupes sont utilisés uniquement ici, ils servent"
            Print #1, "à gérer des ensembles de tests *)"
            Print #1, "#sfc=end"
            Print #1, "_next GT1, GT4;"
            Print #1, ""
            Print #1, "_trans GT1;"
            Print #1, "#sfc=GT1(0,2)"
                Select Case Len(rst.Fields("DOC_NUM").Value)
                Case 1
                    Print #1, "NOM_MODE = 'ENTREE_DC00" & rst.Fields("DOC_NUM").Value & "'"
                Case 2
                    Print #1, "NOM_MODE = 'ENTREE_DC0" & rst.Fields("DOC_NUM").Value & "'"
                Case 3
                    Print #1, "NOM_MODE = 'ENTREE_DC" & rst.Fields("DOC_NUM").Value & "'"
                End Select
            Print #1, "and BL_RECUP = FALSE;"
            Print #1, "#sfc=end"
            Print #1, "_next GS2;"
            Print #1, ""
            Print #1, "_trans GT4;"
            Print #1, "#sfc=GT4(1,2)"
                Select Case Len(rst.Fields("DOC_NUM").Value)
                Case 1
                    Print #1, "NOM_MODE = 'ENTREE_DC00" & rst.Fields("DOC_NUM").Value & "'"
                Case 2
                    Print #1, "NOM_MODE = 'ENTREE_DC0" & rst.Fields("DOC_NUM").Value & "'"
                Case 3
                    Print #1, "NOM_MODE = 'ENTREE_DC" & rst.Fields("DOC_NUM").Value & "'"
                End Select
            Print #1, "and BL_RECUP = TRUE;"
            Print #1, "#sfc=end"
            Print #1, "_next GS4;"
            Print #1, ""
            Print #1, "_step GS2;"
            Print #1, "#sfc=GS2(0,3)"
            Print #1, "Action (P):"
            Print #1, ""
                Select Case Len(rst.Fields("DOC_NUM").Value)
                Case 1
                    Print #1, "NOM_TEST_LOCAL := 'DC00" & rst.Fields("DOC_NUM").Value & "';"
                Case 2
                    Print #1, "NOM_TEST_LOCAL := 'DC0" & rst.Fields("DOC_NUM").Value & "';"
                Case 3
                    Print #1, "NOM_TEST_LOCAL := 'DC" & rst.Fields("DOC_NUM").Value & "';"
                End Select
            Print #1, "(* déclaration du nom du fichier de résultats du document " & rst.Fields("DOC_NUM").Value & "*)"
            Print #1, "If FICHIER_OUVERT = False Then"
            Print #1, "    NOM_FICH := NOM_REPERTOIRE + NOM_TEST_LOCAL + '_' + MSG(OTOMATIC_N) + '.TXT';"
            'Print #1, "    If RIGHT(RECUP_MODE, 6) = NOM_TEST_LOCAL THEN"
            Print #1, "    IF BL_RECUP THEN"
            Print #1, "        (* Ouverture du fichier en ajout *)"
            Print #1, "        FICHIER_OUVERT := LOG_ON(nom_fich, TRUE);"
            Print #1, "        RETOUR_FONCTION := LOG_MSG('Fichier ouvert en ajout, le fichier de recupération existe');"
            Print #1, "    Else"
            Print #1, "        (* Ouverture du fichier en écrasement *)"
            Print #1, "        FICHIER_OUVERT := LOG_ON(nom_fich, FALSE);"
            Print #1, "    END_IF;"
            Print #1, "    RETOUR_FONCTION := LOG_MSG('Fichier de compte rendu de tests du ' + NOM_TEST_LOCAL);"
            Print #1, "    RETOUR_FONCTION := LOG_DATE(TRUE);"
            Print #1, "END_IF;"
            Print #1, ""
            Print #1, "END_ACTION;"
            Print #1, "#sfc=end"
            Print #1, "_next GT2;"
            Print #1, ""
            Print #1, "_step GS4;"
            Print #1, "#sfc=GS4(1,3)"
            Print #1, "(* Mettre ici la gestion du groupe de mise en conditions initiales"
            Print #1, "Cela peut faire plusieurs étapes*)"
            Print #1, "#sfc=end"
            Print #1, "_next GT5;"
            Print #1, ""
            Print #1, "_trans GT2;"
            Print #1, "#sfc=GT2(0,4)"
            Print #1, "TRUE;"
            Print #1, "#sfc=end"
            Print #1, "_next GS3;"
            Print #1, ""
            Print #1, "_trans GT5;"
            Print #1, "#sfc=GT5(1,4)"
            Print #1, "TRUE;"
            Print #1, "#sfc=end"
            Print #1, "_next GS2;"
            Print #1, ""
            Print #1, "_step GS3;"
            Print #1, "_next GT3;"
            Print #1, ""
            Print #1, "_trans GT3;"
            Print #1, "#sfc=GT3(0,6)"
            Print #1, "TRUE;"
            Print #1, "#sfc=end"
            Print #1, "_next GS5;"
            Print #1, ""
                cptLigne = 7
                cptColonne = 0
                cptEtape = 5
                cptTransition = 6
        End If 'Fin nouveau document
        
        ' Génération de la liste des tests par document
        Print #1, "_step GS" & cptEtape & ";"
        Print #1, "#sfc=GS" & cptEtape & "(" & cptColonne & "," & cptLigne & ")"
        Print #1, "Action (P):"
        Print #1, "(* Mettre le nom du groupe uniquement ICI *)"
        Print #1, "    TMP_TEST := '" & rst.Fields("TST_CODE").Value & rst.Fields("TST_INDICE").Value & "';"
        Print #1, ""
        Print #1, "(* Cette fonction détermine si le groupe doit être exécuté ou non"
        Print #1, "Elle positionne FLAG_GROUPE, pas besoin d'utiliser le retour"
        Print #1, "de la fonction pour l'instant. *)"
        Print #1, "    RETOUR_FONCTION := RECH_TST(TMP_TEST);"
        Print #1, ""
        Print #1, "END_ACTION;"
        Print #1, "#sfc=end"
        Print #1, "_next GT" & cptTransition & ";"
        Print #1, ""
        Print #1, "_trans GT" & cptTransition & ";"
            cptLigne = cptLigne + 1
        Print #1, "#sfc=GT" & cptTransition & "(" & cptColonne & "," & cptLigne & ")"
        Print #1, "NOM_TEST = 'SORTIE_" & rst.Fields("TST_CODE").Value & rst.Fields("TST_INDICE").Value & "';"
        Print #1, "#sfc=end"
            cptEtape = cptEtape + 1
        Print #1, "_next GS" & cptEtape & ";"
        Print #1, ""
            cptLigne = cptLigne + 1
            cptTransition = cptTransition + 1
        
        rst.MoveNext
        If (rst.EOF) Then
            rst.MoveLast
''''''            Call GenCodeSeqDocumentFin(rst)
            rst.MoveNext
        End If
    Wend

    Close #1
    'MsgBox "Les fichiers de mode ont été générés dans " & Left(CurrentDb.Name, Len(CurrentDb.Name) - Len("Valids.mdb")) & "ISaGRAF"
    rst.Close
    Set bds = Nothing
    '------------------------------------------------
    GoTo GenCodeSeqDocument_Exit
GenCodeSeqDocument_Err:
    ' Affiche les informations sur l'erreur.
    MsgBox "Numéro d'erreur " & Err.Number & ": " & Err.Description
    ' Reprend l'exécution à l'instruction suivant la ligne où l'erreur s'est produite.
    Resume Next
GenCodeSeqDocument_Exit:
End Sub

Public Sub GenCodeSeqModeFin(rst As Recordset)

    Print #1, "_step GS" & cptEtape & ";"
    Print #1, "#sfc=GS" & cptEtape & "(" & cptColonne & "," & cptLigne & ")"
    Print #1, "Action (P):"
    Print #1, ""
    Print #1, "(*fermeture fichier resultats du mode en cours*)"
    Print #1, "IF (FICHIER_OUVERT = TRUE) THEN"
    Print #1, "    RETOUR_FONCTION := LOG_MSG('Fin de mode');"
    Print #1, "    RETOUR_FONCTION := LOG_DATE(TRUE);"
    Print #1, "    RETOUR_FONCTION := LOG_DATE(FALSE);"
    Print #1, "    RETOUR_FONCTION := LOG_OFF(FALSE);"
    Print #1, "    FICHIER_OUVERT := FALSE;"
    Print #1, "END_IF;"
    Print #1, ""
    Print #1, "END_ACTION;"
    Print #1, "#sfc=end"
    Print #1, "_next GT" & cptTransition & ";"
    Print #1, ""
    Print #1, "_trans GT" & cptTransition & ";"
    Print #1, "#sfc=GT" & cptTransition & "(" & cptColonne & "," & cptLigne + 1 & ")"
    Print #1, "TRUE;"
    Print #1, "#sfc=end"
    Print #1, "_next GS" & cptEtape + 1 & ";"
    Print #1, ""
    Print #1, "_step GS" & cptEtape + 1 & ";"
    Print #1, "#sfc=GS" & cptEtape + 1 & "(" & cptColonne & "," & cptLigne + 2 & ")"
    Print #1, "Action (P):"
    Print #1, "    NOM_MODE := 'SORTIE_MODE_" & rst.Fields("MIS_CODE").Value & "';"
    Print #1, "END_ACTION;"
    Print #1, "#sfc=end"
    Print #1, "_next GT" & cptTransition + 1 & ";"
    Print #1, ""
    Print #1, "_trans GT" & cptTransition + 1 & ";"
    Print #1, "#sfc=GT" & cptTransition + 1 & "(" & cptColonne & "," & cptLigne + 3 & ")"
    Print #1, "TRUE;"
    Print #1, "#sfc=end"
    Print #1, "_next GS1;"

End Sub

Public Sub GenCodeMode(blnGenDoc As Boolean)
    On Error GoTo GenCodeMode_Err
    
    Dim bds As Database, rst As Recordset, modeCrt As String
    Set bds = CurrentDb
    If Not blnGenDoc Then
        Set rst = bds.OpenRecordset("SELECT TEST.MIS_CODE, TEST.TST_CODE, TEST.TST_INDICE FROM TEST ORDER BY TEST.MIS_CODE, TEST.TST_ORDRE", dbOpenDynaset)
    Else
        Set rst = bds.OpenRecordset("SELECT TEST.MIS_CODE, TEST.TST_CODE, TEST.TST_INDICE, FROM TEST_GEN_SEL INNER JOIN TEST ON (TEST_GEN_SEL.TST_INDICE = TEST.TST_INDICE) AND (TEST_GEN_SEL.TST_CODE = TEST.TST_CODE) GROUP BY TEST.TST_ORDRE, TEST.MIS_CODE, TEST.TST_CODE, TEST.TST_INDICE ORDER BY TEST.MIS_CODE, TEST.TST_ORDRE", dbOpenDynaset)
    End If
    
    If rst.RecordCount > 0 Then
        modeCrt = ""
    Else
        MsgBox "Il n'y a pas de fichier de Mode à générer !"
        GoTo GenCodeMode_Exit
    End If
    

    'MsgBox "Les fichiers de mode ont été générés dans " & Left(CurrentDb.Name, Len(CurrentDb.Name) - Len("Valids.mdb")) & "ISaGRAF"
    rst.Close
    Set bds = Nothing
    '------------------------------------------------
    GoTo GenCodeMode_Exit
GenCodeMode_Err:
    ' Affiche les informations sur l'erreur.
    MsgBox "Numéro d'erreur " & Err.Number & ": " & Err.Description
    ' Reprend l'exécution à l'instruction suivant la ligne où l'erreur s'est produite.
    Resume Next
GenCodeMode_Exit:
End Sub

Public Sub InitGeneration() 'Cette fonction permet d'initialiser les valeurs avant la génération d'un PR ou d'un groupe de PR
Dim texte As String

If Conf_bancReel = 2 Then
    texte = "FileGlobals.NbEnv = 1, FileGlobals.NbEmb=2, FileGlobals.NbSection=2"
ElseIf Conf_bancReel = 3 Then
    texte = "FileGlobals.NbEnv = 2, FileGlobals.NbEmb=3, FileGlobals.NbSection=3"
ElseIf Conf_bancReel = 4 Then
    texte = "FileGlobals.NbEnv = 2, FileGlobals.NbEmb=4, FileGlobals.NbSection=4"
Else
    texte = "FileGlobals.NbEnv = 1, FileGlobals.NbEmb=2, FileGlobals.NbSection=2"
End If

genStatement "MainSequence", 0, "MainSequence", texte, "", StepGroup_Setup

TestStandIHMInited = False  'Initialise l'interface de l'IHM

End Sub

Public Sub GenCodeTout()
   
'''On Error GoTo GenCodeTout_Err
    
Dim rg_Test, rg_Etape As Integer  'Rang des lignes au sein des séquences (test et etape)
Dim TempsDepart, TempsArrive As Date
Dim sNomSeqSortie, sTexteSortie As String
Dim Nouveau_FichierTS As Boolean  'Nouveau fichier TS à générer
Dim ConfTest As String

iContinue = 0   'Pas d'erreur

Set db_ihm = CurrentDb
Set rs_test = db_ihm.OpenRecordset("SELECT * FROM Tempo_Test;", dbOpenDynaset)
Set rs_document = db_ihm.OpenRecordset("SELECT * FROM Tempo_Document;", dbOpenDynaset)

If rs_test.EOF = True Then
    iContinue = 2
    GoTo GenCodeTout_Err
End If

If rs_document.EOF = True Then
    iContinue = 3
    GoTo GenCodeTout_Err
End If

Form_Gene_en_cours.visible = True
DoEvents

Nouveau_FichierTS = True
TempsDepart = Time
Call Temporaire 'Remplissage de tableau pour économiser des temps requetes
PriseCabmission = False
'''**************************
''''on désactive le calcul de la présence de la variable dans la STR
Calcul_VarDansSTR = False
'''**************************
                                    
rs_document.MoveFirst
rs_test.MoveFirst
While Not rs_test.EOF 'Scrute la table Tempo_Test
    rg_Test = 0 'Variable permettant de savoir à quel niveau on ajoute le test dans la séquence
    rg_Etape = 0
    
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    While rs_test.Fields("Test_Code") <> "" And iContinue = 0 And rs_test.Fields("Test_Code") <> "-"
        If Nouveau_FichierTS = True Then
            'Création du nouveau PR de Test, Init des valeurs de base
            iContinue = 0
            CompteVar_ES True   'Init du compteur de variables E/S
            'Génération de l'init (Mode mission en entrée + Verif modeles en Auto)
            Call TestStandTCS.InitTestStand 'Init de TestStand
            InitGeneration      'Init générale avant de générer un PR, reprend le choix de la conf_BancReel
            
            sequence = "Sequence d'Entrée" 'la batterie est prise par défaut dans la section 1
            genInsertSequence rg_Test, sequence     'génération de la séquence
            genAppelSéquence "MainSequence", rg_Test, sequence, sequence, "", "", StepGroup_Main: rg_Test = rg_Test + 1  'ajout de la séquence au fichier
            rg_Etape = 0  'Remise à 0 du compteur de lignes
            '-------------------------------------------------
            genAppelSéquenceAutreFichier sequence, 0, "Séquence d'Entrée", "MainSequence", "C:\Teststand\Sequence_Entree.seq", False, False, ""
            '-------------------------------------------------
  '          Call GenCodeCheckModeleAuto(rg_Etape, "M", StepGroup_Setup, iNbVehicleMaxi)     ' Dans le Setup de l'init, pour l'instant correspond au mode mission M (en dur)
           ' Call GenMission(0, 2, StepGroup_Main, , "E")        ' Dans le MainSequence de l'init
            Nouveau_FichierTS = False
        End If
        '-------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------
        'Génération de la séquence (PR)
        If iContinue = 0 Then
            PriseCabmission = False
            
            ' **********************
            ' Récupération du type de conf PR-banc avant génération
            ConfTest = rs_test.Fields("Test_Mission")
            Select Case ConfTest
            Case 2
                sequence = CStr(rg_Test) & "_TEST : " & rs_test.Fields("Test_Code") & " - Conf = A1"
            Case 4
                sequence = CStr(rg_Test) & "_TEST : " & rs_test.Fields("Test_Code") & " - Conf = B"
            Case 5
                sequence = CStr(rg_Test) & "_TEST : " & rs_test.Fields("Test_Code") & " - Conf = C"
            Case 6
                sequence = CStr(rg_Test) & "_TEST : " & rs_test.Fields("Test_Code") & " - Conf = D"
            Case 7
                sequence = CStr(rg_Test) & "_TEST : " & rs_test.Fields("Test_Code") & " - Conf = A2"
            End Select
            
            sequence = CStr(rg_Test) & "_TEST : " & rs_test.Fields("Test_Code")
            genInsertSequence rg_Test, sequence     'génération de la séquence
            genAppelSéquence "MainSequence", rg_Test, sequence, sequence, "", "", StepGroup_Main: rg_Test = rg_Test + 1 'ajout de la séquence au fichier
            ' **********************
            'ConfTest       :conf du Test du PR
            'Conf_bancReel  :conf pour la génération
            Call GenCodeTest(rs_test.Fields("Test_Id"), rs_test.Fields("Test_Code"), rs_test.Fields("Test_Indice"), ConfTest, Conf_bancReel)   'Ajout des tests dans les séquences
        End If
        If iContinue > 0 Then
            GoTo GenCodeTout_Err
        End If
        '-------------------------------------------------------------------------------------
        rs_test.MoveNext
    Wend
    '-------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------
    If Nouveau_FichierTS = False And iContinue = 0 Then
    'Génération de la sequence de sortie
        PriseCabmission = False
        rs_test.MovePrevious
        sequence = "Sequence de Sortie"
        genInsertSequence rg_Test, sequence     'génération de la séquence
        genAppelSéquence "MainSequence", rg_Test, sequence, sequence, "", "", StepGroup_Main: rg_Test = rg_Test + 1  'ajout de la séquence au fichier
     '   Call GenMission(0, rs_test.Fields("Test_Mission"), StepGroup_Main, , "S")        ' Dans le MainSequence
        genAppelSéquenceAutreFichier sequence, 0, "Séquence de sortie", "MainSequence", "C:\Teststand\Sequence_Sortie.seq", False, False, ""
        'Affichage du nombre de variables d'E/S
''        MsgBox "Entrée = " & intNBVar_E & vbCr & "Sortie = " & intNBVar_S
 '''       genLabel sequence, rg_Test, "", True, StepGroup_Main
        
        TestStandEngine.ReloadGlobals
        sNomSeqSortie = rs_document.Fields("Document_Num") & " " & rs_document.Fields("Document_Indice")

        ''------------------------------------------------
        '' Gestion par gamme (EP20 par défaut)
        If PR_Destination = 1 Then      'Prima proto1
            sNomSeqSortie = "EP" & Mid(sNomSeqSortie, 3)
        End If
        
        ''------------------------------------------------
        '' Gestion mode UM
        If Memo_ModeMultiple Then
            sNomSeqSortie = sNomSeqSortie '& "_UM"
        End If
        ''------------------------------------------------
        
        Call TestStandTCS.CloseTestStand(sNomSeqSortie)        'Fermeture de TestStand et Sauvegarde de la séquence
        rs_document.MoveNext
        Nouveau_FichierTS = True
    End If
    rs_test.MoveNext
Wend
        
Form_Gene_en_cours.visible = False
TempsArrive = Time - TempsDepart
Call ReportError(ArrTable_ReportError)  'Création d'un fichier reportant toutes les erreurs de compilation

If iContinue = 0 And iCounterError = 0 Then sTexteSortie = "Génération effectuée dans " & strPathTestStandOutput & vbCr & "Temps : " & TempsArrive & " (h:m:s)"
If iContinue = 0 And iCounterError > 0 Then sTexteSortie = "Génération effectuée dans " & strPathTestStandOutput & vbCr & "Temps : " & TempsArrive & " (h:m:s)" & vbCr & "Il y a eu " & CStr(iCounterError) & " Erreurs" & vbCr & "Cf. D:\Data\Report_Err_Gen_TS.txt"
If iContinue > 0 Then sTexteSortie = "La génération a eu un problème, Veuillez vérifier vos paramètres d'init des PR." & vbCr & "Il y a eu " & CStr(iCounterError) & " Erreurs" & vbCr & "Cf. D:\Data\Report_Err_Gen_TS.txt"
If Mode_NonReg = True Then sTexteSortie = sTexteSortie & vbCr & "Mode NON-REG activé !!!"
MsgBox sTexteSortie, vbInformation, "Fin de la génération"

rs_test.Close
rs_document.Close
'db_ihm.Close
'------------------------------------------------
Exit Sub
GenCodeTout_Err:
    Select Case iContinue
    Case 1
        MsgBox "Le Mode Mission est introuvable !", vbCritical, "Erreur Génération !"
    Case 2
        MsgBox "Il y a aucun test à générer !", vbCritical, "Erreur Génération !"
    Case 3
        MsgBox "Il y a aucun PR à générer !", vbCritical, "Erreur Génération !"
    Case 4
        MsgBox "Check des Modèles en Auto !", vbCritical, "Erreur Génération !"
    Case Else
        MsgBox "Numéro d'erreur " & Err.Number & ": " & Err.Description, vbCritical, "Erreur Génération !"
    End Select
    Form_Gene_en_cours.visible = False
    
    
    saveTestStand ""
End Sub


Public Sub Temporaire() 'Fonction exécutée qu'une seule fois au tout début du lancement d'une génération

Dim iCounterY As Long
Dim bds As Database
Dim rst As Recordset
 
Set bds = CurrentDb
If iContinue = 0 Then
    'Tableau temporaire de la liste et de l'adressage des equipements du Model
    Set rst = bds.OpenRecordset("SELECT * FROM Tref_EquipementCB;")
    iCounterY = 0
    While (rst.EOF = False)
        iCounterY = iCounterY + 1
        ArrTable_Eqt_Model(1, iCounterY) = rst.Fields("EquipementCB_Code").Value
        If IsNull(rst.Fields("EquipementCB_Prefixe").Value) Then
            ArrTable_Eqt_Model(2, iCounterY) = ""
        Else
            ArrTable_Eqt_Model(2, iCounterY) = rst.Fields("EquipementCB_Prefixe").Value
        End If
        If IsNull(rst.Fields("EquipementCB_AdresseNommage").Value) Then
            ArrTable_Eqt_Model(3, iCounterY) = ""
        Else
            ArrTable_Eqt_Model(3, iCounterY) = rst.Fields("EquipementCB_AdresseNommage").Value
        End If
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    'Tableau de la config réseau pour la loco 1
    Set rst = bds.OpenRecordset("SELECT * FROM Tref_ConfigGen")
    ArrTable_Config_Rezo(0) = rst.Fields("ENV_PC5").Value    'IP Env 1
    ArrTable_Config_Rezo(1) = rst.Fields("ENV_PC6").Value    'Env 2
    ArrTable_Config_Rezo(2) = rst.Fields("EMB_PC1").Value    'Emb 1
    ArrTable_Config_Rezo(3) = rst.Fields("EMB_PC2").Value    'Emb 2
    ArrTable_Config_Rezo(4) = rst.Fields("EMB_PC3").Value    'Emb 3
    ArrTable_Config_Rezo(5) = rst.Fields("EMB_PC4").Value    'Emb 4
    sVariable_Version_CB = VersionCB                            'Variable permettant de connaitre la version de l'appli CB
''    sChemin_SL2 = rst.Fields("DIRECTOR_SL2").Value               'Répertoire d'accès au fichier ihm
    A1 = rst.Fields("N_Surveillance").Value
    a2 = rst.Fields("ZP_Surveillance").Value
    
    Poste1 = rst.Fields("NET_VEH1").Value
    Poste2 = rst.Fields("NET_VEH2").Value
    
    rst.Close
    Set rst = Nothing
    
    'Sauve toutes les plages des variables, valeur max, min et tolérance des variables
    iCounterY = 0
    Set rst = CurrentDb.OpenRecordset("SELECT T_Variable.Variable_Code, Tref_PlageVariable.PlageVariable_Type, Tref_PlageVariable.PlageVariable_Precision FROM T_Variable INNER JOIN Tref_PlageVariable ON T_Variable.Variable_Plage = Tref_PlageVariable.ID_PlageVariable;")
    While Not rst.EOF
        ArrTable_VarPlage(0, iCounterY) = rst.Fields("Variable_Code")    'Variable
        ArrTable_VarPlage(1, iCounterY) = rst.Fields("PlageVariable_Type")    'Pression, bool, manip, speed, tempo
 '       ArrTable_VarPlage(2, iCounterY) = rst.Fields("PlageVariable_Unite")   'Unité
 '       ArrTable_VarPlage(3, iCounterY) = rst.Fields("PlageVariable_Min")    'Valeur min
 '       ArrTable_VarPlage(4, iCounterY) = rst.Fields("PlageVariable_Max")     'Valeur max
        ArrTable_VarPlage(5, iCounterY) = rst.Fields("PlageVariable_Precision")     'Précision en %
        iCounterY = iCounterY + 1
        rst.MoveNext
    Wend
    rst.Close
    Set rst = Nothing
    'Remise à zéro de la table des erreurs
    iCounterError = 0
    For iCounterY = 0 To 1500
        ArrTable_ReportError(iCounterY) = ""
    Next
    'Remise à zéro du pointeur du listing des variables Forcées/déforcées
    intCounter_ListeVarForce = 0
End If
bds.Close

End Sub

Sub generationSelective(ID_Doc As Integer, ByVal ConfAGenerer As Integer) 'ajouter les tests du PR dans Tempo_Test
'Seul les tests pouvant etre déroulés sur le banc seront générés
Dim bds As Database
Dim rst, rstTEST_GEN_SEL As Recordset
Dim ConfPR As Integer
Dim b_genere As Boolean

Set bds = CurrentDb
Gsql = ""
Gsql = "SELECT T_PR_Test.Ordre, T_Test.Test_Code, T_Test.Test_Indice,T_Test.ID_Test, T_Document.ID_Document, T_Document.Document_Titre, T_Test.Test_ConfigTrain"
Gsql = Gsql & " FROM T_Test INNER JOIN (T_Document INNER JOIN T_PR_Test ON T_Document.ID_Document = T_PR_Test.ID_PR) ON T_Test.ID_Test = T_PR_Test.ID_Test"
Gsql = Gsql & " WHERE (((T_Document.ID_Document)=" & ID_Doc & "))"
Gsql = Gsql & " ORDER BY T_PR_Test.Ordre;"

Set rst = bds.OpenRecordset(Gsql, dbOpenDynaset)
Set rstTEST_GEN_SEL = bds.OpenRecordset("Tempo_Test")

With rst
    While Not .EOF
        b_genere = CompatibilitePR_Test(!Test_ConfigTrain, ConfAGenerer)
        If b_genere Then
            rstTEST_GEN_SEL.AddNew
            rstTEST_GEN_SEL.Fields("Test_Id") = !ID_TEST
            rstTEST_GEN_SEL.Fields("Test_Code") = !Test_Code
            rstTEST_GEN_SEL.Fields("Test_Indice") = !Test_Indice
            rstTEST_GEN_SEL.Fields("Test_Mission") = !Test_ConfigTrain
            rstTEST_GEN_SEL.Fields("Test_Titre") = !Document_Titre
            rstTEST_GEN_SEL.Update
        End If
        .MoveNext
    Wend
    .Close
End With

rstTEST_GEN_SEL.AddNew
rstTEST_GEN_SEL.Fields("Test_Code") = "-"
rstTEST_GEN_SEL.Fields("Test_Indice") = "-"
rstTEST_GEN_SEL.Update
rstTEST_GEN_SEL.Close
bds.Close

End Sub


Function Trouve_CheminFBS(strVariableFip) As String
Dim i As Integer
Dim iMax As Integer
Dim sChemin As String

''A voir s'il faut augmenter cette valeur !!
    iMax = 9500
    For i = 0 To iMax
        If ArrTable_Var_CheminMPU(0, i) = strVariableFip Then
            sChemin = ArrTable_Var_CheminMPU(1, i)
            Exit For
        End If
    Next
        
Trouve_CheminFBS = sChemin
End Function

Function Trouve_Tolerance(strVariableFip) As Long
Dim i As Long
Dim iMax As Long
Dim sModule_MMAP As String

    ''A revoir car peut prendre du temps !!! exécuté en local (PC rapide)
    iMax = 300000
    For i = 0 To iMax
        If ArrTable_VarPlage(0, i) = strVariableFip Then
            Exit For
        End If
    Next
        
Trouve_Tolerance = i
End Function

Function RemplacePointVirgule(ByVal sValue As String) As String
Dim sTemp As String
Dim i As Integer
    
    sTemp = sValue
    
    If InStr(1, sTemp, ",") > 1 Then
        i = InStr(1, sTemp, ",")
        sTemp = Mid(sTemp, 1, i - 1) & "." & Right(sTemp, Len(sTemp) - i)
    End If

    RemplacePointVirgule = sTemp
End Function

Sub Config_Mpu(Tst_Code As String, Tst_Indice As String)

Dim iCount As Integer
Dim rstMpu As Recordset 'Requete sur la base véhicule
Dim rstconf As Recordset 'Requete pour trouver la config correcte
Dim sTexte As String

For iCount = 1 To 6
    arConfig_Mpu(iCount) = 0  'Remise à zéro de la config Mpu
Next

Set rstMpu = CurrentDb.OpenRecordset("SELECT CONF_CODE FROM VEHICULE WHERE TST_CODE=" & Chr(34) & Tst_Code & Chr(34) & " AND TST_INDICE=" & Chr(34) & Tst_Indice & Chr(34))
If IsNull(rstMpu.Fields("CONF_CODE")) Then GoTo Fin
sTexte = rstMpu.Fields("CONF_CODE")
Set rstconf = CurrentDb.OpenRecordset("SELECT * FROM CONFIG_MPU WHERE CONF_CODE=" & sTexte)
sTexte = rstconf.Fields("CONF_EQT")
For iCount = 1 To 6
    If Mid(sTexte, iCount, 1) <> 0 Then
        arConfig_Mpu(iCount) = Mid(sTexte, iCount, 1)
    Else
        Exit For
    End If
Next
Set rstMpu = Nothing
Set rstconf = Nothing

Exit Sub
Fin:
iContinue = 10 'Erreur de conf MPU

End Sub


Sub GenNonReg()
'Génération d'une séquence spécifique pour appeler chaque séquence pour la NON REG
Dim rst As Recordset
Dim sNomSequuence, sDoc_Code, sDoc_Indice, sequence, sTitre As String
Dim iCounter As Integer
Dim o3, o4 As Object
Dim a, b
    
    If NumeroPREnregistre = 0 Then
    sNomSequuence = Left(Gchemin, Len(Gchemin) - 4) & "Models\Modele_Non_Reg.seq"   'Ne pas modifier ce chemin !!!
    Else
    sNomSequuence = Left(Gchemin, Len(Gchemin) - 4) & "TestStand\Sequence_NonReg_" & CStr(Date$) & "_NRG.seq"
    End If
    Set TestStandEngine = CreateObject("TestStand.Engine")
    Set GenereTCS = TestStandEngine.GetSequenceFile(sNomSequuence)
    
    a = TestStandEngine.VersionString
    b = CDec(Left(a, 1))
    If b >= 3 Then
        GEN_TS = 3
        LoadTypePaletteVer3     'En fonction de la version de TS, appelle la fonction prévue
    Else
        GEN_TS = 2
        LoadTypePaletteVer2     'Pointe sur une fonction obsolète pour les versions 3 et +
    End If
    
 '   iCounter = 1
    If NumeroPREnregistre = 0 Then
        NumeroPREnregistre = 1
    Else
        NumeroPREnregistre = NumeroPREnregistre + 1
    End If
    sNomFichier_NonReg = ""
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM Tempo_Document")
    While (rst.EOF = False)
        sDoc_Code = rst.Fields("Document_Num")
        sDoc_Indice = rst.Fields("Document_Indice")
        If PR_Origine <> PR_Destination Then
            If PR_Origine = 1 Then      'Prima proto1 vers ...
                If PR_Destination = 2 Then sDoc_Code = "MAEL" & Mid(sDoc_Code, 5)
                If PR_Destination = 3 Then sDoc_Code = "AXEL" & Mid(sDoc_Code, 5)
'                If PR_Destination = 4 Then sDoc_Code = "???" & Mid(sDoc_Code, 5)
            ElseIf PR_Origine = 2 Then  'Maroc vers ...
                If PR_Destination = 1 Then sDoc_Code = "ATEL" & Mid(sDoc_Code, 5)
                If PR_Destination = 3 Then sDoc_Code = "AXEL" & Mid(sDoc_Code, 5)
'                If PR_Destination = 4 Then sDoc_Code = "???" & Mid(sDoc_Code, 5)
            ElseIf PR_Origine = 3 Then 'prima proto2 vers ...
                If PR_Destination = 1 Then sDoc_Code = "ATEL" & Mid(sDoc_Code, 5)
                If PR_Destination = 2 Then sDoc_Code = "MAEL" & Mid(sDoc_Code, 5)
'                If PR_Destination = 4 Then sDoc_Code = "???" & Mid(sDoc_Code, 5)
            End If
        End If
        sTitre = sDoc_Code & " " & sDoc_Indice & "_NRG" & ".seq"
        sNomFichier_NonReg = sNomFichier_NonReg & vbCr & sTitre
        sequence = "Main" & CStr(NumeroPREnregistre)
        genInsertSequence NumeroPREnregistre, sequence
        genAppelSéquenceNONREG sequence, 0, sequence, "MainSequence", sTitre, StepGroup_Main
        NumeroPREnregistre = NumeroPREnregistre + 1
        iCounter = iCounter + 1
        rst.MoveNext
    Wend
'    'Ajout du dernier pas
'    sequence = "Main" & CStr(iCounter)
'    genInsertSequence rang, sequence
'    genStatement sequence, 0, sequence, "StationGlobals.bSequenceMultiple=False", "", StepGroup_Main

    'Sauvegarde et fermeture TestStand
 '   NumeroPREnregistre = rang - 1
    NumeroPREnregistre = NumeroPREnregistre - 1
    sNomSequuence = "Sequence_NonReg_" & CStr(Date$)
    CloseTestStand sNomSequuence
End Sub

Public Function Recherche_Precond(ETP_CODE As String, ETP_INDICE As String, bVar_Reseau As Boolean, bEtape As Boolean) As Integer

Dim iRes As Integer
Dim rst As Recordset

    iRes = 0   'Pas de précondition, si > 0 Alors renvoi l'ID
    If bEtape = True Then
    'Mode de génération dans l'étape
        If bVar_Reseau = True Then 'Variable FIP
            Set rst = CurrentDb.OpenRecordset("SELECT ID FROM VARACTION WHERE ETP_CODE=" & Chr(34) & ETP_CODE & Chr(34) & " AND ETP_INDICE=" & Chr(34) & ETP_INDICE & Chr(34))
            If Not (rst.EOF) Then
                If IsNull(rst.Fields("ID")) = False Then iRes = rst.Fields("ID")
            End If
        Else       'Variable modèle
            Set rst = CurrentDb.OpenRecordset("SELECT ID FROM VAR_ACT_MODEL WHERE ETP_CODE=" & Chr(34) & ETP_CODE & Chr(34) & " AND ETP_INDICE=" & Chr(34) & ETP_INDICE & Chr(34))
            If Not (rst.EOF) Then
                If IsNull(rst.Fields("ID")) = False Then iRes = rst.Fields("ID")
            End If
        End If
    Else
    'Mode de génération dans le mode Mission
        If bVar_Reseau = True Then 'Variable FIP
            Set rst = CurrentDb.OpenRecordset("SELECT ID FROM VARACTION_M WHERE ETP_CODE=" & Chr(34) & ETP_CODE & Chr(34) & " AND ETP_INDICE=" & Chr(34) & ETP_INDICE & Chr(34))
            If Not (rst.EOF) Then
                If IsNull(rst.Fields("ID")) = False Then iRes = rst.Fields("ID")
            End If
        Else       'Variable modèle
            Set rst = CurrentDb.OpenRecordset("SELECT ID FROM VAR_ACT_MODEL_M WHERE ETP_CODE=" & Chr(34) & ETP_CODE & Chr(34) & " AND ETP_INDICE=" & Chr(34) & ETP_INDICE & Chr(34))
            If Not (rst.EOF) Then
                If IsNull(rst.Fields("ID")) = False Then iRes = rst.Fields("ID")
            End If
        End If
    End If
Recherche_Precond = iRes
End Function

Public Function Ajout_Precond(ID As Integer, bVar_Reseau As Boolean, bEtape As Boolean, ByVal iDest As Integer, Optional sAdress_FIP As String) As Integer

' ETP_CODE      -> Code de l'étape
' ETP_INDICE    -> Indice de l'étape
' bVar_Reseau   -> Booléen de reconnaissance du type de variable (Réseau ou modèle)
' bEtape        -> Booléen de présence de la variable (Etape ou mode Mission)

Dim rst, rst_Plage As Recordset
Dim iError, iNumConsist As Integer
Dim strVariable_FIP, strVariable_Model As String
Dim strEquipment_FIP, strEquipment_Model As String
Dim bTypeVariable As Boolean  'Si vrai alors la variable est du type FIP
Dim strTolerance, sVehicule, sValeur, str_EQT As String
Dim bMPU1, bMPU2, boolPresentRP, boolAccesNickname As Boolean
Dim strContexte, strPlage, sInstance, sPrecond As String

    iError = 0 'Pas d'erreur dans la fonction
    boolPresentRP = False
    
    If bEtape = True Then
    'Mode de génération dans l'étape
        If bVar_Reseau = True Then 'Precond à intégrer avant l'action FIP
            Set rst = CurrentDb.OpenRecordset("SELECT * FROM PRECOND_NORMAL WHERE ID=" & ID)
            If Not (rst.EOF) Then
                If IsNull(rst.Fields("VAR_CODE")) = False Then strVariable_FIP = rst.Fields("VAR_CODE")
                If IsNull(rst.Fields("EQT_CODE")) = False Then strEquipment_FIP = rst.Fields("EQT_CODE")
                If IsNull(rst.Fields("VAR_CODE2")) = False Then strVariable_Model = rst.Fields("VAR_CODE2")
                If IsNull(rst.Fields("EQT_CODE2")) = False Then strEquipment_Model = rst.Fields("EQT_CODE2")
                sVehicule = rst.Fields("VEHICULE")
                If IsNull(rst.Fields("VALEUR")) = False Then
                    sValeur = rst.Fields("VALEUR")
                Else
                    sValeur = "0"
                End If
                str_PreCond = "FileGlobals.intSauve==" & sValeur
                bMPU1 = rst.Fields("MMAP_MPU1")
                bMPU2 = rst.Fields("MMAP_MPU2")
                If Len(strVariable_FIP) > 1 And Len(strEquipment_FIP) > 1 Then
                    bTypeVariable = True
                    strTolerance = ArrTable_VarPlage(5, Trouve_Tolerance(strVariable_FIP))
                ElseIf Len(strVariable_Model) > 1 And Len(strEquipment_Model) > 1 Then
                    bTypeVariable = False
                Else  'Erreur
                    iError = -1 'Pas de variable présent dans les préconditions (Etape - FIP)
                    GoTo Gestion_Sortie
                End If
            End If
        Else       'Precond à intégrer avant l'action Model
            Set rst = CurrentDb.OpenRecordset("SELECT * FROM PRECOND_MODEL WHERE ID=" & ID)
            If Not (rst.EOF) Then
                If IsNull(rst.Fields("VAR_CODE")) = False Then strVariable_FIP = rst.Fields("VAR_CODE")
                If IsNull(rst.Fields("EQT_CODE")) = False Then strEquipment_FIP = rst.Fields("EQT_CODE")
                If IsNull(rst.Fields("VAR_CODE2")) = False Then strVariable_Model = rst.Fields("VAR_CODE2")
                If IsNull(rst.Fields("EQT_CODE2")) = False Then strEquipment_Model = rst.Fields("EQT_CODE2")
                sVehicule = rst.Fields("VEHICULE")
                If IsNull(rst.Fields("VALEUR")) = False Then
                    sValeur = rst.Fields("VALEUR")
                Else
                    sValeur = "0"
                End If
                str_PreCond = "FileGlobals.intSauve==" & sValeur
                bMPU1 = rst.Fields("MMAP_MPU1")
                bMPU2 = rst.Fields("MMAP_MPU2")
                If Len(strVariable_FIP) > 1 And Len(strEquipment_FIP) > 1 Then
                    bTypeVariable = True
                    strTolerance = ArrTable_VarPlage(5, Trouve_Tolerance(strVariable_FIP))
                ElseIf Len(strVariable_Model) > 1 And Len(strEquipment_Model) > 1 Then
                    bTypeVariable = False
                Else  'Erreur
                    iError = -2 'Pas de variable présent dans les préconditions (Etape - Model)
                    GoTo Gestion_Sortie
                End If
            End If
        End If
    ' Génération générale de la ligne de l'étape supplémentaire
        If bTypeVariable = True Then 'Variable du type FIP
            Set rst_Plage = CurrentDb.OpenRecordset("SELECT * FROM VARIABLE WHERE ((VAR_CODE= " & Chr(34) & strVariable_FIP & Chr(34) & "))", dbOpenDynaset)
            strPlage = rst_Plage.Fields("PLG_CODE").Value
            Set rst_Plage = Nothing
            If strEquipment_FIP = "SERVEUR_RP" Then boolPresentRP = True
  ''          strVariable_FIP = GenVariableFip(strVariable_FIP, sVehicule, boolPresentRP, strEquipment_FIP, , , sValeur, True, , ID & " (PRECOND_NORMAL)")
            str_EQT = strNomEqtFip
            strContexte = " Vehicle=" & sVehicule & " ; Equipment=" & strEquipment_FIP & " ; "
            If Left(str_EQT, 1) = "@" Then
                If bMPU1 = True Or bMPU2 = True Then
                    If bMPU1 = True Then str_EQT = "@mpu1" & Mid(str_EQT, 2, Len(str_EQT) - 1) & strVariable_FIP 'Seul le MPU1 est utilisé
                    If bMPU2 = True Then str_EQT = "@mpu2" & Mid(str_EQT, 2, Len(str_EQT) - 1) & strVariable_FIP 'Seul le MPU2 est utilisé
                    intSauve = CInt(sValeur)
                Else
                    iError = -10 'Aucun MPU de renseigner pour la variable FIP
                    GoTo Gestion_Sortie
                End If
            Else
                If strPlage = "Booléen" Then
                    genLectureFIP sequence, rg_EtapeDansSequence, sAdress_FIP, strVariable_FIP, sValeur, sVehicule, str_EQT, True, strContexte, iDest, "", True: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    str_PreCond = "FileGlobals.boolSauve==" & sValeur    'Si la variable attendue est un booléen
                Else
                    genLectureFIP sequence, rg_EtapeDansSequence, sAdress_FIP, strVariable_FIP, sValeur, sVehicule, str_EQT, False, strContexte, iDest, "", True: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                End If
            End If
        Else  'Variable du type Modèle
            If CInt(sVehicule) <= 6 Then iNumConsist = 3
            If CInt(sVehicule) <= 4 Then iNumConsist = 2
            If CInt(sVehicule) <= 2 Then iNumConsist = 1
            Set rst_Plage = CurrentDb.OpenRecordset("SELECT * FROM VARIABLE_MODEL WHERE VAR_CODE= " & Chr(34) & strVariable_Model & Chr(34))
            strPlage = rst_Plage.Fields("PLG_CODE").Value
            boolAccesNickname = rst_Plage.Fields("VAR_DOUBLON").Value
            Set rst_Plage = Nothing
            sValeur = RemplacePointVirgule(sValeur)
            strContexte = " Vehicle=" & sVehicule & " ; Equipment=" & strEquipment_Model & " ; "
            If (boolAccesNickname = False) Then 'Traitement cas Mnémonique du LV_VEHICLE
                sInstance = GenCheminVariable(sVehicule, strEquipment_Model, False)
                If strPlage = "Booléen" Then
                    genTestModel sequence, rg_EtapeDansSequence, sInstance, sVehicule, strVariable_Model, sValeur, True, strContexte, True, iDest: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    str_PreCond = "FileGlobals.boolSauve==" & sValeur    'Si la variable attendue est un booléen
                Else
                    genTestModel sequence, rg_EtapeDansSequence, sInstance, sVehicule, strVariable_Model, sValeur, False, strContexte, True, iDest: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                End If
            Else
                strVariable_Model = GenVariableModel(strVariable_Model, False)
                sInstance = "L" & sVehicule & "_"
                If strPlage = "Booléen" Then
                    genLectureNNModel sequence, rg_EtapeDansSequence, sVehicule, sInstance & strVariable_Model, sValeur, True, strContexte, iDest, "", True: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                    str_PreCond = "FileGlobals.boolSauve==" & sValeur    'Si la variable attendue est un booléen
                Else
                    genLectureNNModel sequence, rg_EtapeDansSequence, sVehicule, sInstance & strVariable_Model, sValeur, False, strContexte, iDest, "", True: rg_EtapeDansSequence = rg_EtapeDansSequence + 1
                End If
            End If
        End If
    Else
' Les tables pour les modes mission n'existent pas (évolution futur)
'    'Mode de génération dans le mode Mission
'        If bVar_Reseau = True Then 'Precond à intégrer avant l'action FIP
'            Set rst = CurrentDb.OpenRecordset("SELECT * FROM PRECOND_NORMAL_M WHERE ID=" & ID)
'            If Not (rst.EOF) Then
'                strVariable_FIP = rst.Fields("VAR_CODE")
'                strEquipment_FIP = rst.Fields("EQT_CODE")
'                strVariable_Model = rst.Fields("VAR_CODE2")
'                strEquipment_Model = rst.Fields("EQT_CODE2")
'
'                If Len(strVariable_FIP) > 1 And Len(strEquipment_FIP) > 1 Then
'                    bTypeVariable = True
'                ElseIf Len(strVariable_Model) > 1 And Len(strEquipment_Model) > 1 Then
'                    bTypeVariable = False
'                Else  'Erreur
'                    iError = -3 'Pas de variable présent dans les préconditions (Mission - FIP)
'                    GoTo Gestion_Sortie
'                End If
'
'
'            End If
'        Else       'Precond à intégrer avant l'action Model
'            Set rst = CurrentDb.OpenRecordset("SELECT * FROM PRECOND_MODEL_M WHERE ID="  & ID)
'            If Not (rst.EOF) Then
'                strVariable_FIP = rst.Fields("VAR_CODE")
'                strEquipment_FIP = rst.Fields("EQT_CODE")
'                strVariable_Model = rst.Fields("VAR_CODE2")
'                strEquipment_Model = rst.Fields("EQT_CODE2")
'
'                If Len(strVariable_FIP) > 1 And Len(strEquipment_FIP) > 1 Then
'                    bTypeVariable = True
'                ElseIf Len(strVariable_Model) > 1 And Len(strEquipment_Model) > 1 Then
'                    bTypeVariable = False
'                Else  'Erreur
'                    iError = -4 'Pas de variable présent dans les préconditions (Mission - Model)
'                    goto Gestion_Sortie
'                End If
'
'
'            End If
'        End If
    End If
Gestion_Sortie:
Ajout_Precond = iError
End Function

Public Function ParametersSequence(ByVal TypeMarcheP As String)
'A partir du type de Marché, la variable local au PC est mis à jour
Dim seq As String
Dim N As Integer
Dim texte As String

If (TypeMarcheP = "EP") Then
    N = 1    'EP20
    'bTempoPrima = True
End If
If (TypeMarcheP = "MA") Then
    N = 2    'Maroc
   ' bTempoPrima = False
End If

If ModeMultiple = True Or Memo_ModeMultiple = True Then
    texte = "True"
Else
    texte = "False"
End If

seq = "MainSequence"
genStatement seq, 0, "Definition du type de Marché (Gamme)", "StationGlobals.Prima = " & N, "", StepGroup_Setup, True
genStatement seq, 1, "Definition de l'UM", "FileGlobals.ModeMultiple = " & texte, "", StepGroup_Setup, True
End Function

Public Function InsertLabel(NomSequence As String, ByVal Groupe As Integer, ByVal Ligne As Integer, Champ As String) As Integer
genLabel NomSequence, Ligne, Champ, True, Groupe
InsertLabel = Ligne + 1
End Function



Public Function CompteVar_ES(RemiseAZero As Boolean, Optional NombreVar_Entree As Long = 0, Optional NombreVar_Sortie As Long = 0) As Boolean
'Fonction qui comptabilise le nombre de variables en Entrées/Sorties
Dim Resultat As Boolean

Resultat = False
If RemiseAZero Then
    intNBVar_E = 0
    intNBVar_S = 0
    Resultat = True
Else
    If NombreVar_Entree >= 0 And NombreVar_Sortie >= 0 Then
        intNBVar_E = intNBVar_E + NombreVar_Entree
        intNBVar_S = intNBVar_S + NombreVar_Sortie
        Resultat = True
    End If
End If

CompteVar_ES = Resultat
End Function

Private Function getValTolérance(ByVal strVariable As String) As Double
Dim rstVar As Recordset

'Set rstVar = CurrentDb.OpenRecordset("SELECT PLAGE.* FROM VARIABLE INNER JOIN PLAGE ON VARIABLE.PLG_CODE = PLAGE.PLG_CODE WHERE VARIABLE.VAR_CODE = " & Chr(34) & strVariable & Chr(34))
Set rstVar = Nothing
If Not rstVar.EOF Then
    If Not IsNull(rstVar!PLG_PRECISION) Then
        getValTolérance = CDec(rstVar!PLG_PRECISION)
    Else
       getValTolérance = -1
    End If
Else
    getValTolérance = -1
End If

rstVar.Close
Set rstVar = Nothing

End Function



