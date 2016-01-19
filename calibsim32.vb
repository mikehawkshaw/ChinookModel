Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module basSimulation
	'--------------------------------------------------------------
	'  PACIFIC SALMON COMMISSION CHINOOK MODEL
    '  VB.net version 1.0
    '  December 5, 2003
	'  FILE:     CALIBSIM.BAS        = computation module
	'
	'  COMPANION FILES:
	'  ----------------
	'     CTCinput.BAS        = read CTC input files
	'     CALIBOUT.BAS        = OUTPUT GENERATION MODULE
	'     MAIN.BAS            = define global variables, control call to modules
	'     FRM.MAIN            = prompts for files names and some constants
	'     CoastModelInput.BAS = auxilliary data input module (read Coast Model input files)
	'
	'  this version updates scrCohrt after each time step
	'  do NOT define a time step in timestep.txt without a fishery
	'  because of instantaneous maturity rates are being used
	'  and to prevent excapements from being calculated when there are no fisheries
	'--------------------------------------------------------------
	
	
	Dim AEQCohort As Double 'shared between calcTerminalRunAndEscapement and calcTotalFishingMortalities
	Dim age1Fish As Double 'shared between CalcAge1Production and GSHcheck
    Dim baseER(,,) As Double
    Dim biggestProb, BigProb As Double
    Dim CCCFileID As Integer
    Dim CCCfish_ISBMFileID As Integer
    Dim CCCfish_AABMFileID As Integer
    Dim CCCstkFileID As Integer
    Dim ceil(,) As Double
    Dim CLB_HISfileID As Integer
    Dim CLB_DATfileID As Integer
    Dim isCNR As Boolean
    Dim CNRLegal(,,,) As Double
    Dim CNRShakCat(,,,) As Single
    Dim ConvFlag As Integer
    Dim ConvFlagMR As Integer
    Dim encounterRatio(,) As Double
    Dim enhSpawn As Double
    Dim EscapeT(,), Escape(,) As Double
    Dim escapement As Double
    Dim EVS_RECfileID As Integer
    Dim fracNV(,,) As Single
    Dim GSH_Cat_EV(,) As Double
    Dim harvestRate(,,) As Double
    Dim is_Canadian() As Boolean
    Dim is_ISBM_fishery() As Boolean, is_ISBM_stock() As Boolean
    Dim lastRepeat As Integer
    Dim MATFRL_DATFileID As Integer
    Dim legalCatch(,,,) As Double
    Dim legalDropOffs(,,,) As Double
    Dim maturityRateApplied(,) As Integer
    Dim numLoops As Integer
    Dim OcnAECat() As Double
    Dim oceanCohort(,) As Double 'cohort/survival rate
    Dim OPT As Double 'shared between CalcAge1Production and GSHcheck
    Dim otherDat(,) As Double 'array shared between checkEV, calcEV, and calcMaturityRateAdjustments
    Dim probBY As Double
    Dim ProbStock As String
    Dim RT_scalar() As Single
    Dim RT_temp(,) As Double
    Dim scrCohrt(,) As Double
    Dim Spring() As Double
    Dim stkShakCat(,,,) As Double
    'Dim stockMonTest As Integer 'shared between calcEscapement, calcTerminalRunAndEscapement, and calcTotalFishingMortalities
    Dim region As Integer 'preterm / term
    Dim tempCat(,) As Double
    Dim TempStkCat(,,) As Double
    Dim OceanEscape_EV(,,) As Double
    Dim TermRun(,) As Double
    Dim timeStep As Integer
    Dim timeStep1Cohort(,) As Double
    Dim timeStep1TermRun(,) As Double
    Dim totAECat() As Double
    Dim totBasePeriodCatch(,) As Double
    Dim totEsc_allAges() As Double 'shared between calcTerminalRunAndEscapement and calcTotalFishingMortalities
    Dim totalEscape_EV(,,) As Double 'for EV calculations
    Dim yr As Integer

    Dim AEQfileID As Integer
    Dim catchFileID As Integer
    Dim CNRfileID As Integer
    Dim cohortFileID As Integer
    Dim exploitationRateFileID As Integer
    Dim ISBMcohortFileID As Integer, ISBMfileID As Integer
    Dim spawnEscFileID As Integer
    Dim shakerFileID As Integer
    Dim termRunFileID As Integer
    Dim termRun_OcnNetFileID As Integer
    Dim totalMortFileID As Integer


    '--------------------------------------------------------------
    Sub CalibSim()
        '--------------------------------------------------------------
        'purpose:  this is the entry point for the simulation module
        '          call annualLoop until ConvFlag% = 1 Or Iter = NumIter
        '          at the end of each annualLoop, call checkEV
        '          At the end of this subroutine, control goes back to main.bas
        '
        'Called By: main.bas
        '--------------------------------------------------------------


        Dim endTime, startTime As Double

        ReDim baseER(numFisheries, numStocks, NumAges)
        ReDim Escape(NumAges, numStocks), EscapeT(NumAges, numStocks)
        ReDim harvestRate(numFisheries, NumAges, numStocks)
        ReDim oceanCohort(numStocks, NumAges)
        ReDim ocnAEHR(NumYears, numStocks)
        ReDim Spring(numStocks)

        Iter = 0
        lastRepeat = 0

        Select Case isCalibration
            Case True 'stage 1 or 2 calibration
                Call Calibrate(1) 'housekeeping
                Call Calibrate(2) 'more housekeeping
            Case False 'projection run
                '..... Check if abundance index data are requested
                If SaveFile(7) = 0 Then
                    '..... Erase unnecessary array
                    Erase AI_fisheryIndex
                    Erase abundCohrt
                    Erase relAbund
                    Erase relAbundNonVuln
                Else
                    ReDim abundCohrt(NumYears, numStocks, NumAbundanceIndex)
                    ReDim relAbund(NumYears, NumAbundanceIndex)
                    ReDim relAbundNonVuln(NumYears, NumAbundanceIndex)
                End If
                '..... Erase unnecessary arrays
                Erase totalEscape_EV
                Erase GSH_Cat_EV
                Erase OceanEscape_EV
                Erase EV_lastYr
                Erase NumEVBroodYr
                Erase EVageFlag
                Erase PointBack
                Erase numStocksUsingSameEV
                Erase StartAge
        End Select

        '..... Determine start time
        startTime = VB.Timer()
        ConvFlagMR = -1

        Select Case isCalibration
            Case True
                'if calibration, then call annualLoop and checkEV until ConvFlag% = 1 Or Iter = NumIter
                Do
                    If ConvFlag = 1 Or Iter = NumIter Then
                        lastRepeat = 1
                        '..... Open Age-specific Calibration data file
                        'calibCheck (evaluate calibrations) is NOT the same as chkClb (compares projection runs)
                        Call CalibCheck(0, 1)
                    End If
                    Iter = Iter + 1
                    'Print #97, "start annual loop from calibsim, iteration"; Iter
                    Call AnnualLoop()

                    If lastRepeat = 1 Then
                        '..... Close age-specific calibration file
                        'calibCheck (evaluate calibrations) is NOT the same as chkClb (compares projection runs)
                        Call CalibCheck(0, 2)
                        FileClose(EVS_RECfileID)
                    Else
                        '..... Set convergence flag to 1 then check EV Convergence
                        ConvFlag = 1
                        Call CheckEV()
                    End If 'lastRepeat = 1

                Loop Until lastRepeat = 1
            Case False
                'if projection run, then call annualLoop only once
                lastRepeat = 1
                Call AnnualLoop()
        End Select

        '..... Done, still more housekeeping.
        If isCalibration = True Then Call Calibrate(3) 'stage 1 or 2 calibration

        '..... Determine end time for model run
        endTime = VB.Timer()
        '..... Compute elapsed time for model/calibration run
        runTimer = endTime - startTime
        If runTimer < 0 Then runTimer = runTimer + 86400

        If Stm = 4 Then
            FileClose(CCCfish_ISBMFileID)
            FileClose(CCCfish_AABMFileID)
            FileClose(CCCstkFileID)
        End If

        FileClose(EVS_RECfileID)

        If Outputs = 1 Then
            '..... Reset Calibration method for output generation for stage 1 or 2 calibration
            If isCalibration = True And CalibCycle Then CalibMethod = CalibSwitch \ CalibCycle
            '..... Call output module
            Call basOutput.CalibOut()
        End If

    End Sub

    '--------------------------------------------------------------
    Sub CalcRelativeAbundance()
        '--------------------------------------------------------------
        '  Purpose:  Calculates abundance index:
        '            Cohort x Survival x Harvest Rate x Prop Vuln
        '
        '  Arguments: Yr% = Year of simulation run.
        '
        '  Inputs:    AI_fisheryIndex()
        '             AbundanceIndex_PNV()  ***** NEW ARRAY 10/29/93
        '             Cohort_()
        '             harvestRate()
        '             NumAbundanceIndex
        '             NumAges
        '             numStocks
        '             PNV()
        '             PreTerm%
        '             survivalRate
        '             TermRun()
        '
        '  Called By: calcEscapement
        '
        '  Output:    AbundCohohrt()
        '             RelAbund()
        '
        '  Externals: FTerminal
        '
        '--------------------------------------------------------------

        Dim AABMfishery, age, fish, stk As Integer
        Dim tmp1, tmp2 As Double
        Dim tmp3, tmp4 As Double



        For AABMfishery = 1 To NumAbundanceIndex
            'RelAbund(yr, AABMfishery) = 0
            If timeStep = 1 Then
                relAbund(yr, AABMfishery) = 0
                relAbundNonVuln(yr, AABMfishery) = 0
            End If
            '.....  Determine index for fishery to compute abundance index
            fish = AI_fisheryIndex(AABMfishery)
            For stk = 1 To numStocks
                For age = 2 To NumAges
                    '****** CHANGE 10/29/93 TO estimate Abundance WRT PNVs in effect in a specified year
                    If AbundanceIndexBaseYr(AABMfishery) >= 0 Then
                        '..... Use specified PNV base year for abundance calculations
                        'eq 271 and eq 272 part 1
                        tmp2 = harvestRate(fish, age, stk) * (1 - AbundanceIndex_PNV(AABMfishery, age))
                        tmp4 = harvestRate(fish, age, stk) * AbundanceIndex_PNV(AABMfishery, age)
                    Else
                        '**********OLD CODE BELOW*********************************
                        '..... Use PNV's in effect each year for abundance calculations
                        'eq 271 and eq 272 part 1
                        tmp2 = harvestRate(fish, age, stk) * (1 - PNV(fish, age))
                        tmp4 = harvestRate(fish, age, stk) * PNV(fish, age)
                    End If
                    '*********END CHANGE**************************************
                    If FTerminal(stk, fish, age) = PreTerm Then
                        'eq 271 part 2
                        tmp1 = cohort_(age, stk) * tmp2
                        tmp3 = cohort_(age, stk) * tmp4
                    Else
                        'eq 272 part 2
                        tmp1 = TermRun(age, stk) * tmp2
                        'tmp = TotalTermRun(age%, stk%) * tmp2
                        tmp3 = TermRun(age, stk) * tmp4
                    End If
                    'eq 271 and eq 272 part 3
                    relAbund(yr, AABMfishery) = relAbund(yr, AABMfishery) + tmp1
                    relAbundNonVuln(yr, AABMfishery) = relAbundNonVuln(yr, AABMfishery) + tmp3
                    abundCohrt(yr, stk, AABMfishery) = abundCohrt(yr, stk, AABMfishery) + tmp1
                Next age
            Next stk
        Next AABMfishery

    End Sub

    '--------------------------------------------------------------
    Sub CalcEncounterRatio(ByRef location As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Calculate the shaker encounter rate
        '            (Nonvulnerable/Vulnerable population)
        '
        '  Arguments: location = Terminal or PreTerminal Fishery
        '             Yr%  = Simulation year
        '
        '  Inputs:    numFisheries
        '             oceanNetFlag%()
        '             PNV()
        '             PreTerm%
        '             ScrCohrt()
        '             TempCat()
        '             TempStkCat()
        '             TermRun()
        '
        '  Called By: calcPreterminalHarvest
        '             calcEscapement
        '
        '  Output:    FracNV!()
        '             StkWgt()
        '             TotPNV
        '             TotPV
        '
        '  Externals: CalcShakersAndCNR
        '             FTerminal
        '
        '--------------------------------------------------------------

        Dim totPNV, accum, totPV As Double
        Dim fish, age, stk As Integer
        Dim cohrt(,) As Double
        ReDim fracNV(numFisheries, numStocks, NumAges)

        For fish = 1 To numFisheries
            accum = 0
            If timeStepFlag(fish, timeStep) > 0 Then
                '..... Reset Cohrt()= 0, Cohrt is not indexed for fishery
                ReDim cohrt(NumAges, numStocks)
                totPNV = 0
                totPV = 0
                For stk = 1 To numStocks
                    If TempStkCat(location, fish, stk) > 0 Then
                        For age = 2 To NumAges
                            If FTerminal(stk, fish, age) = location Then
                                Select Case location
                                    Case PreTerm
                                        cohrt(age, stk) = scrCohrt(age, stk)
                                    Case Term
                                        cohrt(age, stk) = TermRun(age, stk)
                                End Select
                                'eq 41
                                totPNV = totPNV + cohrt(age, stk) * PNV(fish, age)
                                'eq 42
                                totPV = totPV + cohrt(age, stk) * (1 - PNV(fish, age))
                            Else
                                cohrt(age, stk) = 0
                            End If
                        Next age
                    End If
                Next stk

                If totPV > 0 Then
                    '..... Compute encounter ratio
                    'eq 43
                    encounterRatio(location, fish) = totPNV / totPV
                End If
                If totPNV > 0 Then
                    For stk = 1 To numStocks
                        'restore this line for original version of FracNV (i.e. not age specific FracNV)
                        If TempStkCat(location, fish, stk) > 0 Then 'use this to replicate QB with one step
                            For age = 2 To NumAges
                                'use this line for age specific FracNV
                                'If legalCatch(location, fish%, stk%, age%) > 0 Then 'use this for multiple time steps with age specific FracNV
                                'eq 48
                                fracNV(fish, stk, age) = (cohrt(age, stk) * PNV(fish, age)) / totPNV
                                accum = accum + fracNV(fish, stk, age)
                                'use this line for age specific FracNV
                                'End If 'legalCatch(location, fish%, stk%, age%) > 0 for multiple time steps with age specific FracNV
                            Next age
                            'restore this line for original version of FracNV
                        End If 'TempStkCat(location, fish%, stk%) > 0 to replicate QB with one step
                    Next stk

                    'standardize FracNV so the sum will add up to 100%
                    '            For stk% = 1 To NumStocks
                    '                For age% = 2 To NumAges
                    '                    If FracNV!(fish%, stk%, age%) > 0 Then FracNV!(fish%, stk%, age%) = FracNV!(fish%, stk%, age%) / accum
                    '                Next age%
                    '            Next stk%

                End If
            End If 'timeStepFlag(fish%, timeStep%) > 0

        Next fish

    End Sub

    '--------------------------------------------------------------
    Sub CalcEV(ByRef StN As Integer, ByRef stk As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Computes survival scalars for calibration run.
        '
        '  Arguments: StN%         = Starting age for calibration
        '             Stk%         = Stock index number
        '
        '  Inputs:    EV_firstYr%
        '             NumYears
        '             EVageFlag%()
        '             EV()
        '             otherDat() = Temporary data for model
        '                            convergence computations
        '                          shared with calcEV and calcMaturityRateAdjustments
        '
        '  Called By: CheckEV.BAS
        '
        '  Output:    BrdYrFlg%()
        '             ConTolerance
        '             EV()
        '
        '  Externals: ABS
        '
        '--------------------------------------------------------------

        Dim AdjControl, a, age As Integer
        Dim AvgStkSclr As Double
        'Dim BrdYrEV 
        Dim BrdYr, BYAdjust As Integer
        Dim ModelData As Double
        Dim Num As Double
        Dim obsEVdata() As Double
        'Dim PseudoEV, PseudoEVDat As Integer
        Dim ratio_s() As Double
        Dim StartEVAveYr As Integer
        Dim StkSclrMax As Double
        Dim str8, str11 As String
        Dim SumModel, SumEV, SumStk As Double
        Dim totalModDat, TotalModData, TotModData As Double
        Dim year_ As Integer

        ReDim obsEVdata(NumAges)

        str8 = New String("       ", 8)
        str11 = New String("           ", 11)

        '..... Reset arrays for calibration
        '2/24/96 This is a REPLACEMENT SUBROUTINE to allow EVs for broods prior to first model year ****
        ReDim ratio_s(NumYears + numAges1)

        'eq 243
        StkSclrMax = 1 / StkSclrMin

        '..... Compute adjustment ratios between calibration data and model estimates
        For BrdYr = -NumAges To NumEVBroodYr(stk)
            'Call GetEV(BrdYr, stk, BrdYrEV, PseudoEVDat)
            For age = 2 To NumAges
                obsEVdata(age) = BYEVdat(BrdYr + NumAges, stk, age) 'working copy 
            Next age
            If isEVforThisBroodYear(BrdYr + NumAges, stk) = True Then
                '..... EV is to be estimated for this Brood Year
                SumEV = 0
                SumModel = 0
                For age = StN To NumAges
                    year_ = BrdYr + age
                    '..... Check for last year of data
                    If year_ > EV_lastYr(stk) Then Exit For
                    If year_ >= 0 Then
                        If obsEVdata(age) > 0 Then 'observed data from .fcs file
                            ModelData = otherDat(year_, age)
                            SumModel = SumModel + ModelData
                            If EVageFlag(year_, stk) = 0 Then
                                If isPseudoAge(BrdYr + NumAges, stk) = False Then
                                    '.... No pseudo age data generated, treat as before
                                    TotalModData = 0
                                    For a = StN To NumAges
                                        TotalModData = TotalModData + otherDat(year_, a)
                                    Next a
                                    If TotModData > 0 Then
                                        '..... Prorate calibration data according to model-generated estimates
                                        'eq 237
                                        'note there is a problem here:  totalModDat is not define
                                        'however we don't get a divide by zero error message because
                                        'the program never goes here because EVageFlag%(year_%, stk%) = 1
                                        obsEVdata(age) = obsEVdata(age) * ModelData / totalModDat
                                    End If
                                End If
                            End If
                            SumEV = SumEV + obsEVdata(age)
                        End If

                    End If
                Next age

                If SumModel > 0 Then
                    'eq 235
                    ratio_s(BrdYr + NumAges) = SumEV / SumModel
                Else
                    ratio_s(BrdYr + NumAges) = StkSclrMax
                End If

            End If
        Next BrdYr

        '..... Adjust EV Scalars
        '   Set up EV calculation cycle for EV estimation, based on iteration
        '
        '   0 = every year
        '   1 = odd-even
        '   2 = Iter MOD 3; BrdYr MOD 2
        '   3 = Iter MOD 3; BrdYr MOD 3
        '   4 = Iter MOD 4; BrdYr MOD 3
        '****** NEW PATTERNS 2/23/96 *******
        '   5 = Iter MOD 4; All-Odd-All-Even
        '   6 = Adjust all on 5-year cycles
        ' AdjControl% = control patterns depending on simulation year

        Select Case CalibMethod
            Case 0
            Case 1
                AdjControl = Iter Mod 2
            Case 2, 3
                AdjControl = Iter Mod 3
            Case 4, 5
                AdjControl = Iter Mod 4
            Case 6
                AdjControl = Iter Mod 5
            Case Else
        End Select

        For BrdYr = -NumAges To NumEVBroodYr(stk)
            If Iter > CalibSwitch Then
                Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
            Else
                Select Case CalibMethod

                    Case 0
                        Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                    Case 1 '..... Odd-Even
                        BYAdjust = System.Math.Abs(BrdYr Mod 2)
                        If AdjControl = BYAdjust Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                    Case 2 '..... Iter MOD 3; BrdYr MOD 2
                        BYAdjust = System.Math.Abs(BrdYr Mod 3)
                        Select Case AdjControl
                            Case 1
                                If BYAdjust = 1 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 2
                                If BYAdjust = 0 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 0
                                Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                        End Select
                    Case 3 '..... Iter MOD 3; BrdYr MOD 3
                        BYAdjust = System.Math.Abs(BrdYr Mod 3)
                        If AdjControl = 0 Then
                            Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                        Else
                            If AdjControl = BYAdjust Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                        End If
                    Case 4 '..... Iter MOD 4; BrdYr MOD 3
                        BYAdjust = System.Math.Abs(BrdYr Mod 3)
                        Select Case AdjControl
                            Case 1
                                If BYAdjust = 1 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 2
                                If BYAdjust = 2 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 3
                                If BYAdjust = 0 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 0
                                Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                        End Select
                    Case 5 '..... 2/23/96 New Pattern All-Odd-All-Even
                        BYAdjust = System.Math.Abs(BrdYr Mod 2)
                        Select Case AdjControl
                            Case 1, 3
                                Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 2
                                If BYAdjust = 1 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 0
                                If BYAdjust = 0 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                        End Select
                    Case 6 '..... 2/23/96 New Pattern ALL-1-2-3-4-ALL
                        BYAdjust = System.Math.Abs(BrdYr Mod 4)
                        Select Case AdjControl
                            Case 2 '..... Adjust EVs for -NumAges to 0
                                If BYAdjust = 1 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 3
                                If BYAdjust = 2 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 4 '..... Adjust EVs for increments of 5
                                If BYAdjust = 3 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 0
                                If BYAdjust = 0 Then Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                            Case 1 '..... Adjust all EVs
                                Call AdjustEV(BrdYr, stk, ratio_s, StkSclrMax)
                        End Select

                End Select
            End If
        Next BrdYr
        '..... Compute average EV
        '...... 3/00 CDS  number of recent years to use in the average
        If NumEV2Ave > 0 Then
            StartEVAveYr = NumEVBroodYr(stk) - NumEV2Ave + 1
        Else
            StartEVAveYr = EV_firstYr
        End If

        SumStk = 0
        Num = 0
        For BrdYr = StartEVAveYr To NumEVBroodYr(stk)
            'Call GetEV(BrdYr, stk, BrdYrEV, PseudoEV)
            If isEVforThisBroodYear(BrdYr + NumAges, stk) = True Then
                'eq 250
                SumStk = SumStk + EV(BrdYr + NumAges, stk)
                Num = Num + 1
            End If
        Next BrdYr
        If Num > 0 Then
            'eq 251
            AvgStkSclr = SumStk / Num
        Else
            AvgStkSclr = 1
        End If

        '..... Set EVs to Average for indicated brood years
        For BrdYr = EV_firstYr To NumEVBroodYr(stk)
            'Call GetEV(BrdYr, stk, BrdYrEV, PseudoEV)
            'eq 252
            If isEVforThisBroodYear(BrdYr + NumAges, stk) = False Then EV(BrdYr + NumAges, stk) = AvgStkSclr
        Next BrdYr

        '..... Set EVs to average for all years after the last one to be estimated
        If NumEVBroodYr(stk) < MaxYrs Then
            For BrdYr = NumEVBroodYr(stk) + 1 To MaxYrs
                'eq 253
                EV(BrdYr + NumAges, stk) = AvgStkSclr
            Next BrdYr
        End If
    End Sub


    Sub AdjustEV(ByRef BrdYr As Integer, ByRef stk As Integer, ByRef ratio_s() As Double, ByRef StkSclrMax As Double)
        'this is called only by Sub CalcEV
        Dim newEV As Double
        Dim previousEV As Double
        Dim Z As Double


        'eq 244
        previousEV = EV(BrdYr + NumAges, stk)
        'eq 241
        '..... Compute ratio_S between calibration data and model estimate
        newEV = previousEV * ratio_s(BrdYr + NumAges)
        '&&&&& 6/99 GM Change to Select Case
        'eq 242
        Select Case newEV
            Case Is < StkSclrMin
                newEV = StkSclrMin
            Case Is > StkSclrMax
                newEV = StkSclrMax
            Case Else
                'no change
        End Select
        'save the average of the new and previous EV
        'eq 247
        EV(BrdYr + NumAges, stk) = (previousEV + newEV) / 2
        '..... Convergence Test - compute difference between old and new EVs
        If ConvMethod Then
            'eq 245
            Z = previousEV - newEV '..... Absolute difference
        Else
            Z = (previousEV - newEV) / previousEV '..... Relative difference
        End If

        'eq 246
        If System.Math.Abs(Z) > Convtolerance Then
            ConvFlag = 0
            If System.Math.Abs(Z) > biggestProb Then
                BigProb = EV(BrdYr + NumAges, stk)
                biggestProb = System.Math.Abs(Z)
                ProbStock = stock_shortName(stk)
                probBY = BrdYr + 79
            End If
        End If
    End Sub

    '--------------------------------------------------------------
    Sub CalcMaturityRateAdjustments(ByRef StN As Integer, ByRef stk As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Computes maturation rate adjustments for
        '            individual stocks as specified.
        '
        '  Arguments: StN%         = Starting age for calibration
        '             Stk%         = Stock index number
        '             otherDat() = Temporary data for model
        '                            convergence computations
        '                          and shared with checkEV and calcEV
        '  Inputs:    EV_firstYr%
        '             NumYears
        '             EVageFlag%()
        '             EV()
        '
        '  Called By: CheckEV
        '
        '  Output:    ConvFlagMR%
        '             maturationRate()
        '             AdltEqv()
        '             Age1toAdultSurvivalRate()
        '
        '  Externals: ABS
        '             FRecruitRate
        '
        '--------------------------------------------------------------

        'this subroutine was not called by stage 1, stage 2 or projection run
        'because AdjMatRate% = 0 in all of the .op7 files.

        Dim BrdYr As Integer, age As Integer
        'dim BrdYrEV As Integer
        Dim ageRatio(,) As Double
        Dim MRTest() As Integer
        Dim obsEVdata() As Double
        'Dim PseudoEV As Integer
        Dim TempMod As Double
        Dim TotModelDat() As Double
        Dim year_ As Integer

        '..... Reset arrays for calibration
        '***** 2/26/96 eliminated multiple comparisons and reduced arrays accordingly
        'ReDim ageRatio(3 To 4, NumYears)
        ReDim ageRatio(4, NumYears)
        ReDim obsEVdata(NumAges)
        ReDim TotModelDat(NumYears)

        '..... Compute total model data estimates by brood year
        For year_ = EV_firstYr To NumYears
            '..... Exit if EV factor is not to be estimated for this Stk% and Year
            If year_ > EV_lastYr(stk) Then Exit For
            For age = StN To NumAges
                '..... Compute total value of convergence variable
                '..... when age-specific data are not used for fitting
                TotModelDat(year_) = TotModelDat(year_) + otherDat(year_, age)
            Next age
        Next year_

        'ReDim MRTest(3 To 4)
        ReDim MRTest(4)
        '..... Compute ratios between calibration data and model estimates
        For BrdYr = EV_firstYr To NumEVBroodYr(stk)
            'Call GetEV(BrdYr, stk, BrdYrEV, PseudoEV)
            For age = 2 To NumAges
                obsEVdata(age) = BYEVdat(BrdYr + NumAges, stk, age) 'working copy
            Next age
            If isEVforThisBroodYear(BrdYr + NumAges, stk) = True Then
                For age = StN To NumAges
                    year_ = BrdYr + age

                    If year_ > EV_lastYr(stk) Then Exit For
                    If obsEVdata(age) > 0 Then
                        '..... Data specified for fitting
                        '..... store corresponding model estimate in temporary variable
                        TempMod = otherDat(year_, age)
                        '..... check if age-specific data are specified
                        If EVageFlag(year_, stk) = 0 Then
                            '..... age-specific data are NOT specified
                            '..... prorate by age using model-generated estimate
                            obsEVdata(age) = obsEVdata(age) * TempMod / TotModelDat(year_)
                        End If
                        If TempMod > 0 Then
                            '..... calculate ratios for maturation rate adjustments
                            Select Case age
                                Case 3
                                    ageRatio(3, BrdYr) = obsEVdata(age) / TempMod
                                Case 4
                                    ageRatio(4, BrdYr) = obsEVdata(age) / TempMod
                                Case Else
                            End Select
                        End If
                    End If
                Next age
            End If
        Next BrdYr

        '..... Test for convergence
        Select Case StkMat(stk)
            Case Is = 1 '..... Age 3 & 4 maturation rate adjustments
                age = 3
                Call MatRateCheck(ageRatio, MRTest)
                age = 4
                Call MatRateCheck(ageRatio, MRTest)
                If MRTest(3) = 1 And MRTest(4) = 1 Then
                    Exit Sub
                Else
                    ConvFlagMR = 0
                End If
            Case 3 '..... Age 3 maturation rate adjustments
                age = 3
                Call MatRateCheck(ageRatio, MRTest)
                If MRTest(3) = 1 Then
                    Exit Sub
                Else
                    ConvFlagMR = 0
                End If
            Case 4 '..... Age 4 maturation rate adjustments
                age = 4
                Call MatRateCheck(ageRatio, MRTest)
                If MRTest(4) = 1 Then
                    Exit Sub
                Else
                    ConvFlagMR = 0
                End If
        End Select

    End Sub

    Sub MatRateCheck(ByRef ageRatio(,) As Double, ByRef MRTest() As Integer)
        Dim stk As Integer
        Dim BrdYr As Integer
        Dim age As Integer
        'this is called only by Sub CalcMaturityRateAdjustments
        Dim AvgRatio, SumRatio As Double
        Dim Num As Integer


        MRTest(age) = 0
        '..... Compute average ratios for calibration data:model estimates
        SumRatio = 0
        Num = 0
        For BrdYr = EV_firstYr To NumEVBroodYr(stk)
            SumRatio = SumRatio + ageRatio(age, BrdYr)
            Num = Num + 1
        Next BrdYr
        If Num Then
            AvgRatio = SumRatio / Num
        Else
            AvgRatio = 1
        End If
        'eq 243
        If System.Math.Abs(1 - AvgRatio) <= Convtolerance Then
            MRTest(age) = 1
        Else
            '..... Adjust maturation rates
            Print(CLB_HISfileID, "MR Adjust Stk " & stk & " Age " & age & " Iter " & Iter & " Old: " & VB6.Format(maturationRate(age, stk), "0.####0 "))
            ConvFlag = 0
            maturationRate(age, stk) = FMaxOne(maturationRate(age, stk) * AvgRatio)
            PrintLine(CLB_HISfileID, "New " & VB6.Format(maturationRate(age, stk), "0.####0"))
            Select Case StkMat(stk)
                Case 1
                    If age = 4 Then Call CalcAEQ()
                Case 3
                    If age = 3 Then Call CalcAEQ()
                Case 4
                    If age = 4 Then Call CalcAEQ()
            End Select
        End If
    End Sub

    Sub CalcAEQ()
        Dim stk As Integer
        Dim age As Integer
        Dim ag As Integer
        'this is called only by Sub MatRateCheck
        '..... Compute Adult Equivalence Factors
        For ag = age To 2 Step -1
            AdltEqv(ag, stk) = maturationRate(ag, stk) + ((1 - maturationRate(ag, stk)) * survivalRate(ag + 1) * AdltEqv(ag + 1, stk))
        Next ag
        'get % age 1 to adult survival rate
        Age1toAdultSurvivalRate(stk) = FRecruitRate(stk)
    End Sub

    '--------------------------------------------------------------
    Sub CalcShakersAndCNR(ByRef location As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Calculates the number of fish lost to incidental
        '            mortality.
        '
        '  Arguments: Fish% = Fishery index number
        '             location  = Terminal/Pretreminal
        '             FracNV!() = Proportion of population not
        '                     vulnerable to fishery by stock and age
        '             Yr%   = Simulation year
        '
        '  Inputs:    CNRData()
        '             CNRFlag%()
        '                0 = no CNR fishery
        '                1 = CNR fishery
        '             CNRMethod%()
        '                0 = RT Method Model ceiling ratio
        '                1 = Season length
        '                2 = reported encounters
        '                    NOTE: input values must be scaled by
        '                    model catch/true catch
        '             dropoffRate()
        '             encounterRatio()
        '             legalCatch()
        '             legalReleaseMortRate()
        '             pointerCNR%()
        '             RT()
        '             sublegalReleaseMortRte()
        '             TempCat
        '
        '  Called By: calcPreterminalHarvest
        '             calcTotalFishingMortalities
        '
        '  Output:    CNRLegal()
        '             CNRShakCat()
        '             StkShakCat()
        '
        '  Externals: None.
        '
        '--------------------------------------------------------------

        Dim Adjust, DropRate As Double
        Dim age, CNRIndx As Integer
        Dim ERIt, adjustBaseHR As Double
        Dim fish, ICheck As Integer
        Dim instBaseHR As Double
        Dim legalCNRMortRate As Double
        Dim RT_workingCopy As Double
        Dim stk As Integer
        Dim subLegalCNRscalar, LegalCNRscalar As Double
        Dim TestFish, TestYr As Integer
        Dim CNRsubLegalRate, StkShakCatWithOutDropOffs, CNRlegalRate As Double
        Dim TotShak As Double

        For fish = 1 To numFisheries
            If timeStepFlag(fish, timeStep) > 0 Then
                RT_workingCopy = RT(yr, location, fish)
                isCNR = False
                CNRsubLegalRate = 0
                CNRlegalRate = 0

                '...... eq 44
                ShakMortRate(fish) = sublegalReleaseMortRate(fish, yr) + dropoffRate(fish, yr)

                '...... eq 45
                TotShak = ShakMortRate(fish) * encounterRatio(location, fish) * tempCat(location, fish)
                '...... so that takes care of sublegals during legal season??

                If CNRFlag(yr, fish) = 1 Then
                    '..... Calculate CNR Mortality losses
                    isCNR = True
                    CNRIndx = pointerCNR(fish)
                    Select Case CNRMethod(yr, CNRIndx)
                        Case 0 '....RT or Harvest Ratio Method
                            If RT_workingCopy > 0 And RT_workingCopy < 1 Then
                                'eq 81
                                '..... SubLegals
                                CNRsubLegalRate = CNRSelectivity(0, CNRIndx) * (1 - RT_workingCopy) / RT_workingCopy
                                '..... CNRSelectivity is the selectivity factors
                                'compute legal IM+ dropoff 
                                'eq 83
                                legalCNRMortRate = legalReleaseMortRate(fish, yr) + dropoffRate(fish, yr)

                                'eq 82
                                '..... Legals
                                CNRlegalRate = CNRSelectivity(1, CNRIndx) * legalCNRMortRate * (1 - RT_workingCopy) / RT_workingCopy
                            End If

                        Case 1 '.... use ratio of CNR/Regular Season
                            'eq 84
                            CNRsubLegalRate = CNRSelectivity(0, CNRIndx) * CNRdata(yr, CNRIndx, 1) / CNRdata(yr, CNRIndx, 0) '..... SubLegal
                            'eq 83
                            'legalCNRMortRate = legalReleaseMortRate(fish%)
                            legalCNRMortRate = legalReleaseMortRate(fish, yr) + dropoffRate(fish, yr)
                            'changed to legalCNRMortRate = legalReleaseMortRate(fish%)+ dropoffRate(fish%) after side by side test with QB

                            'eq 85
                            CNRlegalRate = CNRSelectivity(1, CNRIndx) * legalCNRMortRate * CNRdata(yr, CNRIndx, 1) / CNRdata(yr, CNRIndx, 0) '..... Legal
                            '.....above here have legal morts done for CNR including CNR legal Dropoff
                        Case 2 '....Reported Encounter Method
                            'eq 86
                            If TotShak > 0 Then CNRsubLegalRate = (tempCat(location, fish) / TotShak) * ShakMortRate(fish) * (CNRdata(yr, CNRIndx, 1) / CNRdata(yr, CNRIndx, 2))
                            'eq 83
                            legalCNRMortRate = legalReleaseMortRate(fish, yr) + dropoffRate(fish, yr)
                            'eq 87
                            CNRlegalRate = legalCNRMortRate * CNRdata(yr, CNRIndx, 0) / CNRdata(yr, CNRIndx, 2)
                    End Select
                End If 'If CNRFlag%(yr%, fish%) = 1
                For stk = 1 To numStocks
                    For age = 2 To NumAges
                        If FTerminal(stk, fish, age) = location Then
                            '..... Distribute shakers across stocks and ages
                            'eq 49
                            stkShakCat(location, fish, stk, age) = fracNV(fish, stk, age) * TotShak
                            StkShakCatWithOutDropOffs = stkShakCat(location, fish, stk, age)
                            'add legal catch dropoff morts here
                            DropRate = dropoffRate(fish, yr)
                            'eq 51
                            legalDropOffs(location, fish, stk, age) = legalCatch(location, fish, stk, age) * DropRate
                            'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, StkShakCat(location, fish%, stk%, age%); legalDropOffs; "shakers CNR"; fish%; age%; location; FracNV!(fish%, stk%, age%); TotShak
                            'eq 50
                            stkShakCat(location, fish, stk, age) = stkShakCat(location, fish, stk, age) + legalDropOffs(location, fish, stk, age)
                            If isCNR = True Then
                                '*************** 2/96 CNR Adjustment for Multiple Encounters ************
                                '..... Compute adjustment to shakers due to multiple encounters
                                '..... For description of method, see Morishima memo of 2/21/96
                                Select Case CNRMethod(yr, CNRIndx)
                                    '..... For methods 0 and 1, adjustment is the proportion of season operating under CNR
                                    '      multiplied by the number of potential reencounter periods
                                    Case 0 '....RT Method
                                        If RT_workingCopy > 0 And RT_workingCopy < 1 Then
                                            'eq 102
                                            instBaseHR = -System.Math.Log(1 - harvestRate(fish, age, stk)) / baseEncounterPeriods(CNRIndx)
                                            'eq 103
                                            Adjust = (1 - RT_workingCopy) * baseEncounterPeriods(CNRIndx)
                                            'eq 104
                                            ERIt = (1 - System.Math.Exp(-instBaseHR * Adjust))
                                            If harvestRate(fish, age, stk) > 0 Then
                                                'eq 105
                                                subLegalCNRscalar = ERIt / (harvestRate(fish, age, stk) * (1 - RT_workingCopy))
                                                LegalCNRscalar = subLegalCNRscalar
                                            Else
                                                subLegalCNRscalar = 0
                                            End If
                                            Adjust = subLegalCNRscalar
                                        Else
                                            Adjust = 1
                                        End If
                                    Case 1 '....Season Length
                                        'eq 108
                                        Adjust = CNRdata(yr, CNRIndx, 1) / (CNRdata(yr, CNRIndx, 0) + CNRdata(yr, CNRIndx, 1))
                                        'eq 109 sublegal
                                        adjustBaseHR = Adjust * harvestRate(fish, age, stk) * CNRSelectivity(0, CNRIndx)
                                        'eq 110 1 of 2 instances
                                        instBaseHR = -System.Math.Log(1 - adjustBaseHR) / baseEncounterPeriods(CNRIndx)
                                        'eq 111
                                        ERIt = (1 - System.Math.Exp(-instBaseHR * legalCNRMortRate * Adjust * baseEncounterPeriods(CNRIndx)))
                                        If adjustBaseHR > 0 Then
                                            'eq 112
                                            subLegalCNRscalar = ERIt / (adjustBaseHR * legalCNRMortRate)
                                        Else
                                            subLegalCNRscalar = 0
                                        End If
                                        'eq 113 legal
                                        adjustBaseHR = Adjust * harvestRate(fish, age, stk) * CNRSelectivity(1, CNRIndx)
                                        'eq 110 2 of 2 instances
                                        instBaseHR = -System.Math.Log(1 - adjustBaseHR) / baseEncounterPeriods(CNRIndx)
                                        ERIt = (1 - System.Math.Exp(-instBaseHR * legalCNRMortRate * Adjust * baseEncounterPeriods(CNRIndx)))
                                        If adjustBaseHR > 0 Then
                                            LegalCNRscalar = ERIt / (adjustBaseHR * legalCNRMortRate)
                                        Else
                                            LegalCNRscalar = 0
                                        End If
                                    Case Else '..... Estimated Encounters (no change cause multiple encounters are imbedded in estimate)
                                        Adjust = 1
                                End Select
                                ' so if adjust >1 then did multiple encounters
                                If Adjust > 1 Then
                                    '..... Estimate CNR mortality losses
                                    'eq 106
                                    CNRShakCat(location, CNRIndx, stk, age) = StkShakCatWithOutDropOffs * CNRsubLegalRate * subLegalCNRscalar
                                    'eq 107
                                    CNRLegal(location, CNRIndx, stk, age) = legalCatch(location, fish, stk, age) * CNRlegalRate * LegalCNRscalar
                                Else
                                    '..... Estimate CNR mortality losses
                                    '.....recall that StkShakCatWithOutDropOffs has TotShak (hence IM rate) already in it
                                    '.....eq 88
                                    CNRShakCat(location, CNRIndx, stk, age) = StkShakCatWithOutDropOffs * CNRsubLegalRate
                                    '.....eq 89
                                    CNRLegal(location, CNRIndx, stk, age) = legalCatch(location, fish, stk, age) * CNRlegalRate
                                End If
                            End If
                        End If
                    Next age
                Next stk
            End If 'timeStepFlag(fish%, timeStep%) > 0
        Next fish
        Exit Sub

        '****************** END Multiple Encounter Modifications ***********

    End Sub

    '--------------------------------------------------------------
    Sub CalcHarvestStockComposition()
        '--------------------------------------------------------------
        '  Purpose:  Computes stock composition estimates for model
        '            stocks (% of catch by stock) in specified
        '            fisheries.
        '
        '  Arguments: Yr% = Simulation Year
        '
        '  Inputs:    NumStkCatch
        '             numStocks
        '             PreTerm%
        '             TempCat()
        '             Term%
        '
        '
        '  Output:    StkProp()
        '
        '  Externals: None.
        '
        '--------------------------------------------------------------

        Dim i, fish, stk As Integer
        Dim TotCatFish As Double

        For i = 1 To NumStkCatch
            '..... Determine fishery number
            fish = StkCompFish(i)
            '..... Compute total preterminal and terminal catch
            TotCatFish = tempCat(PreTerm, fish) + tempCat(Term, fish)
            If TotCatFish > 0 Then
                For stk = 1 To numStocks
                    '..... Compute stock compositions
                    stkProp(yr, stk, i) = (TempStkCat(PreTerm, fish, stk) + TempStkCat(Term, fish, stk)) / TotCatFish
                Next stk
            End If
        Next i

    End Sub

    '--------------------------------------------------------------
    Sub CalibCheck(ByRef year_ As Integer, ByRef Ctl As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Generate reports comparing age-specific
        '            calibration and model-generated data.
        '            'calibCheck (evaluate calibrations) is NOT the same as chkClb (compares projection runs)
        '
        '  Arguments: Cntl% = Control variable
        '                     1 = Set Up Files For Output
        '                     2 = Close files
        '                     3 = Not Used
        '                     4 = Produce comparison data for report
        '             year_%   = Year of simulation
        '
        '  Inputs:    TotalEscape_EV()
        '             EVType%()
        '             numStocksUsingSameEV%()
        '             EVageFlag%()
        '             GSH_Cat_EV()
        '             OceanEscape_EV()
        '
        '  Called By: calcEscapement
        '
        '  Output:    Disk file containing detailed statistics
        '
        '  Externals: None.
        '
        '--------------------------------------------------------------
        Dim st, age, i, stk As Integer
        Dim fileExists, file1 As String
        Dim temp As Double

        Select Case Ctl
            Case 1 '..... Open Calibration Data File

                file1 = saveFilePrefix & "CLB.DAT"
                'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC/commoner/redir/redirect.htm?keyword="vbup1041"'
                fileExists = Dir(file1)
                'check to see if file exists
                If fileExists <> "" Then Kill(file1)
                CLB_DATfileID = FreeFile()
                FileOpen(CLB_DATfileID, file1, OpenMode.Output)

            Case 2 '..... Close Calibration Data File
                FileClose(CLB_DATfileID)

            Case 3 '..... Unused

            Case 4 '..... Dump Relevant statistics to file for comparison
                For stk = 1 To numStocks
                    Select Case EVType(stk)
                        Case 1 '..... Adult escapement
                            PrintLine(CLB_DATfileID, year_)
                            For age = StartAge(stk) To NumAges
                                PrintLine(CLB_DATfileID, totalEscape_EV(year_, stk, age))
                            Next age

                        Case 2 '..... Terminal Run Size
                            PrintLine(CLB_DATfileID, year_)
                            For age = StartAge(stk) To NumAges
                                PrintLine(CLB_DATfileID, OceanEscape_EV(year_, stk, age))
                            Next age

                        Case 3 '..... Georgia Strait Hatchery Catch
                            PrintLine(CLB_DATfileID, year_)
                            For age = StartAge(stk) To NumAges
                                PrintLine(CLB_DATfileID, GSH_Cat_EV(year_, age))
                            Next age

                        Case 4 '..... Terminal Run Sizes for more than one stock combined
                            PrintLine(CLB_DATfileID, year_)
                            For age = StartAge(stk) To NumAges
                                temp = 0
                                For i = 1 To numStocksUsingSameEV(stk, 0)
                                    '..... Get stock pointer
                                    st = numStocksUsingSameEV(stk, i)
                                    temp = temp + OceanEscape_EV(year_, st, age)
                                    '..... Hardwire for RBT and RBH stocks, data not avail prior to 1985
                                    If stk = 4 And year_ < 5 Then GoTo DumpIt
                                Next i
DumpIt:
                                PrintLine(CLB_DATfileID, temp)
                            Next age
                        Case Else
                    End Select
                Next stk
        End Select
    End Sub

    '--------------------------------------------------------------
    Sub Calibrate(ByRef Ctl As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Controls calibration procedure.
        '
        '  Arguments: Ctl% = Control variable
        '
        '  Inputs:    AdjMatRate%
        '             Age%
        '             BiggestProb
        '             BigProb
        '             BrdYr%
        '             ConvFlag%
        '             EV_firstYr%
        '             Iter
        '             MatIter
        '             maturationRate()
        '             NameRun$
        '             NumEVBroodYr()
        '             numStocks
        '             ProbBY
        '             ProbStock$
        '             SaveFilePrefix$
        '             Stk%
        '             EV()
        '
        '
        '  Output:    ConTolerance
        '             EV()
        '
        '  Externals: DATE$
        '             CheckEV
        '             LEFT$
        '             Write_EV_File
        '
        '
        '--------------------------------------------------------------

        Dim file As String
        Dim stk As Integer

        Select Case Ctl
            Case 1
                '..... Open file to record history of EV estimates
                file = saveFilePrefix & "EVS.REC"
                EVS_RECfileID = FreeFile()
                FileOpen(EVS_RECfileID, file, OpenMode.Output)
                '..... Dump unchanging information
                PrintLine(EVS_RECfileID, EV_firstYr)
                For stk = 1 To numStocks
                    PrintLine(EVS_RECfileID, stk, NumEVBroodYr(stk))
                Next stk
                '..... Open file to record calibration history
                file = saveFilePrefix & "CLB.HIS"
                CLB_HISfileID = FreeFile()
                FileOpen(CLB_HISfileID, file, OpenMode.Output)
            Case 2
                file = saveFilePrefix & "HIS.REC"
                HIS_RECfileID = FreeFile()
                FileOpen(HIS_RECfileID, file, OpenMode.Output)
                '..... Record date and time of calibration
                PrintLine(HIS_RECfileID, "Date of EV Calculations: " & DateString)
                PrintLine(HIS_RECfileID, "Time of EV Calculations: " & Left(TimeString, 5))
                PrintLine(HIS_RECfileID)
                '..... Record run title
                PrintLine(HIS_RECfileID, NameRun)
                PrintLine(HIS_RECfileID)

                '..... Erase unnecessary array
                Erase AI_fisheryIndex
                'Erase abundCohrt
                'Erase relAbund

            Case 3
                '..... Calibration completed.  Finish housekeeping.
                FileClose(HIS_RECfileID)
                '.....  Reset Iter to total iterations
                Iter = Iter - 1
                '..... Write EV file from calibration
                Call Write_EV_File()

        End Select

    End Sub

    '--------------------------------------------------------------
    Sub CalcLegalHarvest(ByRef age As Integer, ByRef stk As Integer, ByRef location As Integer, ByRef Chort2(,) As Double)
        '--------------------------------------------------------------
        '  Purpose:  Computes legal catch by stock and age and accumulates total catch
        '            by fishery.
        '
        '  Arguments: Age%     = Age of fish
        '             Stk%     = Stock index number
        '             location     = Terminal/PreTerminal
        '             Chort2() = scrCohrt() if this sub is called by CalcPreTerminalHarvest
        '             Chort2() = termRun() if this sub is called CalcTerminalHarvest
        '             yr% = year, one of the indices for FP()
        '
        '  Inputs:    FP()
        '             harvestRate()
        '             PNV()
        '
        '  Called By: calcPreterminalHarvest
        '             CalcTerminalHarvest
        '
        '  Output:    legalCatch()
        '             TempCat()
        '             TempStkCat()
        '
        '  Externals: FTerminal
        '
        '--------------------------------------------------------------

        Dim fish As Integer
        Dim F As Double

        For fish = 1 To numFisheries
            If timeStepFlag(fish, timeStep) > 0 Then
                If harvestRate(fish, age, stk) > 0 Then
                    If FTerminal(stk, fish, age) = location Then

                        'the following is for simple linear rates
                        'comment out this line if instantaneous rates are being used
                        F = harvestRate(fish, age, stk) * (1 - PNV(fish, age))

                        'the following is for instantaneous rate
                        'If numTimeStepsActuallyFished(fish%) = 1 Then F = harvestRate(fish%, age%, stk%) * (1 - PNV(fish%, age%))
                        'If numTimeStepsActuallyFished(fish%) > 1 Then
                        'instantaneous mortality rate = F
                        '        F = -(Log(1 - harvestRate(fish%, age%, stk%)) / numTimeStepsActuallyFished(fish%))
                        'mortality rate with vulnerability = F * (1-PNV)
                        '        F = F * (1 - PNV(fish%, age%))
                        'harvest rate = 1 - survival rate = 1 - exp(-F)
                        '        F = 1 - Exp(-F)
                        'End If

                        'Compute legal-size pre-terminal catches
                        'Cohort * Harvest Rate * Fishery Policy * Proportion Vulnerable
                        '1/96 Age-Specific FPs
                        'Chort2() = scrCohrt() if this sub is called by CalcPreTerminalHarvest
                        'Chort2() = termRun() if this sub is called CalcTerminalHarvest

                        'eq 12
                        legalCatch(location, fish, stk, age) = Chort2(age, stk) * F * FP(yr, stk, fish, age)


                        'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, Chort2(age%, stk%); F; FP(yr%, stk%, fish%, age%); legalCatch(location, fish%, stk%, age%); "legal harvest"; fish%; location; numTimeStepsActuallyFished(fish%)
                        'If stk% = 20 Then Print #99, yr%; fish%; age%; Chort2(age%, stk%); legalCatch(location, fish%, stk%, age%)

                        'eq 46
                        tempCat(location, fish) = tempCat(location, fish) + legalCatch(location, fish, stk, age)
                        TempStkCat(location, fish, stk) = TempStkCat(location, fish, stk) + legalCatch(location, fish, stk, age)


                        'If yr = 0 And location = 0 Then Print #99, yr%; stk%; age%; fish%; timeStep%; legalCatch(location, fish%, stk%, age%); Chort2(age%, stk%); harvestRate(fish%, age%, stk%); PNV(fish%, age%); "21pretermCat"
                        'If yr = 0 And location = 1 Then Print #99, yr%; stk%; age%; fish%; timeStep%; legalCatch(location, fish%, stk%, age%); Chort2(age%, stk%); harvestRate(fish%, age%, stk%); PNV(fish%, age%); "41termCat"



                    End If 'FTerminal(stk%, fish%, age%) = location
                End If 'harvestRate(fish%, age%, stk%) > 0
            End If 'timeStepFlag(fish%, timeStep%) > 0
        Next fish

    End Sub


    '--------------------------------------------------------------
    Sub CalcPreterminalHarvest()
        '--------------------------------------------------------------
        '  Purpose:  Clears array contents, updates scratch cohorts,
        '            computes catches by preterminal fisheries,
        '            controls ceiling setup, and calls terminal
        '            catch routines.
        '
        '  Arguments: Yr% = Simulation year
        '
        '  Inputs:    CNRLegal()
        '             CNRShakCat()
        '             Cohort_()
        '             encounterRatio()
        '             CeilingEnd
        '             MaxAges
        '             legalCatch()
        '             NumAges
        '             NumCeil
        '             numStocks
        '             OcnAECat()
        '             PreTerm%
        '             ScrCohrt()
        '             StartCeil
        '             StkShakCat()
        '             survivalRate()
        '             TotAECat()
        '             TempStkCat()
        '             TermRun()
        '
        '  Externals: CalcencounterRatio
        '             calcLegalHarvest
        '             CalcTerminalHarvest
        '             ceilingSetup
        '             FActive
        '
        '
        '--------------------------------------------------------------

        Dim AdltEqvCatch, CNRTot As Double
        Dim fish, age, CNRIndx, stk As Integer
        Dim preterminalModelMortality, TotPreTermCat As Double

        '..... Reset arrays to zero at the start of each time step
        ReDim CNRLegal(1, NumCNRFish, numStocks, NumAges)
        ReDim CNRShakCat(1, NumCNRFish, numStocks, NumAges)
        ReDim encounterRatio(1, numFisheries)
        ReDim legalCatch(1, numFisheries, numStocks, NumAges)
        ReDim legalDropOffs(1, numFisheries, numStocks, NumAges)
        ReDim OcnAECat(numStocks)
        ReDim RT_temp(1, numFisheries)
        ReDim stkShakCat(1, numFisheries, numStocks, NumAges)
        ReDim tempCat(1, numFisheries)
        ReDim TempStkCat(1, numFisheries, numStocks)
        ReDim TermRun(NumAges, numStocks)
        ReDim totAECat(numStocks)

        '..... Reset arrays to zero at the start of each year
        'If timeStep% = 1 Then
        '    ReDim TotalTermRun( NumAges,  NumStocks)
        'End If 'If timeStep% = 1

        'scrCohrt is working copy Cohort_.  Apply natural mortality, fishing mortality, etc
        'to scrCohrt.  Test if ratio of modeled to observed ceiling fishery < cutOffLevel
        'if not, do not save scrCohrt.  Go back and start over again with another scratch
        'copy of Cohort_.  Cohort_ is unchanged just in case the iterative search fails
        'and you need to start again.
        'When ratio of modeled to observed ceiling fishery is < cutOffLevel
        'then save scrCohrt as Cohort_ in Sub CalcTerminalRunAndEscapement (which is after
        'CalcTerminalHarvest.


        For stk = 1 To numStocks
            For age = 2 To NumAges
                'get a working copy of Cohort_()
                '..... get a copy of cohort(at start of this year) = cohort(at end of previous year) * survivalRate
                scrCohrt(age, stk) = cohort_(age, stk)
                'PrintLine(traceFileID, "caclPerterm", age, stk, scrCohrt(age, stk))
                If timeStep = 1 Then timeStep1Cohort(age, stk) = scrCohrt(age, stk)
                '..... Compute legal-size pre-terminal catches eq 11
                Call CalcLegalHarvest(age, stk, PreTerm, scrCohrt)
            Next age
        Next stk

        '..... Check for ceiling management option
        If NumCeil > 0 And FActive(startCeil, CeilingEnd) Then
            '..... Set ceiling level
            Call CeilingSetup(PreTerm)
        End If

        '..... Calculate shaker encounter rates
        Call CalcEncounterRatio(PreTerm)

        '..... Compute shaker loss
        Call CalcShakersAndCNR(PreTerm)

        'calculate total preterminal catch and subtract to obtain terminal run before maturation
        For stk = 1 To numStocks
            '..... Set accumulator to zero
            AdltEqvCatch = 0
            For age = 2 To NumAges
                '..... Compute total preterminal catch + shakers for stock
                TotPreTermCat = 0
                For fish = 1 To numFisheries
                    If timeStepFlag(fish, timeStep) > 0 Then
                        If FTerminal(stk, fish, age) = PreTerm Then
                            If CNRFlag(yr, fish) = 1 Then
                                '..... Determine CNR Fishery
                                CNRIndx = pointerCNR(fish)
                                '..... Total CNR Legal & Sublegal Mortality
                                CNRTot = CNRLegal(PreTerm, CNRIndx, stk, age) + CNRShakCat(PreTerm, CNRIndx, stk, age)
                            Else
                                CNRTot = 0
                            End If
                            '..... Compute preterminal mortality eq 133
                            preterminalModelMortality = legalCatch(PreTerm, fish, stk, age) + stkShakCat(PreTerm, fish, stk, age) + CNRTot
                            AdltEqvCatch = AdltEqvCatch + preterminalModelMortality * AdltEqv(age, stk)
                            'accum preterminal mortality eq 132
                            TotPreTermCat = TotPreTermCat + preterminalModelMortality
                        End If
                    End If 'timeStepFlag(fish%, timeStep%) > 0
                Next fish

                '..... Compute # fish after pre-terminal catch
                '..... i.e. scrCohrt = starting cohort size - preterm catch only
                'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, ScrCohrt(age%, stk%); "preterm catch"; TotPreTermCat; "scrCohrt after preterminal catch";
                'eq 131
                scrCohrt(age, stk) = FMinZed(scrCohrt(age, stk) - TotPreTermCat)
                'PrintLine(traceFileID, "caclPerterm2", age, stk, scrCohrt(age, stk))
                'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, ScrCohrt(age%, stk%)
            Next age
            '..... Compute total and ocean adult equivalent catches
            totAECat(stk) = totAECat(stk) + AdltEqvCatch
            OcnAECat(stk) = totAECat(stk)
        Next stk

    End Sub

    '--------------------------------------------------------------
    Sub CalcTerminalHarvest()
        '--------------------------------------------------------------
        '  Purpose:  Calculates (for individual stocks) interim
        '            catches by terminal fisheries, adult equivalent
        '            exploitation rates, and escapements, and then
        '            updates the cohort size.
        '
        '  Arguments: Yr% = Simulation year
        '
        '  Inputs:    AdltEqv()
        '             BaseCeilStart
        '             CNRFlag%()
        '             CNRLegal()
        '             CNRShakCat()
        '             CeilingEnd
        '             legalCatch()
        '             NumAges
        '             numFisheries
        '             numStocks
        '             isOceanTerminal
        '             PreTerm%
        '             pointerCNR%()
        '             StartCeil
        '             StkShakCat()
        '             isCeiling
        '
        '  Called By: calcPreterminalHarvest
        '
        '  Output:    OcnAECat()
        '             ScrCohrt()
        '             TermRun()
        '             TotAECat()
        '             TotPreTermCat
        '
        '  Externals: calcLegalHarvest
        '             CeilingTestIfModeledCatchMatchesObserved
        '             ceilingSetup
        '             FActive
        '             FTerminal
        '
        '
        '--------------------------------------------------------------


        Dim age, stk As Integer


        For stk = 1 To numStocks
            For age = 2 To NumAges
                'the following assumes the first preterminal fishery is in time step 2
                'the following if statement has to be change if that is no longer true
                If timeStep = 2 Then timeStep1TermRun(age, stk) = TermRun(age, stk)
                '..... calculate terminal catch eq 141
                Call CalcLegalHarvest(age, stk, Term, TermRun)
            Next age
        Next stk

        'ceiling fishery
        If isCeiling = True And yr > BaseCeilStart Then
            Call CeilingSetup(Term)
            If FActive(startCeil, CeilingEnd) And isOceanTerminal = True Then
                'eq 142.....call subroutine to compare CurrentCatch / TotCeil ratio with cutOffLevel
                Call CeilingTestIfModeledCatchMatchesObserved()
            End If
        End If

        'note the next 2 lines of code could be placed here instead of
        'calcTotalFishingMortalities
        'Shakers and CNR are not used in the ceiling harvest calculations
        'To speed up the program it is placed in calcTotalFishingMortalities
        'so it is not called over and over
        'again in the iterative process to find the perterminal ceiling harvest rates

        'Call CalcencounterRatio(Term%, yr%)
        '..... Compute shaker loss
        'Call CalcShakersAndCNR(Term%, yr%)


    End Sub


    '--------------------------------------------------------------
    Sub CheckEV()
        '--------------------------------------------------------------
        '  Purpose:  Sets calibration data for each stock and then
        '            calls CalcEV to compute EV scalars.
        '
        '  Arguments: None.
        '
        '  Inputs:    TotalEscape_EV()
        '             Age%
        '             ConvFlag%
        '             EVType%()
        '             I%
        '             InStk
        '             NumAges
        '             numStocks
        '             NumYears
        '             numStocksUsingSameEV%()
        '             StartAge%()
        '             Stk%
        '             GSH_Cat_EV()
        '             StckSclr()
        '             OceanEscape_EV()
        '             year_%
        '
        '
        '  Output:    OtherDat()  temp array shared with
        '                         calcEV and calcMaturityRateAdjustments
        '
        '  Externals: CalcEV
        '
        '
        '--------------------------------------------------------------

        Dim stk, i, age, inStk, year_ As Integer

        Select Case AdjMatRate
            Case 0
                '..... Do not adjust maturation rate schedules
                biggestProb = 0
                BigProb = 0
                probBY = 0
                ProbStock = ""
                '..... Compute new EV scalars
                Call CalcEVFactors()
                '..... Dump new EV scalars to disk
                Call CalibDump()

            Case Is > 0
                '..... Adjust Maturation Schedules
                '..... Set Convergence Flag for Maturation Rates = 1 (yes)
                ConvFlagMR = 1
                '..... Reset variables to identify stock & BY with biggest convergence problem
                biggestProb = 0
                BigProb = 0
                probBY = 0
                ProbStock = ""
                '..... Compute new EV scalars
                Call CalcEVFactors()
                '..... Dump EV Scalars to disk
                Call CalibDump()
                If ConvFlag = 1 Then Exit Sub
                '..... Go through time loop with revised EV scalars
                Call AnnualLoop()
                '..... Compute new Maturation Rates
                For stk = 1 To numStocks
                    If StkMat(stk) = 1 Or StkMat(stk) > 2 Then
                        '..... Adjust maturation rates for this stock
                        '****** 2/26/96 reduced dim to eliminate multi-comparisons ******
                        ReDim otherDat(NumYears + 5, NumAges)
                        Select Case EVType(stk)

                            Case 1 '..... Adult escapement only/start @ appropriate age
                                For year_ = 0 To NumYears
                                    For age = 2 To NumAges
                                        otherDat(year_, age) = totalEscape_EV(year_, stk, age)
                                    Next age
                                Next year_
                                Call CalcMaturityRateAdjustments(StartAge(stk), stk)

                            Case 2 '..... Total terminal run/start @ appropriate age
                                For year_ = 0 To NumYears
                                    For age = 2 To NumAges
                                        otherDat(year_, age) = OceanEscape_EV(year_, stk, age)
                                    Next age
                                Next year_
                                Call CalcMaturityRateAdjustments(StartAge(stk), stk)

                            Case 3 '..... Georgia St Hatchery /GS catches only/start @ appropriate age
                                'note: EVType 3 and 4 is not used
                                '*******2/24/96  ADDED TO PROVIDE FOR CORRECT CALIBRATION TO ESCAPEMENT ***********
                            Case 4 '..... Combine more than one set of data/ESCAPEMENT/start @ appropriate age
                                'note: EVType 3 and 4 is not used
                                For year_ = 0 To NumYears
                                    For i = 1 To numStocksUsingSameEV(stk, 0)
                                        inStk = numStocksUsingSameEV(stk, i)
                                        For age = 2 To NumAges
                                            otherDat(year_, age) = otherDat(year_, age) + totalEscape_EV(year_, inStk, age)
                                        Next age
                                    Next i
                                Next year_
                                Call CalcMaturityRateAdjustments(StartAge(stk), stk)
                                '..... Set maturation rates etc. for all stocks to computed values
                                For i = 2 To numStocksUsingSameEV(stk, 0)
                                    inStk = numStocksUsingSameEV(stk, i)
                                    For age = 3 To 4
                                        maturationRate(age, inStk) = maturationRate(age, stk)
                                    Next age
                                    Age1toAdultSurvivalRate(inStk) = Age1toAdultSurvivalRate(stk)
                                Next i

                                '****** 2/24/96 ADDED TO PROVIDE FOR CORRECT CALIBRATION TO TERMINAL RUN
                            Case 5 '..... Combine more than one set of data/terminal run/start @ appropriate age
                                For year_ = 0 To NumYears
                                    For i = 1 To numStocksUsingSameEV(stk, 0)
                                        inStk = numStocksUsingSameEV(stk, i)
                                        For age = 2 To NumAges
                                            otherDat(year_, age) = otherDat(year_, age) + OceanEscape_EV(year_, inStk, age)
                                        Next age
                                    Next i
                                Next year_
                                Call CalcMaturityRateAdjustments(StartAge(stk), stk)
                                '..... Set maturation rates etc. for all stocks to computed values
                                For i = 2 To numStocksUsingSameEV(stk, 0)
                                    inStk = numStocksUsingSameEV(stk, i)
                                    For age = 3 To 4
                                        maturationRate(age, inStk) = maturationRate(age, stk)
                                    Next age
                                    Age1toAdultSurvivalRate(inStk) = Age1toAdultSurvivalRate(stk)
                                Next i

                            Case Else
                        End Select
                    End If
                Next stk

        End Select
    End Sub

    Sub CalcEVFactors()
        'this is called only by Sub CheckEV
        Dim stk, i, age, inStk, year_ As Integer

        For stk = 1 To numStocks
            '******** 2/96 eliminated array PaulDump and reduced size of OtherDat
            ReDim otherDat(NumYears + 5, NumAges)
            'EV factors are based on ratio_S() of observed to model data
            'EVtype%(stk%) indicates what type of data to use for ratio_S() (Table 3.10 p 47 1991 PSC CMD)
            Select Case EVType(stk)

                Case 1 '..... Adult escapement only/start @ appropriate age
                    For year_ = 0 To NumYears
                        For age = 2 To NumAges
                            otherDat(year_, age) = totalEscape_EV(year_, stk, age) 'from eq 165
                        Next age
                    Next year_
                    Call CalcEV(StartAge(stk), stk)

                Case 2 '..... Total terminal run/start @ appropriate age
                    For year_ = 0 To NumYears
                        For age = 2 To NumAges
                            otherDat(year_, age) = OceanEscape_EV(year_, stk, age)
                        Next age
                    Next year_
                    Call CalcEV(StartAge(stk), stk)

                Case 3 '..... Georgia St Hatchery /GS catches only/start @ appropriate age
                    'note: EVType 3 and 4 is not used
                    For year_ = 0 To NumYears
                        For age = 2 To NumAges
                            otherDat(year_, age) = GSH_Cat_EV(year_, age)
                        Next age
                    Next year_
                    Call CalcEV(StartAge(stk), stk)

                    '********* 2/24/96 ADDED TO PROVIDE FOR CORRECT CALIBRATION TO ESCAPEMENT *****
                Case 4 '..... Combine more than one set of data/ESCAPEMENT/start @ appropriate age
                    'note: EVType 3 and 4 is not used
                    For year_ = 0 To NumYears
                        For i = 1 To numStocksUsingSameEV(stk, 0)
                            inStk = numStocksUsingSameEV(stk, i)
                            For age = 2 To NumAges
                                otherDat(year_, age) = otherDat(year_, age) + totalEscape_EV(year_, inStk, age)
                            Next age
                        Next i
                    Next year_
                    Call CalcEV(StartAge(stk), stk)
                    '..... Set EV factors for all stocks in group
                    '***** 2/22/96 Modified to allow broods prior to first model year ***********
                    For year_ = -NumAges To MaxYrs
                        For i = 2 To numStocksUsingSameEV(stk, 0)
                            inStk = numStocksUsingSameEV(stk, i)
                            EV(year_ + NumAges, inStk) = EV(year_ + NumAges, stk)
                        Next i
                    Next year_

                    '*****2/24/96 ADDED TO PROVIDE FOR CORRECT CALIBRATION TO TERMINAL RUN **************
                Case 5 '..... Combine more than one set of data/terminal run/start @ appropriate age
                    For year_ = 0 To NumYears
                        For i = 1 To numStocksUsingSameEV(stk, 0)
                            inStk = numStocksUsingSameEV(stk, i)
                            For age = 2 To NumAges
                                otherDat(year_, age) = otherDat(year_, age) + OceanEscape_EV(year_, inStk, age)
                            Next age
                        Next i
                    Next year_
                    Call CalcEV(StartAge(stk), stk)
                    '..... Set EV factors for all stocks in group
                    '***** 2/22/96 Modified to allow broods prior to first model year ***********
                    For year_ = -NumAges To MaxYrs
                        For i = 2 To numStocksUsingSameEV(stk, 0)
                            inStk = numStocksUsingSameEV(stk, i)
                            EV(year_ + NumAges, inStk) = EV(year_ + NumAges, stk)
                        Next i
                    Next year_
                Case Else
            End Select

        Next stk

    End Sub


    Sub CalibDump()
        'this is called only by Sub CheckEV
        Dim BrdYr, stk As Integer

        '..... Dump EV estimates to disk
        For stk = 1 To numStocks
            If NumEVBroodYr(stk) Then
                PrintLine(EVS_RECfileID, stk)
                For BrdYr = EV_firstYr To NumEVBroodYr(stk)
                    PrintLine(EVS_RECfileID, BrdYr, EV(BrdYr + NumAges, stk), Iter)
                Next BrdYr
            End If
        Next stk

    End Sub

    '--------------------------------------------------------------
    Sub CeilingLevel(ByRef NextCeil As Integer, ByRef location As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Computes observed Preterminal and terminal
        '            ceiling levels relative to the base
        '            period.
        '            Note change in ceiling specs to avoid
        '            necessity to compute scalar values.
        '
        '  Arguments: Yr%      = Simulation year
        '             NextCeil = Counter for ceiling number
        '             location     = Terminal/PreTerminal
        '
        '  Inputs:    AddCeil%
        '             BaseCeilingEnd%
        '             BaseCeilStart%
        '             basePeriodToCeilingScalar()
        '             CeilingFisheryFlag()
        '             NumCeil
        '             PreTerm%
        '             TempCat()
        '             Term%
        '
        '  Called By: calcTotalFishingMortalities
        '             ceilingSetup
        '
        '  Output:    Ceil()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        'Note when the CEI file was read by Sub ReadCEIFile, there was
        'a local variable named ceilingCatch which was the aggregate ceiling catch
        'from which the basePeriodtoCeilingScalar was calculated.

        'This subroutine splits the base period catch into perterminal and terminal
        'using the basePeriodtoCeilingScalar
        'and saves the results as CeilLevel(perTerm%,yr,fish%) and CeilLevel(Term%yr,fish%),
        'which is the starting values for Ceil(perTerm%,fish%) and Ceil(Term%,fish%).

        'Note Ceil()is working copy of CeilLevel().
        'The only difference, is CeilLevel has an extra subscript for year.


        'this subroutine does 3 things:
        '(1) accum base period catch during base period years,
        '(2) calculate CeilLevel when last ceiling year is encountered, and
        '(3) set Ceil = CeilLevel for all years after base period.

        'this subroutine is called,
        '(1) after the terminal harvest has been calculated
        'during the base period years to accum the base period catches

        '(2) after the terminal harvest has been calculated
        'when the last base period year is encountered
        'to calculate average base period preterminal and terminal legal catches
        'and set ceiling levels for the first year with ceiling level changes
        '(defined as number of ceiling level changes in rows 6 & 7 in CEI file)

        '(3) after the preterminal legal catch has been calculated
        'and convergence criteria has been satisfied
        'to set ceiling levels for the next year

        'note NextCeil = 1 when this sub is called by CalcTotalFishingMortalities (after calcTerminalHarvest)
        'NextCeil > 1 when this sub is called by CeilingSetup (after preterminal legal catch has been calculated)

        'note yr% <= BaseCeilingEnd% when this sub is called by CalcTotalFishingMortalities (after calcTerminalHarvest)
        'yr% > BaseCeilingEnd% when this sub is called by CeilingSetup (after preterminal legal catch has been calculated)

        'in other words, when this subroutine is called by CalcTotalFishingMortalities,
        'NextCeil = 1 always
        'but yr% will range from 0 to 5
        'when this subroutine is called by CeilingSetup, NextCeil > 2
        'and yr% >= 7

        'Remember ceiling fisheries began after the end of the base period.

        'Remember basePeriodtoCeilingScalar(i%, fish%) = ceilingCatch / BasePeriodCatch
        'is calculated in Sub ReadCEIfile when the CEI file is read, not in this subroutine.


        Dim averageBasePeriodCatch As Double
        Dim i, fish, j As Integer

        '..... do this first
        '..... note this loop must precede follow the previous loop in order for CeilLevel() to be defined

        If yr = BaseCeilingEnd Then 'yr% = BaseCeilingEnd% only happens with calcTotalFishingMortalities, not with CeilingSetup
            For fish = 1 To numFisheries
                If timeStepFlag(fish, timeStep) > 0 Then
                    If CeilingFisheryFlag(fish) = 1 Then
                        For i = 0 To 1
                            For j = 1 To AddCeil + 1
                                CeilLevel(i, j, fish) = 0
                            Next j
                        Next i
                    End If 'CeilingFisheryFlag(fish%) = 1
                End If 'timeStepFlag(fish%, timeStep%) > 0
            Next fish
        End If


        '..... do this during the 1979-1984 ceiling base period years
        'do not confuse with 1979-1982 abundance index base period years
        If NextCeil = 1 Then 'NextCeil = 1 only happens with calcTotalFishingMortalities, not with CeilingSetup

            For fish = 1 To numFisheries
                If timeStepFlag(fish, timeStep) > 0 Then
                    If CeilingFisheryFlag(fish) = 1 Then
                        '..... Set initial ceilings
                        'note the ceiling base period years = 1979-1984
                        'do not confuse with 1979-1982 abundance index base period
                        If FActive(BaseCeilStart, BaseCeilingEnd) Then
                            '..... Accumulate preterminal and terminal catches during base period
                            totBasePeriodCatch(PreTerm, fish) = totBasePeriodCatch(PreTerm, fish) + tempCat(PreTerm, fish)
                            totBasePeriodCatch(Term, fish) = totBasePeriodCatch(Term, fish) + tempCat(Term, fish)
                        End If
                    End If 'CeilingFisheryFlag(fish%) = 1
                End If 'timeStepFlag(fish%, timeStep%) > 0
            Next fish
            '..... do this after the ceiling base period years 1979-1984 (do not confuse with 1979-1982 base period)
        Else '..... ceilingSetup

            For fish = 1 To numFisheries
                If timeStepFlag(fish, timeStep) > 0 Then
                    If CeilingFisheryFlag(fish) = 1 Then
                        '..... Set ceiling levels other than the first
                        '..... in other words, get a working copy without the year subscript
                        'eq 24
                        ceil(location, fish) = CeilLevel(location, NextCeil, fish)
                    End If
                End If 'timeStepFlag(fish%, timeStep%) > 0
            Next fish

        End If

        '..... do this only when you have reached the last base period year
        '..... note this loop must follow the previous loop in order for TotBasePeriodCatch() to accum properly
        If yr = BaseCeilingEnd Then 'yr% = BaseCeilingEnd% only happens with calcTotalFishingMortalities, not with CeilingSetup
            For fish = 1 To numFisheries
                If timeStepFlag(fish, timeStep) > 0 Then
                    If CeilingFisheryFlag(fish) = 1 Then
                        'eq 22.....Compute avg Preterminal base period catch
                        'note the ceiling base period is the years before ceiling fisheries 1979-1984
                        'do not confuse with the 1979-1982 base period
                        averageBasePeriodCatch = totBasePeriodCatch(PreTerm, fish) / (BaseCeilingEnd - BaseCeilStart + 1)
                        For j = 1 To AddCeil + 1
                            '..... calculate catch that would have been obtained with 1979-1982 base period harvest rates
                            'eq 23
                            CeilLevel(PreTerm, j, fish) = basePeriodToCeilingScalar(j, fish) * averageBasePeriodCatch
                        Next j
                        'eq 22.....Compute avg Terminal base period Catch
                        averageBasePeriodCatch = totBasePeriodCatch(Term, fish) / (BaseCeilingEnd - BaseCeilStart + 1)
                        For j = 1 To AddCeil + 1
                            '..... calculate catch that would have been obtained with 1979-1982 base period harvest rates
                            'eq 23
                            CeilLevel(Term, j, fish) = basePeriodToCeilingScalar(j, fish) * averageBasePeriodCatch
                        Next j

                        'eq 24.....Set levels for first ceiling year (update working copy)
                        ceil(PreTerm, fish) = CeilLevel(PreTerm, 1, fish)
                        ceil(Term, fish) = CeilLevel(Term, 1, fish)
                    End If 'CeilingFisheryFlag(fish%) = 1
                End If 'timeStepFlag(fish%, timeStep%) > 0
            Next fish
        End If

    End Sub

    '--------------------------------------------------------------
    Sub CeilingTestIfModeledCatchMatchesObserved()
        '--------------------------------------------------------------
        '  Purpose:  Manages the calculation of ceilings for
        '            fisheries that are modeled as terminal for some
        '            stocks and preterminal for others.  The model
        '            calculates preterminal catches prior to terminal
        '            catches.  Consequently, a reduction in preterm-
        '            inal catches will increase terminal catches.
        '            Ceilings for these types of fisheries must be
        '            computed iteratively until the change becomes
        '            minimal (less than CutOffLevel).  Note that
        '            stock-specific terminal catches resulting from
        '            truncating spawning escapements at the goal are
        '            not included in computations of ceiling catches
        '            for terminal fisheries.
        '
        '  Arguments: Yr% = Simulation year.
        '
        '  Inputs:    Ceil()
        '             CeilingControlFlag%()
        '             CutOffLevel
        '             CeilingFisheryFlag()
        '             NumCeil
        '             isOcnTerm()
        '             PreTerm%
        '             RT()
        '             TempCat()
        '             Term%
        '
        '  Called By: CalcTerminalHarvest
        '
        '  Output:    isAnotherLoop
        '             RT_scalar!()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------

        Dim ratio, CurrentCatch, TotCeil As Double
        Dim fish As Integer

        '..... Set flag to indicate that another pass is not necessary
        isAnotherLoop = False
        For fish = 1 To numFisheries
            If timeStepFlag(fish, timeStep) > 0 Then
                If CeilingFisheryFlag(fish) = 1 Then
                    If isOcnTerm(fish) = True Then
                        '..... Compute total preterminal and terminal ceiling level
                        TotCeil = ceil(PreTerm, fish) + ceil(Term, fish)
                        '..... Compute total preterminal and terminal model catch level
                        CurrentCatch = tempCat(PreTerm, fish) + tempCat(Term, fish)
                        '..... Compute ratio between Model Catch and Ceiling eq 144
                        ratio = CurrentCatch / TotCeil
                        '..... Compute new ceiling adjustment factor  eq 145
                        If ratio Then
                            RT_scalar(fish) = RT_scalar(fish) / ratio
                        End If
                        '..... This trap collects situations where the CurrentCatch is less than the Ceiling
                        '..... The effect of the RT_scalar! Variable was to negate the effect of the ForceCeilFlg
                        '..... and to force the catch to the ceiling even if that effect was not wanted.

                        If RT(yr, PreTerm, fish) * RT_scalar(fish) > 1 Then
                            If CeilingControlFlag(yr, fish) = 0 Then
                                RT_scalar(fish) = RT_scalar(fish) * ratio
                                GoTo NxtJ
                            End If
                        End If
                        'eq 147
                        If System.Math.Abs(1 - ratio) > CutOffLevel Then
                            '..... Set flag to indicate that another pass is required for convergence
                            isAnotherLoop = True
                            '&&&&&& 6/99 GM DEBUG
                            If numLoops > 5 Then PrintLine(logFileID, TimeOfDay, "Fishery " & fish & " Test " & System.Math.Abs(1 - ratio))
                        End If
                    End If 'If CeilingFisheryFlag(fish%) = 1
                End If 'If isOcnTerm(fish%) = True
            End If 'timeStepFlag(fish%, timeStep%) > 0
NxtJ:
        Next fish

    End Sub

    '--------------------------------------------------------------
    Sub CeilingSetup(ByRef location As Integer)
        '--------------------------------------------------------------
        '  Purpose:  call subroutine to get a working copy of observed ceiling catch
        '            or call subroutine to calculate ceiling catch
        '
        '  Arguments: Yr%  = Simulation year
        '             location = Terminal/PreTerminal
        '
        '  Inputs:    AddCeil%
        '             isAnotherLoop
        '             CeilYr%()
        '             CeilingEnd
        '             StartCeil
        '
        '  Called By: calcPreterminalHarvest
        '             CalcTerminalHarvest
        '
        '  Output:    None.
        '
        '  Externals: CeilReduc
        '             ceilingLevel
        '             FActive
        '
        '
        '--------------------------------------------------------------

        Dim i, NextCeil As Integer

        If isAnotherLoop = False Then
            '..... Ceiling convergence criteria satisfied
            If AddCeil > 0 Then
                For i = 1 To AddCeil
                    NextCeil = i + 1
                    If yr = CeilYR(NextCeil) Then
                        '..... Set next ceiling level
                        Call CeilingLevel(NextCeil, location)
                        Exit For
                    End If
                Next i
            End If
        End If
        '..... Check if year is within ceiling management period
        If FActive(startCeil, CeilingEnd) Then
            '..... Adjust catches to ceilings
            Call CeilingHarvest(location)
        End If
    End Sub


    '--------------------------------------------------------------
    Function FActive(ByRef s As Integer, ByRef E As Integer) As Boolean
        '--------------------------------------------------------------
        '  Purpose:  Determines if option is active (i.e. start <=
        '            simulation year <= end)
        '
        '  Arguments: S% = Start year for option
        '             E% = End year for option
        '             yr% = Current simulation year
        '
        '  Inputs:    None.
        '
        '  Called By: calcPreterminalHarvest
        '             CalcTerminalHarvest
        '             calcEscapement
        '             ceilingSetup
        '
        '  Output:    TRUE if S <= yr5 <=E
        '             FALSE% otherwise
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        FActive = False
        If yr >= s And yr <= E Then FActive = True

    End Function

    '--------------------------------------------------------------
    Function FEnhSpawners(ByRef SmoltProductionGoal As Double, ByRef SmoltSurvRt As Double, ByRef ProductivityPerAdult As Double) As Double
        '--------------------------------------------------------------
        '  Purpose:  Computes number of adults required to meet
        '            broodstock requirements for enhancement.
        '
        '  Arguments: SmoltProductionGoal      = Number of smolts to be produced
        '             SmoltSurvRt = Survival factor from smolt
        '                           release to age 1 fish
        '             AValue      = Production efficiency per adult
        '                           taken for broodstock
        '
        '  Inputs:    None.
        '
        '  Called By: calcEscapement
        '
        '  Output:    Number of adults required to produce desired
        '             smolt production
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------


        'eq 190
        FEnhSpawners = (SmoltProductionGoal * SmoltSurvRt) / System.Math.Exp(ProductivityPerAdult)

    End Function


    '--------------------------------------------------------------
    Function FMaxOne(ByRef Express As Single) As Single
        '--------------------------------------------------------------
        '  Purpose:  Prevents value from exceeding 1.
        '
        '  Arguments: Express! = expresssion to be evaluated
        '
        '  Inputs:    None.
        '
        '
        '  Output:    1 if Express!>1
        '             Express! Otherwise
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        If Express > 1 Then
            FMaxOne = 1
        Else
            FMaxOne = Express
        End If

    End Function

    Function FMinZed(ByRef Express As Double) As Double
        '--------------------------------------------------------------
        '  Purpose:  Prevents value from being less than zero.
        '
        '  Arguments: Express! = expresssion to be evaluated
        '
        '  Inputs:    None.
        '
        '
        '  Output:    0 if Express! < 0
        '             Express! Otherwise
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        If Express < 0 Then
            FMinZed = 0
        Else
            FMinZed = Express
        End If

    End Function

    '--------------------------------------------------------------
    Function FRecruitRate(ByRef s As Integer) As Double
        '--------------------------------------------------------------
        '  Purpose:  Computes number of recruits for stock, based
        '            on standard survivals and stock specific
        '            maturity rates.
        '
        '  Arguments: S% = Stock number
        '
        '  Inputs:    None.
        '
        '  Called By: CalcEV
        '             readMATFRL_DATfile
        '
        '  Output:    Number of recruits at age 1
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------

        Dim age As Integer
        Dim spawnerEQ As Double

        '..... Compute factor to convert calculated "spawner equivalents" to actual numbers
        '..... of recruited Fish% at Age% 1, i.e. age 1 to adult survival rate
        'if there is new maturity data in the mat.dat file (which will replace those in the .stk file)
        'then recalculate the age1ToAdultSurvivalRate here

        'eq 181 version 2 of 2, this one using maturity data in the mat*.dat file
        FRecruitRate = 0
        spawnerEQ = survivalRate(1)
        For age = 2 To NumAges
            spawnerEQ = spawnerEQ * survivalRate(age)
            'eq 137
            FRecruitRate = FRecruitRate + spawnerEQ * maturationRate(age, s)
            spawnerEQ = spawnerEQ * (1 - maturationRate(age, s))
        Next age

    End Function

    '--------------------------------------------------------------
    Function TotalAdultReturn(ByRef escapement As Double, ByRef stk As Integer, ByRef yr As Integer) As Double
        '--------------------------------------------------------------
        '  Purpose:  Computes the production, expressed in terms of
        '            adult equivalents using a truncated form of a
        '            Ricker stock-recruitment function.  Prior to
        '            1988, the model did not allow spawning escmts to
        '            exceed escapement goals.  However, not all
        '            stocks have terminal fisheries capable of
        '            controlling escapements.  While escapements are
        '            allowed at any level, production is truncated at
        '            the maximum.  This formulation thus prevents
        '            production from decreasing at high levels of
        '            spawning escapements.  An optimistic view may
        '            result, however, production response at high
        '            escapement levels is unknown for nearly all
        '            stocks.
        '
        '  Arguments: escapement  = Age 3+ spawning escapement
        '             Stk% = Stock index number
        '
        '  Inputs:    IDL!()
        '             optimumSpawners()
        '             RickerA()
        '
        '  Called By: calcEscapement
        '
        '  Output:    Adult equivalent production
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------

        Dim a, b As Double
        Dim MaxEsc As Single

        a = RickerA(stk)
        '&&&&& gm 6/99 enable use of exact Ricker parameters if known
        If CompNewSRParams(stk) = 1 Then
            '..... Use Old S-R Parameters and truncate production at max
            'eq 217
            b = optimumSpawners(stk) / (0.5 - 0.07 * a)
            TotalAdultReturn = escapement * System.Math.Exp(a * (1 - escapement / b))
            '..... If escapement exceeds level producing maximum recruitment,
            '..... keep recruits at maximum.  cf Ricker 1975, p. 347, eq. 10
            'eq 218
            MaxEsc = b / (a * IDL(stk, yr))
            If escapement > MaxEsc Then
                'eq 219
                TotalAdultReturn = b * System.Math.Exp(a - 1) / a
            End If
        Else
            '..... Use Biologically-based parameters, do not truncate production
            b = RickerB(stk)
            'eq 179
            TotalAdultReturn = escapement * System.Math.Exp(a * (1 - escapement / b))
        End If

    End Function

    '--------------------------------------------------------------
    Function FTerminal(ByRef stk As Integer, ByRef fish As Integer, ByRef age As Integer) As Integer
        '--------------------------------------------------------------
        '  Purpose:  Determines if a fishery is terminal or preterm
        '
        '  Arguments: Stk% = Stock index number
        '             Fish%= Fishery index number
        '             Age% = Age of fish
        '
        '  Inputs:    oceanNetFlag%()
        '             PreTerm%
        '             Term%
        '             TermNetAge
        '             terminalFlag%()
        '
        '  Called By: CalcencounterRatio
        '             CalcShakersAndCNR
        '             CalcTerminalHarvest
        '             calcEscapement
        '             ceilingOcnTerm
        '
        '  Output:    TRUE if terminal fishery or
        '                  if ocean net fishery and age >= TermNetAge
        '             FALSE otherwise
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        'terminalFlag%, oceanNetFlag%, TermNetAge are defined in BSE file
        'fishery is preterminal by default
        'if terminalFlag% = 1 then fishery is terminal for ages 2-maxAge
        'if oceanNetFlag%(fish%) = 1 And age% >= TermNetAge (age 4) then fishery is terminal for ages 4-maxAge

        FTerminal = PreTerm
        If (terminalFlag(stk, fish)) Or (oceanNetFlag(fish) And age >= TermNetAge) Then FTerminal = Term

    End Function



    '--------------------------------------------------------------
    Sub GSHatchCalib()
        '--------------------------------------------------------------
        '  Purpose:  Computes estimated catches for Lower Georgia
        '            Strait Hatchery Stock (GSH) to GS troll and
        '            sport fisheries for EV calculations.
        '
        '  Inputs:    legalCatch()
        '             NumAges
        '             PreTerm%
        '             Term%
        '
        '
        '  Output:    GSH_Cat_EV()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        'this subroutine was never called

        Dim age As Integer
        Dim tmp6, tmp24 As Double

        For age = 2 To NumAges
            tmp6 = legalCatch(PreTerm, 6, 9, age) + legalCatch(Term, 6, 9, age)
            tmp24 = legalCatch(PreTerm, 24, 9, age) + legalCatch(Term, 24, 9, age)
            GSH_Cat_EV(yr, age) = GSH_Cat_EV(yr, age) + tmp6 + tmp24
        Next age

    End Sub


    '--------------------------------------------------------------
    Sub InitializeArrays()
        '--------------------------------------------------------------
        '  Purpose:  Resets arrays for each iteration of the
        '            calibration process.
        '
        '  Arguments: None.
        '
        '  Inputs:    TotalEscape_EV()
        '             isCalibration
        '             Ceil()
        '             FirstCohort()
        '             FirstHRCat()
        '             HarvestRateReport$
        '             NumAges
        '             NumYears
        '             numStocks
        '             NumYears
        '             OcnAEHR()
        '             SaveFile%()
        '             GSH_Cat_EV()
        '             StkMort()
        '             StkProp()
        '             OceanEscape_EV()
        '             TotalAdultEscape_SR()
        '             TotAEHR()
        '             TotCatFish_()
        '             TotCNRFish()
        '             TotShakFish()
        '             terminalRunReport$
        '             TrueTermRun()
        '
        '
        '  Output:    TotalEscape_EV()
        '             Ceil()
        '             Cohort_()
        '             harvestRate()
        '             OcnAEHR()
        '             RT()
        '             GSH_Cat_EV()
        '             StkMort()
        '             StkProp()
        '             OceanEscape_EV()
        '             TotalAdultEscape_SR()
        '             TotAEHR()
        '             TotCatFish_()
        '             TotCNRFish()
        '             TotShakFish()
        '             TrueTermRun()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------

        Dim stk, j, f_, age, fish, year_ As Integer

        If isCalibration = True Then 'stage 1 or 2 calibration
            '..... Calibration run, reset arrays used to calculate EV to zero
            ReDim totalEscape_EV(NumYears, numStocks, NumAges)
            '******** 2/96 Eliminate unnecessary array ********
            If EVType(9) = 3 Then
                ReDim GSH_Cat_EV(NumYears, NumAges)
            Else
                Erase GSH_Cat_EV
            End If
            '**************************************************
            ReDim OceanEscape_EV(NumYears, numStocks, NumAges)
        End If

        '..... Reset ceiling catches
        ReDim ceil(1, numFisheries)
        ReDim CeilLevel(1, AddCeil + 1, numFisheries)

        '..... Check if Ocean and Total Exploitation Rate data are  requested
        If (SaveFile(3) + SaveFile(4)) > 0 Or HarvestRateReport <> "N" Then
            ReDim ocnAEHR(NumYears, numStocks)
            ReDim TotAEHR(NumYears, numStocks)
        Else
            Erase ocnAEHR
            Erase TotAEHR
        End If

        '..... Check if stats on individual stock mortality are requested
        If Stm > 0 Then
            ReDim stkMort(NumYears, numFisheries, numStocks)
        Else
            Erase stkMort
        End If

        '..... check if Stock Composition stats are requested
        If Stp = "Y" Then
            ReDim stkProp(NumYears, numStocks, NumStkCatch)
        Else
            Erase stkProp
        End If

        '..... Check if terminal run stats are requested
        If SaveFile(1) > 0 Or terminalRunReport = "Y" Then
            ReDim trueTermRun(NumYears, numStocks)
        Else
            Erase trueTermRun
        End If

        ReDim totalAdultEscape_SR(NumYears, numStocks)
        ReDim totEsc_allAges(numStocks)
        ReDim totBasePeriodCatch(1, numFisheries)

        ReDim TotCatFish_(NumYears, numFisheries)
        ReDim TotCNRFish(1, NumYears, NumCNRFish)
        ReDim totShakFish(NumYears, numFisheries)

        '..... Initialize arrays cohort sizes and Harv Rates ......
        ReDim cohort_(NumAges, numStocks)
        ReDim harvestRate(numFisheries, NumAges, numStocks)

        For stk = 1 To numStocks
            '***** 2/21/96 Modified to allow EVs for broods prior first model year *******
            '************ 2/96 ***** Get initial cohort sizes from random access file ****
            For age = 1 To NumAges
                cohort_(age, stk) = initialAbundance(stk, age) * EV(-age + NumAges, stk)
            Next age
            Spring(stk) = cohort_(1, stk)
            For fish = 1 To numFisheries
                For age = 2 To NumAges
                    harvestRate(fish, age, stk) = baseHarvestRate(fish, age, stk)
                Next age
            Next fish
        Next stk

        '******2/96 ***** Changed loop variables from fish% to f_% because of conflict with yr%
        For year_ = 0 To NumYears
            For f_ = 1 To numFisheries
                For j = 0 To 1
                    RT(year_, j, f_) = 1
                Next j
            Next f_
        Next year_



    End Sub

    '--------------------------------------------------------------
    Sub ResetPNV()
        '--------------------------------------------------------------
        '  Purpose:  Resets proportion non-vulnerable array
        '
        '  Arguments: Yr% = Simulation year
        '
        '  Inputs:    NumAges
        '             numFisheries
        '             PNVfisheryIndex()
        '             YrPNV()
        '
        '
        '  Output:    RT_scalar!()
        '             PNV()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        'note conditional statement for timesteps not required

        Dim FishIndex, age, fish As Integer

        ReDim RT_scalar(numFisheries) 'RT_scalar! was not dimensioned in QB

        For fish = 1 To numFisheries
            RT_scalar(fish) = 1 'starting value eq 143
            '..... Obtain index of fishery with PNV change
            FishIndex = PNVfisheryIndex(fish)
            If FishIndex > 0 Then
                '..... Reset PNV's by age
                For age = 2 To NumAges
                    If yr <= PNVEnd Then
                        PNV(fish, age) = YrPNV(yr, FishIndex, age)
                    Else
                        'If year is later than last year in PNV file, PNV = 0 if no action is taken
                        'therefore use the last year's data (not data in BSE file)
                        PNV(fish, age) = YrPNV(PNVEnd, FishIndex, age)
                    End If
                Next age
            End If
        Next fish

    End Sub

    '--------------------------------------------------------------
    Sub SaveHarvestRatesAfterCeilingFishery()
        '--------------------------------------------------------------
        '  Purpose:  Save harvest rates after they have been adjusted by RT and RT_scalar
        '            to correctly calculate ceiling fishery
        '
        '  Arguments: None.
        '
        '  Inputs:    CeilingEnd
        '             CeilingFisheryFlag()
        '             NumAges
        '             NumCeil
        '             numStocks
        '             oceanNetFlag%()
        '             isOcnTerm()
        '             PreTerm%
        '             RT()
        '             Term%
        '             TermNetAge
        '             terminalFlag%()
        '
        '  Called By: calcEscapement
        '
        '  Output:    harvestRate()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------

        Dim age, fish, stk As Integer

        For fish = 1 To numFisheries
            If timeStepFlag(fish, timeStep) > 0 Then
                If CeilingFisheryFlag(fish) = 1 Then
                    If isOcnTerm(fish) = True Then
                        '..... Preterminal fishery
                        '..... Change harvest rates
                        For stk = 1 To numStocks
                            For age = 2 To NumAges
                                '..... save harvest rates adjusted to calculate the correct ceiling catch
                                'eq 149
                                harvestRate(fish, age, stk) = harvestRate(fish, age, stk) * RT(CeilingEnd, PreTerm, fish)
                            Next age
                        Next stk
                    Else
                        '..... Terminal fishery
                        For stk = 1 To numStocks
                            If terminalFlag(stk, fish) Then
                                '..... Terminal fishery for this stock
                                For age = 2 To NumAges
                                    '..... save harvest rates adjusted to calculate the correct ceiling catch
                                    'eq 149
                                    harvestRate(fish, age, stk) = harvestRate(fish, age, stk) * RT(CeilingEnd, Term, fish)
                                Next age
                            Else
                                '..... Preterminal fishery  for this stock
                                For age = 2 To NumAges
                                    '..... Check if an ocean net fishery
                                    If oceanNetFlag(fish) Then
                                        '..... Check if fishery should be considered terminal
                                        '..... save harvest rates adjusted to calculate the correct ceiling catch eq 149
                                        If age >= TermNetAge Then
                                            harvestRate(fish, age, stk) = harvestRate(fish, age, stk) * RT(CeilingEnd, Term, fish)
                                        Else
                                            harvestRate(fish, age, stk) = harvestRate(fish, age, stk) * RT(CeilingEnd, PreTerm, fish)
                                        End If
                                    Else
                                        harvestRate(fish, age, stk) = harvestRate(fish, age, stk) * RT(CeilingEnd, PreTerm, fish)
                                    End If
                                Next age
                            End If
                        Next stk
                    End If
                End If
            End If 'timeStepFlag(fish%, timeStep%) > 0
        Next fish

    End Sub

    '--------------------------------------------------------------
    Sub ReadMATFRL_DATfile()
        '--------------------------------------------------------------
        '  Purpose:  Gets stock-specific maturity rate and adult
        '            equivalence factors annually from a sequential
        '            ASCII data file.
        '
        '  Arguments: Yr% = Simulation year
        '
        '  Inputs:    FileNum%
        '             nameMATFRL_DATfile$
        '             NumMatSched%
        '
        '
        '  Output:    AdltEqv()
        '             maturationRate()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------

        Dim stk, i, age, Found, j, year_ As Integer
        Dim stock_name As String

        Select Case Iter
            Case 1
                'Do all error checks for first year only
                Input(MATFRL_DATFileID, year_)
                If (year_ - startYear) <> yr Then
                    MsgBox("Fatal Error in " & nameMATFRL_DATFile & ":  year " & year_ & " is not annual time step " & yr & ".  Program will stop")
                    End
                End If
                For i = 1 To NumMATSched
                    Input(MATFRL_DATFileID, stock_name)
                    stock_name = Left(UCase(stock_name), 3)
                    Found = 0
                    For stk = 1 To numStocks
                        If stock_shortName(stk) = stock_name Then
                            Found = 1
                            If StkMat(stk) <> 2 Then
                                MsgBox("Fatal Error in " & nameMATFRL_DATFile & ": Mismatch for stock " & stock_name & " between EVData file and Maturity File.  Stop will stop.")
                                End
                            End If
                            For age = 2 To 4
                                Input(MATFRL_DATFileID, maturationRate(age, stk))
                            Next age
                            For age = 2 To 4
                                Input(MATFRL_DATFileID, AdltEqv(age, stk))
                            Next age
                            'get % age 1 to adult survival rate
                            Age1toAdultSurvivalRate(stk) = FRecruitRate(stk)
                            '******** 2/9/94 ******************
                            For j = 1 To numEnhancedStocks
                                '..... Check if this is an enhanced stock
                                If Enh_stockIndex(j) = stk Then
                                    '..... Compute new enhancement efficiency
                                    EnhEff_(j) = System.Math.Exp(AValue(j)) * Age1toAdultSurvivalRate(stk) / System.Math.Exp(RickerA(stk))
                                    Exit For
                                End If
                            Next j
                            '**********************************
                            Exit For 'stk% = 1 To NumStocks
                        End If
                    Next stk
                    If Found = 0 Then
                        MsgBox("Fatal Error in " & nameMATFRL_DATFile & ": unable to identify " & stock_name & ". Program will stop.")
                        End
                    End If
                Next i
                Exit Sub



            Case Else
                '..... No need to read data file.  Just from one array to next
                Input(MATFRL_DATFileID, year_)
                For i = 1 To NumMATSched
                    Input(MATFRL_DATFileID, stock_name)
                    stock_name = Left(UCase(stock_name), 3)
                    For stk = 1 To numStocks
                        If stock_shortName(stk) = stock_name Then
                            For age = 2 To 4
                                Input(MATFRL_DATFileID, maturationRate(age, stk))
                            Next age
                            For age = 2 To 4
                                Input(MATFRL_DATFileID, AdltEqv(age, stk))
                            Next age
                            'get % age 1 to adult survival rate
                            Age1toAdultSurvivalRate(stk) = FRecruitRate(stk)
                            Exit For
                        End If
                    Next stk
                Next i
        End Select

    End Sub

    '--------------------------------------------------------------
    Sub Write_EV_File()
        '--------------------------------------------------------------
        '  Purpose:  Writes file of EV factors for calibration run
        '            and revised .STK file (new extension .STx)
        '            containing revised maturation rates.
        '
        '  Arguments: None.
        '
        '  Inputs:    EVType%()
        '             MaxYrs
        '             numStocks
        '             PointBack%()
        '             SaveFilePrefix$
        '             Stk%
        '             EV()
        '
        '
        '  Output:    File of computed EV factors for calibration.
        '
        '  Externals: None.
        '
        '--------------------------------------------------------------

        Dim i, age, EVOFileID, IStk As Integer
        Dim NewSTKFile As String
        Dim NewSTKFileID As Integer
        Dim STKFileID, stk, year_ As Integer
        Dim temp As String

        EVOFileID = FreeFile()
        FileOpen(EVOFileID, nameEVOFile, OpenMode.Output)
        '***** 2/21/96 Modified to allow broods prior to first model year ************
        PrintLine(EVOFileID, startYear - NumAges)
        '*****************************************************************************
        PrintLine(EVOFileID, MaxYrs + startYear)
        For stk = 1 To numStocks
            Print(EVOFileID, " " & VB6.Format(stk, "00") & " ")
            If EVType(PointBack(stk)) = 4 Then
                IStk = PointBack(stk)
            Else
                IStk = stk
            End If
            '***** 2/21/96 Modified to allow broods prior to first model year ************
            For year_ = -NumAges To MaxYrs
                Print(EVOFileID, VB6.Format(EV(year_ + NumAges, IStk), "0.########0E+00") & "  ")
            Next year_
            PrintLine(EVOFileID)
        Next stk
        FileClose(EVOFileID)
        If AdjMatRate = 0 Then Exit Sub

        '..... Rewrite .STK file as .STx file with revised maturation rates
        STKFileID = FreeFile()
        FileOpen(STKFileID, nameSTKFile, OpenMode.Input)
        i = InStr(nameSTKFile, ".") - 1
        If i < 1 Then
            i = Len(nameSTKFile)
        End If
        Select Case AdjMatRate
            Case 1
                '..... Age 3 adjustment
                NewSTKFile = saveFilePrefix & "NEW.ST3"
            Case Is > 1
                '..... Age 3/4 adjustment
                NewSTKFile = saveFilePrefix & "NEW.ST4"
            Case Else
        End Select
        FileOpen(NewSTKFileID, NewSTKFile, OpenMode.Output)
        For stk = 1 To numStocks
            '..... Copy stock name and cohort sizes
            For i = 1 To 2
                temp = LineInput(STKFileID)
                PrintLine(NewSTKFileID, temp)
            Next i
            '..... Discard old maturation schedule
            temp = LineInput(STKFileID)
            If StkMat(stk) = 1 Or StkMat(stk) > 2 Then
                For age = 2 To NumAges
                    '..... Replace with new maturation schedule
                    Print(NewSTKFileID, VB6.Format(maturationRate(age, stk), "0.#######0E+00") & " ")
                Next age
                PrintLine(NewSTKFileID)

                '..... Discard AE factors in .STK file
                temp = LineInput(STKFileID)
                '..... Replace with new adult equivalence factors
                For age = 2 To NumAges
                    '..... Replace with new maturation schedule
                    Print(NewSTKFileID, VB6.Format(AdltEqv(age, stk), "0.#######0E+00") & " ")
                Next age
                PrintLine(NewSTKFileID)

            Else
                '..... Copy maturation rates and adult equivalence factors
                PrintLine(NewSTKFileID, temp)
                temp = LineInput(STKFileID)
                PrintLine(NewSTKFileID, temp)
            End If
            '..... Copy Fishery HRs
            For i = 1 To numFisheries
                temp = LineInput(STKFileID)
                PrintLine(NewSTKFileID, temp)
            Next i
        Next stk
        FileClose(STKFileID, NewSTKFileID)
    End Sub




    '--------------------------------------------------------------
    Sub AnnualLoop()
        '--------------------------------------------------------------
        '  Purpose:  Controls the main simulation time loop.
        '
        '  Arguments: None
        '
        '  Inputs:    AdjMatRate%
        '             isAnotherLoop
        '             BaseCeilingEnd%
        '             BaseCeilStart%
        '             BiggestProb
        '             BigProb
        '             isCalibration
        '             CNREnd%
        '             CNRStart%
        '             ConvFlag%
        '             CeilingEnd%
        '             EnhEnd%
        '             Iter
        '             LastRepeat
        '             MatIter
        '             NumStkCatch
        '             PNVEnd%
        '             PNVStart%
        '             ProbBY
        '             ProbStock$
        '             StartCeil%
        '             EnhStart%
        '
        '
        '  Externals: CalcHarvestStockComposition
        '             calcPreterminalHarvest
        '             CalcTerminalHarvest
        '             CalcTotalFishingMortalities
        '             calcEscapement
        '             resetPNV
        '             readMATFRL_DATfile
        '             Message
        '
        '
        '--------------------------------------------------------------

        Call InitializeArrays()
        '..... Check if annual maturity data are to be entered
        '..... Open maturity schedule file
        '..... because time loop will call Sub readMATFRL_DATfile(yr%) one year at a time
        If isAnnualMaturityData = True Then
            MATFRL_DATFileID = FreeFile()
            FileOpen(MATFRL_DATFileID, nameMATFRL_DATFile, OpenMode.Input)
        End If

        '..... Main Time Loop
        For yr = 0 To NumYears
            PrintLine(logFileID, TimeOfDay, "Year = " & startYear + yr)
            FrmMain.sbr_year.Text = "year = " & startYear + yr

            If ConvFlagMR = 0 Then
                PrintLine(logFileID, TimeOfDay, "Second pass after maturation rate adjustment")
            End If
            If isCalibration = True And ConvFlag = 0 Or ConvFlagMR = 0 Then 'stage 1 or 2 calibration
                If biggestProb > 0 Then
                    PrintLine(logFileID, TimeOfDay, "WORST PROB: BY " & probBY & " " & ProbStock & " - NEW EV: " & VB6.Format(BigProb, "00.##0") & " CT: " & VB6.Format(biggestProb, "0.##0E+000") & " Iter: " & Iter - 1)
                    If yr = 0 Then
                        PrintLine(HIS_RECfileID, "WORST PROB: BY " & probBY & " " & ProbStock & " - NEW EV: " & VB6.Format(BigProb, "00.##0") & " CT: " & VB6.Format(biggestProb, "0.##0E+000") & " Iter: " & Iter - 1)
                    End If
                End If
            End If
            If isCalibration = True And lastRepeat = 1 Then 'stage 1 or 2 calibration
                PrintLine(logFileID, TimeOfDay, "Calibration completed.  Final pass to generate stats.")
            End If

            '..... Reset Prop Non-Vulnerable
            Call ResetPNV()
            '..... Reset fishery policies

            '..... Check if annual maturity data are to be used
            If isAnnualMaturityData = True Then
                '..... Read in maturity rates and adult equivalence factors
                Call ReadMATFRL_DATfile()
            End If

            '..... Display active options
            If FActive(BaseCeilStart, BaseCeilingEnd) Then PrintLine(logFileID, TimeOfDay, "      Base Ceiling Period")
            If FActive(startCeil, CeilingEnd) Then PrintLine(logFileID, TimeOfDay, "      Ceilings")
            If FActive(PNVStart, PNVEnd) Then PrintLine(logFileID, TimeOfDay, "      PNV Change")
            If FActive(CNRStart, CNREnd) Then PrintLine(logFileID, TimeOfDay, "      CNR Fisheries")
            If FActive(EnhStart, EnhEnd) Then PrintLine(logFileID, TimeOfDay, "      Enhancement Change")

            '..... Display iteration number for calibration run
            If isCalibration = True Then 'stage 1 or 2 calibration
                PrintLine(logFileID, TimeOfDay, "      Iteration Number:  ", Iter)
                FrmMain.sbr_iter.Text = " Iteration = " & Iter
                If AdjMatRate Then
                    PrintLine(logFileID, TimeOfDay, "      Maturation rate adjustment")
                End If

                '..... Revert to every year EV calculations after CalibCycle% cycles
                If Iter = CalibSwitch Then CalibMethod = 0
            End If

            'Print #97, "--------------------start new year"; yr%

            Call ApplySurvivalRates()

            For timeStep = 1 To numTimeStepsSimulated

                'Print #97, "--------------------start new time step"; timeStep%

                PrintLine(logFileID, TimeOfDay, "time step = " & timeStep)
                FrmMain.sbr_timeStep.Text = "time step = " & timeStep

                numLoops = 1
                isAnotherLoop = False 'must be false for Sub CeilingSetUp to work properly

                PrintLine(logFileID, TimeOfDay, "ocean")
                FrmMain.sbr_OceanTerminal.Text = "ocean"
                Call CalcPreterminalHarvest() 'legal catch, ceiling catch, shakers, CNR

                PrintLine(logFileID, TimeOfDay, "terminal")
                FrmMain.sbr_OceanTerminal.Text = "terminal"
                Call ApplyMaturityRates()
                Call CalcTerminalHarvest() 'legal catch, ceiling catch

                'because pre-terminal catches are computed prior to
                'the calculation of mature run sizes, an iterative
                'procedure must be employed to find a preterminal ceiling
                'harvest rate (Do While isAnotherLoop = True) such that the combined
                'preterminal and terminal ceiling catches = observed catch

                'isAnotherLoop = true in CeilingTestIfModeledCatchMatchesObserved if Abs(1 - ratio) > CutOffLevel
                'isAnotherLoop = false if NumLoops > 10
                Do While isAnotherLoop = True
                    If numLoops < 10 Then
                        PrintLine(logFileID, TimeOfDay, "Number of loops: " & numLoops)
                        numLoops = numLoops + 1
                        PrintLine(logFileID, TimeOfDay, "ocean")
                        FrmMain.sbr_OceanTerminal.Text = "ocean"
                        Call CalcPreterminalHarvest() 'legal catch, ceiling catch, shakers, CNR

                        PrintLine(logFileID, TimeOfDay, "terminal")
                        FrmMain.sbr_OceanTerminal.Text = "terminal"
                        Call ApplyMaturityRates()
                        Call CalcTerminalHarvest() 'legal catch, ceiling catch
                        'NumLoops% = NumLoops% + 1
                    Else
                        isAnotherLoop = False
                    End If
                Loop
                Call CalcTotalFishingMortalities()

                PrintLine(logFileID, TimeOfDay, "escapement")
                FrmMain.sbr_OceanTerminal.Text = "escapement"

                Call CalcTerminalRunAndEscapement()

                If optionCCC = True Then Call write_CCCfile()

                If optionTotalMort = True Then Call write_totalMortality()

                If optionISBM = True And isCalibration = False Then 'projection fun
                    'If yr <= NumYears - 2 Then Call write_ISBM()
                    Call write_ISBM()
                End If

            Next timeStep

            'AdvanceCohortAge must follow CalcTerminalRunAndEscapement
            'and precede CalcAge1Production
            Call AdvanceCohortAge()

            Call CalcAge1Production()

        Next yr

        '..... Close maturation data file.  Must reopen for each iteration
        If isAnnualMaturityData = True Then FileClose(MATFRL_DATFileID)

    End Sub

    '--------------------------------------------------------------
    Sub GSHCheck(ByRef stk As Integer)
        '--------------------------------------------------------------
        'purpose:  Georgia Straits Hatchery
        '--------------------------------------------------------------

        Dim GSHNatSpawn As Double
        Dim s As Integer

        '..... Procedure to handle Production From Natural spawning of hatchery surplus

        If stock_shortName(stk) = "GSH" Then
            '..... Determine level of potential natural escapement
            GSHNatSpawn = FMinZed(totalAdultEscape_SR(yr, stk) - OPT - enhSpawn)
            If GSHNatSpawn > 5000 Then GSHNatSpawn = 5000
            If GSHNatSpawn Then
                '..... Find production parameters for lower Georgia Strait Natural Stock
                For s = 1 To numStocks
                    If stock_shortName(s) = "GST" Then
                        '..... Add production of age 1 recruits from spawning to GSH
                        'total return from Ricker curve/% age 1 to adult survival rate
                        'eq 180 3 of 3 instances
                        age1Fish = age1Fish + TotalAdultReturn(GSHNatSpawn, s, yr) / Age1toAdultSurvivalRate(s)
                        Exit For
                    End If
                Next s
            End If
        Else
            '..... Assume no contriibutions from natural spawning of surplus hatchery fish
        End If

    End Sub

    '--------------------------------------------------------------
    Sub CalcAge1Production()
        '--------------------------------------------------------------
        'purpose:  calculate age 1 production from escapement
        '--------------------------------------------------------------

        Dim age As Integer
        Dim EnhEff, EffEscape As Double
        Dim j As Integer
        Dim MaxBrood As Single
        Dim OldAdultEsc As Double
        Dim stk As Integer
        Dim tmp As Double

        For stk = 1 To numStocks
            'If stk% = 5 And yr% = 0 Then Print #97, TotalAdultEscape_SR(yr,stk%); "escape in calcAge1Prod"
            'eq 179
            OPT = optimumSpawners(stk)
            'eq 172
            escapement = totalAdultEscape_SR(yr, stk)
            enhSpawn = 0
            age1Fish = 0
            '..... Determine if this is a hatchery or natural stock
            Select Case HatchFlg(stk)
                '**************1/96 ****** SPRING STOCK PROVISION **********
            Case 1, 2 '..... fall and spring Hatchery Stock ...........................
                    '***********************************************************
                    If numEnhancedStocks > 0 And FActive(EnhStart, EnhEnd) Then
                        '..... Enhancement Changes Specified .........
                        For j = 1 To numEnhancedStocks
                            '..... Check if this is a stock that has a changed enhancement schedule
                            If Enh_stockIndex(j) = stk Then
                                '..... Determine number of smolts required to meet production goal
                                If SmoltProd(j, yr) Then
                                    '..... Determine number of spawners required to meet smolt production goal
                                    enhSpawn = FEnhSpawners(SmoltProd(j, yr), SmoltSurvRt(j), AValue(j))
                                    Select Case enhSpawn
                                        Case Is > 0
                                            '..... Production increase is indicated.  Determine if escapement
                                            '..... is adequate to meet production goal.
                                            If totalAdultEscape_SR(yr, stk) >= OPT Then
                                                '..... At least the base production can be achieved.
                                                'eq 178 2 of 2 instances
                                                age1Fish = System.Math.Exp(RickerA(stk)) * OPT
                                                '..... Check if total escapement is adequate for production increase
                                                If totalAdultEscape_SR(yr, stk) >= OPT + enhSpawn Then
                                                    '..... Add production increase using production efficiency
                                                    age1Fish = age1Fish + System.Math.Exp(AValue(j)) * enhSpawn
                                                    '..... Special code for contributions from spawning surplus of GSH stock
                                                    Call GSHCheck(stk)
                                                Else
                                                    '..... Full production increase is not possible.
                                                    age1Fish = age1Fish + System.Math.Exp(AValue(j)) * (escapement - OPT)
                                                End If
                                            Else
                                                '..... Escapement insufficuient to meet even base production level
                                                'eq 177 1 of 2 instances
                                                age1Fish = System.Math.Exp(RickerA(stk)) * escapement
                                            End If
                                        Case Else
                                            '..... Production decrease is indicated.
                                            '..... Remember that EnhSpawn is negative.
                                            If totalAdultEscape_SR(yr, stk) < (OPT + enhSpawn) Then
                                                age1Fish = System.Math.Exp(RickerA(stk)) * totalAdultEscape_SR(yr, stk)
                                            Else
                                                age1Fish = System.Math.Exp(RickerA(stk)) * (OPT + enhSpawn)
                                            End If
                                    End Select 'enhSpawn
                                    GoTo NextComp
                                End If 'SmoltProd(j%, yr%)
                                GoTo HatchProd
                            End If 'Enh_stockIndex(j%) = stk%
                        Next j '= 1 To numEnhancedStocks
                        GoTo HatchProd
                    End If 'numEnhancedStocks > 0 And FActive(EnhStart%, EnhEnd%, yr%)

HatchProd:
                    '..... No changes in enhancement have been specified or this
                    '..... is a hatchery stock without changes in enhancement
                    '..... Compute production.
                    Select Case (escapement - OPT)
                        Case Is >= 0
                            '..... Total adult escapement is adequate to meet production capacity
                            'eq 178 1 of 2 instances
                            age1Fish = System.Math.Exp(RickerA(stk)) * OPT
                            '..... Surplus escapement potentially available to contribute
                            '..... to natural production.  Check if this is the GSH stock.
                            Call GSHCheck(stk)
                        Case Else
                            '..... Total adult escapement is not adequate to meet production capacity.
                            'eq 177 2 of 2 instances
                            age1Fish = System.Math.Exp(RickerA(stk)) * escapement
                    End Select
                Case Else '..... 0 = fall, 3 = spring wild/Natural stock
                    '..... Compute Age 1 recruitment .........................
                    '..... Check if supplemental production has been specified
                    If numEnhancedStocks > 0 And FActive(EnhStart, EnhEnd) Then
                        '..... Enhancement Active .........................
                        For j = 1 To numEnhancedStocks
                            '..... Determine if this stock is supplemented ........
                            If Enh_stockIndex(j) = stk Then
                                '..... Compute maximum # spawners allowed for broodstock
                                'eq 191
                                MaxBrood = EnhProp(j) * escapement
                                '..... Determine number of smolts required to meet production goal
                                If SmoltProd(j, yr) > 0 Then
                                    '..... Determine number of spawners required to meet smolt production goal
                                    enhSpawn = FEnhSpawners(SmoltProd(j, yr), SmoltSurvRt(j), AValue(j))
                                    If enhSpawn > 0 Then
                                        '..... Production increase is indicated
                                        '..... Determine if the number of spawners required exceeds policy limit
                                        'eq 192
                                        If enhSpawn > MaxBrood Then enhSpawn = MaxBrood
                                        '..... Compute Age 1 recruits from supplemental production
                                        '..... Determine relative enhancement efficiency for this stock
                                        EnhEff = EnhEff_(j)
                                    End If
                                    '..... Decrease number of natural spawners by number removed for broodstock
                                    OldAdultEsc = escapement 'this is not used
                                    'eq 195
                                    escapement = escapement - enhSpawn
                                End If 'SmoltProd(j%, yr%) > 0
                                '..... Compute natural production
                                GoTo NaturalProd
                            End If 'Enh_stockIndex(j%) = stk%
                        Next j '= 1 To numEnhancedStocks
                        '..... This is a natural stock with no supplemental production.
                        GoTo NaturalProd
                    Else
                        '..... No enhancement changes have been specified.
                        '..... Procedure to compute natural production from escapements.
NaturalProd:
                        If escapement >= OPT Then
                            '..... At least the base production can be achieved.
                            '..... Determine if Escapement is to be truncated at MSH ....
                            If SkipMSH(stk) = 0 Then
                                '..... Truncate escapement at MSH
                                '..... compute additional catch for MSH
                                'surplus = escapement - OPT
                                tmp = escapement - OPT
                                'If surplus < 0 Then surplus = 0
                                If tmp < 0 Then tmp = 0
                                '..... Set escapement at MSH
                                OldAdultEsc = escapement
                                escapement = OPT
                                If isCalibration = True Then ''stage 1 or 2 calibration
                                    If OldAdultEsc > 0 Then
                                        'proportionOPT =OPT/escapement
                                        tmp = escapement / OldAdultEsc
                                        For age = 2 To NumAges
                                            '(yr%, stk%, age%) = proportionOPT * TotalEscape_EV(yr%, stk%, age%)
                                            totalEscape_EV(yr, stk, age) = tmp * totalEscape_EV(yr, stk, age)
                                        Next age
                                    End If
                                End If
                                '..... Add catch to total
                                'there may be a bug here, if isCalibration = True then tmp = proportionOPT, not surplus
                                'could it be if isCalibration = True then don't add surplus to TotAECat(stk%)?
                                totAECat(stk) = totAECat(stk) + tmp
                            End If
                        End If
                        '..... Compute inter-actions with production from wild fish
                        Select Case densityDependenceFlag
                            Case 1
                                '..... Density Dependent ......................
                                Select Case enhSpawn
                                    Case Is > 0
                                        EffEscape = escapement + enhSpawn * EnhEff
                                    Case Else
                                        EffEscape = escapement
                                End Select

                                'age1Fish = total adult return/% age 1 survival to adult rate
                                'eq 180 1 of 3 instances
                                age1Fish = TotalAdultReturn(EffEscape, stk, yr) / Age1toAdultSurvivalRate(stk)

                                '..Added 10/30/01 by John Carlile to handle negative enhancement..
                                If Enh_stockIndex(j) = stk Then
                                    If SmoltProd(j, yr) < 0 Then
                                        age1Fish = age1Fish + SmoltProd(j, yr) * SmoltSurvRt(j)
                                    End If
                                End If

                            Case Else
                                '..... Density Independent ....................
                                'age1Fish = total adult return/% age 1 survival to adult rate
                                'eq 180 2 of 3 instances
                                age1Fish = TotalAdultReturn(escapement, stk, yr) / Age1toAdultSurvivalRate(stk)

                                If enhSpawn > 0 Then
                                    '..... Add new smolt production ............
                                    'eq 196
                                    age1Fish = age1Fish + SmoltProd(j, yr) * SmoltSurvRt(j)

                                End If
                        End Select 'densityDependenceFlag
                    End If 'numEnhancedStocks > 0 And FActive(EnhStart%, EnhEnd%, yr%)
NextComp:
            End Select 'HatchFlg%(stk%)
            '..... Compute age 1 cohort size
            '*********** 1/96 SPRING STOCK PROVISION **********************
            If HatchFlg(stk) > 1 Then
                '..... This is a spring stock, delay recruitment by a year
                'eq 248
                cohort_(1, stk) = Spring(stk) * EV(yr + NumAges, stk)

                Spring(stk) = age1Fish
            Else
                '..... This is a fall stock
                'eq 248
                cohort_(1, stk) = age1Fish * EV(yr + NumAges, stk)
            End If
            '********** END CHANGE ****************************************

        Next stk

    End Sub

    '--------------------------------------------------------------
    Sub CalcTerminalRunAndEscapement()
        '--------------------------------------------------------------
        'purpose:  calculate terminal run (ocean escapement) and spawing escapement
        '--------------------------------------------------------------

        Dim age, CNRIndx As Integer
        Dim CNRTot As Double
        Dim fish As Integer
        Dim OcnTerm As Double
        Dim stk As Integer
        'Dim str2, str4, str7, str8, str9, str10 As String
        Dim TotTerm As Double

        
        For stk = 1 To numStocks
            If timeStep = 1 Then
                totalAdultEscape_SR(yr, stk) = 0
                trueTermRun(yr, stk) = 0
            End If 'If timeStep% = 1
            totEsc_allAges(stk) = 0

            If AEMethod = 2 Then AEQCohort = 0

            For age = 2 To NumAges
                saveCohort(age, stk) = cohort_(age, stk)
            Next age

            For age = 2 To NumAges
                TotTerm = 0
                OcnTerm = 0
                '....... Code for option to compute OER and TER using cohort method
                If AEMethod = 2 Then
                    '..... Compute total adult equivalent size of all ages in this year
                    '..... note Cohort_ in the next line is cohort size at the start of the year
                    AEQCohort = AEQCohort + (cohort_(age, stk) / survivalRate(age)) * AdltEqv(age, stk) * survivalRate(age)
                End If

                '..... ScrCohrt() is now the cohort size at the end of the year or time step
                '..... i.e. scrCohrt = starting cohort size - preterm catch - terminal run (fish that matured)
                '..... ScrCohrt() is not used in the remainder of the year or timestep loop
                'eq 161
                'PrintLine(traceFileID, "before", age, stk, cohort_(age, stk))
                cohort_(age, stk) = scrCohrt(age, stk)
                'PrintLine(traceFileID, "after", age, stk, cohort_(age, stk))
                'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, ScrCohrt(age%, stk%); "scrCohrt after terminal catch, save as cohort"
                For fish = 1 To numFisheries
                    If timeStepFlag(fish, timeStep) > 0 Then
                        If FTerminal(stk, fish, age) = Term Then
                            '..... Terminal fishery
                            If CNRFlag(yr, fish) = 1 Then
                                '..... Compute CNR stats for terminal fishery
                                CNRIndx = pointerCNR(fish)
                                CNRTot = CNRLegal(Term, CNRIndx, stk, age) + CNRShakCat(Term, CNRIndx, stk, age)
                            Else
                                CNRTot = 0
                            End If

                            'eq 162
                            TotTerm = TotTerm + legalCatch(Term, fish, stk, age) + stkShakCat(Term, fish, stk, age) + CNRTot
                            'eq 168
                            If oceanNetFlag(fish) And terminalFlag(stk, fish) = 0 Then OcnTerm = OcnTerm + legalCatch(Term, fish, stk, age) + stkShakCat(Term, fish, stk, age) + CNRTot
                        End If 'FTerminal(stk%, fish%, age%) = Term%
                    End If 'timeStepFlag(fish%, timeStep%) > 0
                Next fish

                '..... Compute actual escapement prior to pre-spawner mortality
                'eq 163
                Escape(age, stk) = FMinZed(TermRun(age, stk) - TotTerm)
                'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, Escape; "escape"; TermRun(age%, stk%); "termRun"; TotTerm; "term catch for all fisheries"; OcnTerm; "OcnTerm"; age%; "age"; stk%
                '..... Compute ocean escapement eq 167
                EscapeT(age, stk) = FMinZed(TermRun(age, stk) - OcnTerm)

                '..... Compute true total terminal run size ...................
                If age >= TrueTRage Then
                    '..... Only use array if requested (otherwise it does not exist)
                    If SaveFile(1) > 0 Or terminalRunReport = "y" Then
                        'TrueTermRun() is used in stage 2 calibration
                        trueTermRun(yr, stk) = trueTermRun(yr, stk) + EscapeT(age, stk)
                    End If
                End If

                '..... Save total terminal catch to calculate EV when calibrating, eq 166
                If isCalibration = True Then OceanEscape_EV(yr, stk, age) = OceanEscape_EV(yr, stk, age) + EscapeT(age, stk) 'stage 1 or 2 calibration
                totAECat(stk) = totAECat(stk) + TotTerm

                totEsc_allAges(stk) = totEsc_allAges(stk) + Escape(age, stk)

                If isCalibration = True Then 'stage 1 or 2 calibration
                    '..... Save total escapement to calculate EV when calibrating, eq 165
                    'apply interdam loss rates eq 171
                    totalEscape_EV(yr, stk, age) = totalEscape_EV(yr, stk, age) + Escape(age, stk) * IDL(stk, yr)
                End If

                '..... save total adult escapement to calculate progeny from spawning escapement, eq 164
                If age > 2 Then totalAdultEscape_SR(yr, stk) = totalAdultEscape_SR(yr, stk) + Escape(age, stk) * IDL(stk, yr)

            Next age

        Next stk


        If lastRepeat = 1 Then
            '..... call to write relevant data to file for evaluation of calibration
            'calibCheck (evaluate calibrations) is NOT the same as chkClb (compares projection runs)
            If isCalibration = True Then Call CalibCheck(yr, 4) 'stage 1 or 2 calibration

            'If timeStep = numTimeStepsSimulated Then
            'Clean up CNR Data Arrays
            'quick basic will print/write an empty variable as zero
            'visual basic will not
            'also
            'visual basic will initialize variants as empty, not zero
            'TotCNRFish(0, yr, Fish%) is assigned to array_()
            'if both are defined as double, then the empty problem goes away
            '           For fish = 1 To NumCNRFish
            '           If IsNothing(TotCNRFish(0, yr, fish)) Then TotCNRFish(0, yr, fish) = 0
            '           If IsNothing(TotCNRFish(1, yr, fish)) Then TotCNRFish(1, yr, fish) = 0
            '           Next fish
            'End If 'If timeStep% = numTimeStepsSimulated

        End If


    End Sub

    '--------------------------------------------------------------
    Sub CalcTotalFishingMortalities()
        '--------------------------------------------------------------
        'purpose:  calculate exploitation rates, stock composition of harvest
        '          and accum data for relative abundance
        '--------------------------------------------------------------


        Dim CNRIndx, age, fish As Integer
        'Dim stockMonTest
        Dim stk As Integer
        Dim str2, str9 As String
        Dim temp, tmp As Double
        Dim TmpAE3, TmpAE, TmpAE2, TmpAE4 As Double
        Dim accum_CNRsubLegal, accum_CNRlegal, accum_Legal As Double
        Dim accum_shakers As Double

        str2 = New String(" ", 2)
        str9 = New String(" ", 9)

        'code to calculate total preterminal catch is found in
        'calcPreterminalHarvest because the total preterminal catch
        'has to be subtracted from scrCohrt to obtain the terminal run
        'Since it is already calculated, it is not repeated here

        'stockMonTest = 0
        'If isCalibration = True And stk = StockMon Then stockMonTest = 1 'stage 1 or 2 calibration

        'note the next 2 lines of code could be placed at then end of calcTerminalHarvest
        'after the ceiling calculations
        'Shakers and CNR are not used in the ceiling harvest calculations
        'To speed up the program it is placed here so it is not called over and over
        'again in the iterative process to find the perterminal ceiling harvest rates
        'eq 151
        Call CalcEncounterRatio(Term)
        '..... Compute shaker loss
        'eq 152
        Call CalcShakersAndCNR(Term)
        'end of terminal harvest code placed here for expediency

        If isCalibration = False And SaveFile(7) Then 'projection run, abundance report requested at the start of the year
            If timeStep = 1 Then
                'eq 153
                Call CalcRelativeAbundance()
            End If 'If timeStep% = 1
        End If

        'these variables are for the sim, tim, lim.CSV files
        If timeStep = 1 Then
            For fish = 1 To numFisheries
                CNRIndx = pointerCNR(fish)
                TotCNRFish(0, yr, CNRIndx) = 0
                TotCNRFish(1, yr, CNRIndx) = 0
                TotCatFish_(yr, fish) = 0
                totShakFish(yr, fish) = 0
            Next fish
        End If 'If timeStep% = 1

        'the following are code to calculate total combined preterminal and terminal catch
        For stk = 1 To numStocks
            '..... Code for CNR fishery ........................
            For fish = 1 To numFisheries
                If timeStepFlag(fish, timeStep) > 0 Then
                    isCNR = False
                    If CNRFlag(yr, fish) = 1 Then
                        isCNR = True
                        CNRIndx = pointerCNR(fish)
                    End If

                    'the following variable are used to calculate abundance indices
                    If timeStep = 1 Then
                        accum_shakers = 0
                        accum_CNRlegal = 0
                        accum_CNRsubLegal = 0
                        accum_Legal = 0
                        If Stm > 10 Then
                            TmpAE = 0
                            TmpAE2 = 0
                            TmpAE3 = 0
                            TmpAE4 = 0
                        End If
                    End If 'If timeStep% = 1

                    For age = 2 To NumAges
                        If isCNR = True Then
                            accum_CNRlegal = accum_CNRlegal + CNRLegal(PreTerm, CNRIndx, stk, age) + CNRLegal(Term, CNRIndx, stk, age)
                            accum_CNRsubLegal = accum_CNRsubLegal + CNRShakCat(PreTerm, CNRIndx, stk, age) + CNRShakCat(Term, CNRIndx, stk, age)
                        End If
                        accum_shakers = accum_shakers + stkShakCat(PreTerm, fish, stk, age) + stkShakCat(Term, fish, stk, age)

                        Select Case Stm
                            Case 0 '..... Do not compute stock distribution data
                            Case Is < 10 '..... Stock distributions in nominal terms
                                accum_Legal = accum_Legal + legalCatch(PreTerm, fish, stk, age) + legalCatch(Term, fish, stk, age)
                            Case Is > 10 '..... Stock distributions in adult equivalents
                                temp = AdltEqv(age, stk)
                                TmpAE = TmpAE + (stkShakCat(PreTerm, fish, stk, age) + stkShakCat(Term, fish, stk, age)) * temp
                                TmpAE4 = TmpAE4 + (legalCatch(PreTerm, fish, stk, age) + legalCatch(Term, fish, stk, age)) * temp
                                If isCNR = True Then
                                    TmpAE2 = TmpAE2 + (CNRLegal(PreTerm, CNRIndx, stk, age) + CNRLegal(Term, CNRIndx, stk, age)) * temp
                                    TmpAE3 = TmpAE3 + (CNRShakCat(PreTerm, CNRIndx, stk, age) + CNRShakCat(Term, CNRIndx, stk, age)) * temp
                                End If
                        End Select

                    Next age
                    totShakFish(yr, fish) = totShakFish(yr, fish) + accum_shakers
                    If isCNR = True Then
                        'TotCNRFish() is used in the projection run
                        TotCNRFish(0, yr, CNRIndx) = TotCNRFish(0, yr, CNRIndx) + accum_CNRlegal
                        TotCNRFish(1, yr, CNRIndx) = TotCNRFish(1, yr, CNRIndx) + accum_CNRsubLegal
                        totShakFish(yr, fish) = totShakFish(yr, fish) + accum_CNRlegal + accum_CNRsubLegal
                    End If
                    Select Case Stm
                        Case 0, 4 '..... Do not generate stock distribution stats..

                        Case 1 '..... Total Mortality ..........................
                            stkMort(yr, fish, stk) = accum_shakers + accum_CNRlegal + accum_CNRsubLegal + accum_Legal
                        Case 11 '..... Total Mortality in Adult Equivalents .....
                            stkMort(yr, fish, stk) = TmpAE + TmpAE2 + TmpAE3 + TmpAE4

                        Case 2 '..... Catch ....................................
                            stkMort(yr, fish, stk) = accum_Legal
                        Case 12 '..... Catch in Adult Equivalents ...............
                            stkMort(yr, fish, stk) = TmpAE4

                        Case 3 '..... Total Incidental Mortality ...............
                            stkMort(yr, fish, stk) = accum_shakers + accum_CNRlegal + accum_CNRsubLegal
                        Case 13 '..... Total Incidental Mortality in Adult Equivalents.....
                            stkMort(yr, fish, stk) = TmpAE + TmpAE2 + TmpAE3

                    End Select
                End If 'timeStepFlag(fish%, timeStep%) > 0
            Next fish

            '..... Code to compute Ocean Exploitation Rate (OER) AND
            '..... Total Exploitation Rate (TER).
            If HarvestRateReport = "Y" Or SaveFile(3) <> 0 Or SaveFile(4) <> 0 Then
                If AEMethod = 2 Then
                    '..... Use total adult equivalent population at start of preterm season
                    tmp = AEQCohort
                Else
                    '..... Use adult equivalent catch and escapements
                    tmp = totAECat(stk) + totEsc_allAges(stk)
                End If
                '..... Compute Adult Equivalent Exploitation Rates ...........
                If tmp = 0 Then
                    TotAEHR(yr, stk) = -1.0!
                    ocnAEHR(yr, stk) = -1
                Else
                    TotAEHR(yr, stk) = totAECat(stk) / tmp
                    ocnAEHR(yr, stk) = OcnAECat(stk) / tmp
                End If
            End If

        Next stk

        For fish = 1 To numFisheries
            If timeStepFlag(fish, timeStep) > 0 Then
                '..... Compute total catch by fishery
                'TotCatFish_(fish%) is used in stage 2 calibration
                TotCatFish_(yr, fish) = TotCatFish_(yr, fish) + tempCat(PreTerm, fish) + tempCat(Term, fish)
            End If 'timeStepFlag(fish%, timeStep%) > 0
        Next fish

        If FActive(BaseCeilStart, BaseCeilingEnd) Then
            '..... Set Up New Catch Ceilings
            Call CeilingLevel(1, Term)
        End If

        If yr = CeilingEnd And NumCeil > 0 Then
            '..... after RT and RT_scaler has been found
            '..... to calculate the ceiling catch correcly,
            '..... save the adjusted harvest rates
            Call SaveHarvestRatesAfterCeilingFishery()
        End If

        'xxxxxxxxxx  Dell changed 3/6/96 cause EVType%() already deleted in CHINPUT
        'estimate catches for Lower Georgia Strait Hatchery Stock (GSH) to GS troll and sport fisheries.
        'IF isCalibration = True AND EVType%(9) = 3 THEN CALL GSHatchCalib(Yr%)
        If isCalibration = True Then 'stage 1 or 2 calibration 'note EVType% 3 and 4 is never used and GSHatchCalib(yr%) is never called
            If EVType(9) = 3 Then Call GSHatchCalib()
        End If
        'xxxxxxxxxx

        If NumStkCatch > 0 Then
            '..... Compute stock compositions
            Call CalcHarvestStockComposition()
        End If



    End Sub


    '--------------------------------------------------------------
    Sub AdvanceCohortAge()
        '--------------------------------------------------------------
        'purpose:  age cohorts
        'note this must follow calcTerminalRunAndEscapement
        'because Cohort_ = scrCohrt
        'and precede CalcAge1Production
        'because need to advance age 1 to age 2 before estimating new cohort of age 1
        'from spawning escapement (to prevent overwriting the previous year age 1)
        '--------------------------------------------------------------

        Dim age, stk As Integer
        For stk = 1 To numStocks
            'bump all ages up by 1
            For age = NumAges To 2 Step -1
                cohort_(age, stk) = FMinZed(cohort_(age - 1, stk))
            Next age
        Next stk

    End Sub


    '--------------------------------------------------------------
    Sub ApplySurvivalRates()
        '--------------------------------------------------------------
        'purpose: calculate cohort size after natural mortality
        'note this must occur at the start of the year before any fisheries
        '--------------------------------------------------------------

        Dim age, stk As Integer
        ReDim scrCohrt(NumAges, numStocks)
        ReDim timeStep1Cohort(NumAges, numStocks)
        ReDim timeStep1TermRun(NumAges, numStocks)

        For stk = 1 To numStocks

            For age = 1 To NumAges

                'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, Cohort_(age%, stk%);
                'apply survival rate eq 01
                'i.e. cohort(at start of this year) = cohort(at end of previous year) * survivalRate
                cohort_(age, stk) = cohort_(age, stk) * survivalRate(age)
                'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, Cohort_(age%, stk%); "Cohort after apply survival rates"
            Next age

        Next stk

    End Sub

    '--------------------------------------------------------------
    Sub ApplyMaturityRates()
        '--------------------------------------------------------------
        'purpose: apply instantaneous maturity rates before
        '         terminal fishery and spawning escapement
        '         but only if fisheries are active during the time step
        '--------------------------------------------------------------

        Dim age As Integer
        Dim applyMaturity As Integer
        Dim fish, stk As Integer
        Dim instantaneousRate As Double
        ReDim maturityRateApplied(numStocks, NumAges)


        For stk = 1 To numStocks
            For age = 2 To NumAges
                For fish = 1 To numFisheries
                    If timeStepFlag(fish, timeStep) > 0 Then
                        applyMaturity = 0
                        'if terminal fishery
                        If FTerminal(stk, fish, age) = Term Then applyMaturity = 1
                        'if only 1 time step
                        If numTimeStepsSimulated = 1 Then applyMaturity = 1
                        'if last time step
                        If timeStep = numTimeStepsSimulated Then applyMaturity = 1
                        'if maturity has not already been applied
                        If applyMaturity = 1 And maturityRateApplied(stk, age) = 0 Then

                            '                    If numTimeStepsActuallyFished(fish%) = 1 Then instantaneousRate = maturationRate(age%, stk%)
                            '                    If numTimeStepsActuallyFished(fish%) > 1 Then
                            'eq 135
                            '                        If maturationRate(age%, stk%) < 1# Then
                            '                            instantaneousRate = 1 - Exp(Log(1 - maturationRate(age%, stk%)) / numTimeStepsSimulated)
                            '                        Else 'for last age where maturation = 1.0
                            '                            instantaneousRate = 0.9999
                            '                        End If
                            '                    End If
                            'use the next line replicate QB
                            'because exp(log(1-maturationRate)) will add more decimal places
                            'and change stage 1 and 2 EV and abundances
                            instantaneousRate = maturationRate(age, stk)


                            '..... Compute terminal run eq 134
                            TermRun(age, stk) = scrCohrt(age, stk) * instantaneousRate

                            'If stk% = 5 And age% = 4 And yr% = 0 Then Print #97, ScrCohrt(age%, stk%); instantaneousRate; TermRun(age%, stk%); "apply maturity"; fish%; timeStep%; maturationRate(age%, stk%)
                            '                    TotalTermRun(age%, stk%) = TotalTermRun(age%, stk%) + TermRun(age%, stk%)

                            '..... Compute immature fish remaining in ocean
                            '..... scrCohrt = starting cohort size - preterm catch - terminal run (fish that matured)
                            scrCohrt(age, stk) = FMinZed(scrCohrt(age, stk) - TermRun(age, stk))
                            'PrintLine(traceFileID, "applyMat", age, stk, scrCohrt(age, stk), TermRun(age, stk))
                            'eq 136
                            maturityRateApplied(stk, age) = 1
                        End If 'applyMaturity = 1 and maturityRateApplied(stk%, age%) = 0
                    End If 'timeStepFlag(fish%, timeStep%) > 0
                Next fish
            Next age
        Next stk

    End Sub

    '--------------------------------------------------------------
    Sub CeilingHarvest(ByRef location As Integer)
        '--------------------------------------------------------------
        '  Purpose:  Computes scalars to adjust catches to ceiling levels (RT),
        '            multiply modeled catches by RT in order to match ceilings,
        '            and accumulate total catch by stock and by fishery.

        '  Arguments: Yr%  = Simulation year
        '             location = Terminal PreTerminal
        '
        '  Inputs:    Ceil()
        '             CeilingControlFlag%()
        '             CeilingFisheryFlag()
        '             legalCatch()
        '             NewScal()
        '             NumAges
        '             NumCeil
        '             numStocks
        '             RT()
        '             TempCat()
        '
        '  Called By: ceilingSetup
        '
        '  Output:    legalCatch()
        '             TempCat()
        '             TempStkCat()
        '             RT()
        '
        '  Externals: None.
        '
        '
        '--------------------------------------------------------------
        'note RT adjusts only perterminal or terminal ceiling catch (not both)
        'However, if you need to adjust the preterminal catch up slightly,
        'then you need to adjust the terminal catch down slightly so the total will be correct
        'Therefore, RT_scaler is used to adjusts the combined modeled ceiling catch
        'RT_scaler is calculated in CeilingTestIfModeledCatchMatchesObserved

        Dim age, fish, stk As Integer
        Dim RT_workingCopy As Double

        For fish = 1 To numFisheries
            If timeStepFlag(fish, timeStep) > 0 Then
                '..... check to see if this is a ceiling fishery
                If CeilingFisheryFlag(fish) = 1 Then
                    '..... Temporary RT
                    RT_temp(location, fish) = 1
                    '..... Compute RT, ratio between ceiling and model catch eq 25
                    If tempCat(location, fish) > 0 Then RT_temp(location, fish) = ceil(location, fish) / tempCat(location, fish)
                    'eq 27
                    If RT_temp(location, fish) > 1 Then

                        '..... Check if ceiling is to be forced
                        'eq 28
                        If CeilingControlFlag(yr, fish) = 0 Then RT_temp(location, fish) = 1

                    End If
                    If isOcnTerm(fish) = True Then
                        '..... Compute RT factor for pre-terminal fishery eq 146
                        RT_workingCopy = RT_temp(PreTerm, fish) * RT_scalar(fish)
                    Else
                        '..... Compute RT factor for either pre-terminal or terminal fishery
                        RT_workingCopy = RT_temp(location, fish)
                    End If

                    '..... Temporary accumulator variable
                    tempCat(location, fish) = 0

                    '..... Compute new ceiling catches
                    For stk = 1 To numStocks
                        '..... Temporary accumulator variable for stock catch
                        TempStkCat(location, fish, stk) = 0
                        For age = 2 To NumAges
                            'eq 26.....Compute catch adjusted for ceiling management
                            legalCatch(location, fish, stk, age) = legalCatch(location, fish, stk, age) * RT_workingCopy
                            '..... Accumulate total catch by age for each stock
                            TempStkCat(location, fish, stk) = TempStkCat(location, fish, stk) + legalCatch(location, fish, stk, age)
                        Next age
                        '..... Accumulate total catch for all stocks
                        tempCat(location, fish) = tempCat(location, fish) + TempStkCat(location, fish, stk)
                    Next stk

                    'save a copy of RT with an index for year eq 148
                    RT(yr, location, fish) = RT_workingCopy

                End If
            End If 'timeStepFlag(fish%, timeStep%) > 0
        Next fish

    End Sub


    Sub write_ISBM()

        'This subroutine will print the following
        '(1) ocean cohort (pre-natural mortality) directly from the ccc file

        '(2) ISBM cohort:  either (a) cohort X survival rates
        '    or (b) terminal run + ocean net catch of age 4 & 5 fish
        '    depending on the terminal run and ocean net flags

        '(3) catch, shakers, and CNR mortalities multipled by AEQ = total mortality for ISBM purposes.

        '(4) exploitation rate based on (total mortality)/ISBM cohort (not ocean cohort)

        'This program will change terminal flag = 1 for all stocks if fishery = 15 (terminal net) or 25 (terminal sport)
        'so far, termFlag is used only in the calculation of cohort for ISBM
        'if termFlag is used elsewhere, may have to rewrite the code to restore orginial value of termFlag


        Dim accumPTcatch As Double, accumObscatch As Double
        Dim age As Single, ag As Single, AEQ As Double
        Dim CNRIndx As Integer
        Dim exploitationRate(,,) As Double
        Dim fish As Single
        Dim ISBMcohort(,,) As Single
        Dim ISBMstock As Integer
        Dim netFish As Integer
        Dim stk As Single
        Dim temp As String
        Dim totalMortality(,,) As Single

        ReDim exploitationRate(numFisheries, numStocks, NumAges)
        ReDim ISBMcohort(numFisheries, numStocks, NumAges)
        ReDim totalMortality(numFisheries, numStocks, NumAges)

        '.....begin code open files and initialize arrays only during first year
        If yr = 0 Then
            ReDim is_Canadian(numStocks)
            ReDim is_ISBM_fishery(numFisheries)
            ReDim is_ISBM_stock(numStocks)

            cohortFileID = FreeFile()
            temp = saveFilePrefix & "_oceanCohort.csv"
            FileOpen(cohortFileID, temp, OpenMode.Output)

            termRunFileID = FreeFile()
            temp = saveFilePrefix & "_termRun.csv"
            FileOpen(termRunFileID, temp, OpenMode.Output)

            spawnEscFileID = FreeFile()
            temp = saveFilePrefix & "_spawnEsc.csv"
            FileOpen(spawnEscFileID, temp, OpenMode.Output)

            shakerFileID = FreeFile()
            temp = saveFilePrefix & "_shaker.csv"
            FileOpen(shakerFileID, temp, OpenMode.Output)

            catchFileID = FreeFile()
            temp = saveFilePrefix & "_catch.csv"
            FileOpen(catchFileID, temp, OpenMode.Output)

            AEQfileID = FreeFile()
            temp = saveFilePrefix & "_AEQ.csv"
            FileOpen(AEQfileID, temp, OpenMode.Output)

            totalMortFileID = FreeFile()
            temp = saveFilePrefix & "_totalMort_AEQ.csv"
            FileOpen(totalMortFileID, temp, OpenMode.Output)

            exploitationRateFileID = FreeFile()
            temp = saveFilePrefix & "_exploitationRate.csv"
            FileOpen(exploitationRateFileID, temp, OpenMode.Output)

            ISBMcohortFileID = FreeFile()
            temp = saveFilePrefix & "_ISBMcohort.csv"
            FileOpen(ISBMcohortFileID, temp, OpenMode.Output)

            termRun_OcnNetFileID = FreeFile()
            temp = saveFilePrefix & "_termRun_ocnNet.csv"
            FileOpen(termRun_OcnNetFileID, temp, OpenMode.Output)

            CNRfileID = FreeFile()
            temp = saveFilePrefix & "_CNR.csv"
            FileOpen(CNRfileID, temp, OpenMode.Output)

            ISBMfileID = FreeFile()
            temp = saveFilePrefix & "_ISBM_index.csv"
            FileOpen(ISBMfileID, temp, OpenMode.Output)

            For stk = 1 To numStocks

                is_ISBM_stock(stk) = False
                is_Canadian(stk) = False
                Select Case stk
                    Case 2 'Northern/Central B.C.
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 3 'Fraser Early
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 4 'Fraser Late
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 5 'WCVI Hatchery
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 6 'WCVI Natural
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 7 'Upper Strait of Georgia
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 8 'Lower Strait of Georgia Natural
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 9 'Lower Strait of Georgia Hatchery
                        is_ISBM_stock(stk) = True
                        is_Canadian(stk) = True
                    Case 12 'Puget Sound Natural F
                        is_ISBM_stock(stk) = True
                    Case 14 'Nooksack Spring
                        is_ISBM_stock(stk) = True
                    Case 15 'Skagit Wild
                        is_ISBM_stock(stk) = True
                    Case 16 'Stillaguamish Wild
                        is_ISBM_stock(stk) = True
                    Case 17 'Snohomish Wild
                        is_ISBM_stock(stk) = True
                    Case 18 'Washington Coastal Hatchery
                        is_ISBM_stock(stk) = True
                    Case 19 'Col. Upriver Brights
                        is_ISBM_stock(stk) = True
                    Case 23 'Lewis River Wild
                        is_ISBM_stock(stk) = True
                    Case 26 'Columbia River Summers
                        is_ISBM_stock(stk) = True
                    Case 27 'Oregon Coastal
                        is_ISBM_stock(stk) = True
                End Select

                If is_ISBM_stock(stk) = True Then

                    If is_Canadian(stk) = True Then
                        For fish = 1 To numFisheries
                            is_ISBM_fishery(fish) = False
                            Select Case fish
                                Case 3 'Central B.C. Troll
                                    is_ISBM_fishery(3) = True
                                Case 6 'Strait of Georgia Troll
                                    is_ISBM_fishery(6) = True
                                Case 8 'Northern B.C. Net
                                    is_ISBM_fishery(8) = True
                                Case 9 'Central B.C. Net
                                    is_ISBM_fishery(9) = True
                                Case 10 'WCVI Net
                                    is_ISBM_fishery(10) = True
                                Case 11 'Juan de Fuca Net
                                    is_ISBM_fishery(11) = True
                                Case 16 'Johnstone Strait Net
                                    is_ISBM_fishery(16) = True
                                Case 17 'Fraser Net
                                    is_ISBM_fishery(17) = True
                                Case 24 'Strait of Georgia Sport
                                    is_ISBM_fishery(24) = True
                            End Select
                        Next fish
                    End If 'is_Canadian(stk) = true

                    If is_Canadian(stk) = False Then
                        For fish = 1 To numFisheries
                            is_ISBM_fishery(fish) = False
                            Select Case fish
                                Case 5 'WA/OR troll
                                    is_ISBM_fishery(5) = True
                                Case 12 'Puget Sound northern net
                                    is_ISBM_fishery(12) = True
                                Case 13 'Puget Sound southern net
                                    is_ISBM_fishery(13) = True
                                Case 14 'WA coastal net
                                    is_ISBM_fishery(14) = True
                                Case 21 'WA coastal sport
                                    is_ISBM_fishery(21) = True
                                Case 22 'Puget Sound northern sport
                                    is_ISBM_fishery(22) = True
                                Case 23 'Puget Sound southern sport
                                    is_ISBM_fishery(23) = True
                                Case 15 'freshwater terminal net
                                    is_ISBM_fishery(15) = True
                                Case 25 'freshwater terminal sport
                                    is_ISBM_fishery(25) = True
                            End Select
                        Next fish
                    End If 'is_Canadian(stk) = false
                End If 'is_ISBM_stock(stk) = true
            Next stk
        End If 'yr = 0
        '.....end code open files and initialize arrays only during first year


        '.....begin code to print ISBM csv files for all years:  same data from ccc file
        For stk = 1 To numStocks
            If is_ISBM_stock(stk) = True Then
                Write(termRunFileID, 1979 + yr, stk, stock_shortName(stk))
                Write(spawnEscFileID, 1979 + yr, stk, stock_shortName(stk))
                Write(cohortFileID, 1979 + yr, stk, stock_shortName(stk))
                For age = 2 To NumAges
                    Write(termRunFileID, EscapeT(age, stk)) 'terminal run = ocean escapement
                    Write(spawnEscFileID, Escape(age, stk) * IDL(stk, yr))
                    Write(cohortFileID, oceanCohort(stk, age))
                Next age
                WriteLine(termRunFileID)
                WriteLine(spawnEscFileID)
                WriteLine(cohortFileID)
                For fish = 1 To numFisheries
                    If is_ISBM_fishery(fish) = True Then
                        Write(catchFileID, 1979 + yr, stk, stock_shortName(stk), fish, FisheryName(fish))
                        Write(shakerFileID, 1979 + yr, stk, stock_shortName(stk), fish, FisheryName(fish))
                        Write(CNRfileID, 1979 + yr, stk, stock_shortName(stk), fish, FisheryName(fish))
                        Write(totalMortFileID, 1979 + yr, stk, stock_shortName(stk), FisheryName(fish))
                        Write(AEQfileID, 1979 + yr, stk, stock_shortName(stk), FisheryName(fish))

                        For ag = 2 To NumAges
                            '..... Determine CNR Fishery
                            CNRIndx = pointerCNR(fish)
                            Write(catchFileID, legalCatch(PreTerm, fish, stk, ag) + legalCatch(Term, fish, stk, ag))
                            Write(shakerFileID, stkShakCat(PreTerm, fish, stk, ag) + stkShakCat(Term, fish, stk, ag))
                            Write(CNRfileID, CNRLegal(PreTerm, CNRIndx, stk, ag) + CNRLegal(Term, CNRIndx, stk, ag) + CNRShakCat(Term, CNRIndx, stk, ag) + CNRShakCat(PreTerm, CNRIndx, stk, ag))

                            'AEQ = 1 if using termRun + oceannet for cohort (terminal fisheries)
                            AEQ = AdltEqv(ag, stk)
                            If (terminalFlag(stk, fish) = 1 Or oceanNetFlag(fish) = 1 And ag >= 4) Then
                                AEQ = 1.0
                            End If
                            'AEQ = 1 for (15) Freshwater net and (25) freshwater sport
                            If (fish = 15 Or fish = 25) Then
                                AEQ = 1.0
                            End If
                            Write(AEQfileID, AEQ)

                            totalMortality(fish, stk, ag) = (legalCatch(PreTerm, fish, stk, ag) + legalCatch(Term, fish, stk, ag) + stkShakCat(PreTerm, fish, stk, ag) + stkShakCat(Term, fish, stk, ag) + CNRLegal(PreTerm, CNRIndx, stk, ag) + CNRLegal(Term, CNRIndx, stk, ag) + CNRShakCat(Term, CNRIndx, stk, ag) + CNRShakCat(PreTerm, CNRIndx, stk, ag)) * AEQ
                            Write(totalMortFileID, totalMortality(fish, stk, ag))
                        Next ag

                        WriteLine(catchFileID)
                        WriteLine(shakerFileID)
                        WriteLine(totalMortFileID)
                        WriteLine(CNRfileID)
                        WriteLine(AEQfileID)
                    End If 'is_ISBM_fishery(fish) = True
                Next fish
            End If 'is_ISBM_stock(stk) = true
        Next stk
        '.....end code to print ISBM csv filesfor all years:  same data from ccc file


        '.....begin code to print ISBM csv files for all years:  data not found in ccc file
        For stk = 1 To numStocks
            For fish = 1 To numFisheries
                If is_ISBM_fishery(fish) = True Then
                    Write(ISBMcohortFileID, 1979 + yr, stk, stock_shortName(stk), fish, FisheryName(fish))
                    Write(exploitationRateFileID, 1979 + yr, stk, stock_shortName(stk), fish, FisheryName(fish))

                    'calculate cohort for ISBM
                    For age = 2 To NumAges
                        'if the fishery is a terminal fishery use the terminal run + ocean net catch for the cohort size
                        If (terminalFlag(stk, fish) = 1 Or oceanNetFlag(fish) = 1 And age >= 4) Then
                            ISBMcohort(fish, stk, age) = EscapeT(age, stk)
                            WriteLine(termRun_OcnNetFileID, stk, stock_shortName(stk), 1979 + yr, fish, FisheryName(fish), age, "TermRun", EscapeT(age, stk))
                            For netFish = 7 To 17
                                If (terminalFlag(stk, netFish) = 0 And oceanNetFlag(netFish) = 1 And age >= 4) Then
                                    ISBMcohort(fish, stk, age) = ISBMcohort(fish, stk, age) + legalCatch(PreTerm, netFish, stk, age) + legalCatch(Term, netFish, stk, age) + stkShakCat(PreTerm, netFish, stk, age) + stkShakCat(Term, netFish, stk, age) + CNRLegal(PreTerm, netFish, stk, age) + CNRLegal(Term, netFish, stk, age) + CNRShakCat(Term, netFish, stk, age) + CNRShakCat(PreTerm, netFish, stk, age)
                                    WriteLine(termRun_OcnNetFileID, stk, stock_shortName(stk), 1979 + yr, netFish, FisheryName(netFish), age, "OceanNet", legalCatch(PreTerm, netFish, stk, age) + legalCatch(Term, netFish, stk, age) + stkShakCat(PreTerm, netFish, stk, age) + stkShakCat(Term, netFish, stk, age) + CNRLegal(PreTerm, netFish, stk, age) + CNRLegal(Term, netFish, stk, age) + CNRShakCat(Term, netFish, stk, age) + CNRShakCat(PreTerm, netFish, stk, age), ISBMcohort(fish, stk, age))
                                End If
                            Next netFish
                        Else 'if not terminal fishery, then use cohort
                            ISBMcohort(fish, stk, age) = oceanCohort(stk, age) * survivalRate(age)
                        End If 'terminal fishery
                        Write(ISBMcohortFileID, ISBMcohort(fish, stk, age))
                    Next age
                    WriteLine(ISBMcohortFileID)

                    'calculate exploitation rate
                    For age = 2 To NumAges
                        exploitationRate(fish, stk, age) = 0
                        If ISBMcohort(fish, stk, age) > 0 Then
                            exploitationRate(fish, stk, age) = totalMortality(fish, stk, age) / ISBMcohort(fish, stk, age)
                            Write(exploitationRateFileID, exploitationRate(fish, stk, age))
                        Else
                            Write(exploitationRateFileID, " ")
                        End If
                    Next age
                    WriteLine(exploitationRateFileID)

                End If 'is_ISBM_fishery
            Next fish
        Next stk
        '.....end code to print ISBM csv files for all years:  data not found in ccc file



        'begin code to calculate average Base period exploitation rates
        If yr <= 3 Then 'only for base period
            For stk = 1 To numStocks
                For fish = 1 To numFisheries
                    For age = 2 To NumAges
                        baseER(fish, stk, age) = baseER(fish, stk, age) + exploitationRate(fish, stk, age)
                        If yr = 3 Then baseER(fish, stk, age) = baseER(fish, stk, age) / 4
                    Next age
                Next fish
            Next stk
        End If 'yr <= 3
        '.....end code to calculate average Base period exploitation rates


        '.....begin code to calculate ISBM index for all years after base period
        If yr >= 4 Then
            For stk = 1 To numStocks
                accumObscatch = 0
                accumPTcatch = 0
                ISBMstock = 0
                For fish = 1 To numFisheries
                    If is_ISBM_fishery(fish) = True Then
                        For age = 2 To NumAges
                            accumObscatch = accumObscatch + totalMortality(fish, stk, age)
                            accumPTcatch = accumPTcatch + ISBMcohort(fish, stk, age) * baseER(fish, stk, age)
                            ISBMstock = 1
                        Next age
                    End If 'is_ISBM_fishery
                Next fish

                If ISBMstock = 1 And accumPTcatch > 0 Then
                    WriteLine(ISBMfileID, 1979 + yr, stk, stock_shortName(stk), accumObscatch / accumPTcatch)
                End If
            Next stk
        End If 'Yr >= 4 
        '.....end code to calculate ISBM index for all years after base period


        'if last year, then close files
        If yr = NumYears - 2 Then
            'FileClose(AEQfileID)
            'FileClose(catchFileID)
            'FileClose(cohortFileID)
            'FileClose(CNRfileID)
            'FileClose(exploitationRateFileID)
            'FileClose(ISBMcohortFileID)
            'FileClose(ISBMfileID)
            'FileClose(spawnEscFileID)
            'FileClose(shakerFileID)
            'FileClose(termRun_OcnNetFileID)
            'FileClose(termRunFileID)
            ''FileClose(totalMortFileID)
        End If 'Yr = NumYears - 2

    End Sub

    Sub write_CCCfile()

        '--------------------------------------------------------------
        'purpose:  write age specific summary of cohort, AEQ, terminal run (ocean escapement), spawing escapement
        'by stock
        'write summary of total catch, shakers, CNR legal and sublegal mortality
        'by fishery
        '--------------------------------------------------------------

        Dim age, CNRIndx As Integer
        Dim fish As Integer
        Dim nameCCFile As String
        Dim stk As Integer
        Dim str2, str4, str7, str8, str9, str10 As String

        str2 = New String(" ", 2)
        str4 = New String(" ", 4)
        str7 = New String(" ", 7)
        str8 = New String(" ", 8)
        str9 = New String(" ", 9)
        str10 = New String(" ", 10)

        If yr = 0 Then
            CCCFileID = FreeFile()
            nameCCFile = saveFilePrefix & ".CCC"
            FileOpen(CCCFileID, nameCCFile, OpenMode.Output)
        End If

        For stk = 1 To numStocks
            '.....dump first line of ccc file
            If Stm = 4 Then

                'note legacy ccc file has 2 digit year
                str2 = RSet(VB6.Format(yr, "#0"), Len(str2))
                Print(CCCFileID, str2 & " ")

                '..... stock
                str2 = RSet(VB6.Format(stk, "#0"), Len(str2))
                Print(CCCFileID, str2 & " ")

                For age = 2 To NumAges
                    '..... AEQ
                    str7 = RSet(VB6.Format(AdltEqv(age, stk), "0.####0"), Len(str7))
                    Print(CCCFileID, str7 & "  ")

                    '..... cohort
                    'oceanCohort(stk, age) = cohort_(age, stk) / survivalRate(age)
                    'PrintLine(traceFileID, "ccc", stk, age, cohort_(age, stk))

                    oceanCohort(stk, age) = savecohort(age, stk) / survivalRate(age)
                    'PrintLine(traceFileID, "ccc", stk, age, savecohort(age, stk))
                    str8 = RSet(VB6.Format(oceanCohort(stk, age), "#######0"), Len(str8))
                    Print(CCCFileID, str8 & "  ")
                Next age
            End If

            For age = 2 To NumAges
                If Stm = 4 Then
                    '..... save true terminal run
                    str10 = RSet(VB6.Format(EscapeT(age, stk), "######0.#0 "), Len(str10))
                    Print(CCCFileID, str10)

                    '..... save spawning escapement
                    str10 = RSet(VB6.Format(Escape(age, stk) * IDL(stk, yr), "######0.#0"), Len(str10))
                    Print(CCCFileID, str10 & "  ")
                End If
            Next age

            If Stm = 4 Then
                '..... end of first line of ccc file
                PrintLine(CCCFileID)

                '..... Dump second line of ccc file
                For fish = 1 To numFisheries
                    For age = 2 To NumAges
                        '..... fishery
                        str2 = RSet(VB6.Format(fish, "#0"), Len(str2))
                        Print(CCCFileID, str2)

                        '..... age
                        str2 = RSet(VB6.Format(age, "0"), Len(str2))
                        Print(CCCFileID, str2)

                        '..... combined catch (legacy ccc file)
                        str9 = RSet(VB6.Format(legalCatch(PreTerm, fish, stk, age) + legalCatch(Term, fish, stk, age), "#####0.#0"), Len(str9))
                        Print(CCCFileID, str9 & " ")

                        '..... combined shakers (legacy ccc file)
                        str9 = RSet(VB6.Format(stkShakCat(PreTerm, fish, stk, age) + stkShakCat(Term, fish, stk, age), "#####0.#0"), Len(str9))
                        Print(CCCFileID, str9 & " ")

                        If CNRFlag(yr, fish) = 1 Then
                            '..... CNR legal (legacy ccc)
                            CNRIndx = pointerCNR(fish)
                            str9 = RSet(VB6.Format(CNRLegal(PreTerm, CNRIndx, stk, age) + CNRLegal(Term, CNRIndx, stk, age), "#####0.#0"), Len(str9))
                            Print(CCCFileID, str9 & " ")

                            '..... CNR sublegal(legacy ccc)
                            str9 = RSet(VB6.Format(CNRShakCat(PreTerm, CNRIndx, stk, age) + CNRShakCat(Term, CNRIndx, stk, age), "#####0.#0"), Len(str9))
                            PrintLine(CCCFileID, str9 & " ")

                        Else 'no CNR
                            '.....(legacy ccc, last 2 columns are blank)
                            PrintLine(CCCFileID)

                        End If
                    Next age
                Next fish
            End If

        Next stk

    End Sub


    Sub write_totalMortality()
        '--------------------------------------------------------------
        'purpose:  write details regarding total mortality, 
        'preterminal and terminal
        'age specific catch, shakers, dropoffs, CNR legal and sublegal mortality
        'by stock and fishery

        'for spreadsheets, there are more rows of fishery data than there are rows in a spreadsheet.
        'therefore split fishery data into an ISBM file and an AABM file

        'for years without CNR fisheries, zeros are used to fill the column.
        'blanks (zero length string) will cause problems with totals in ACCESS
        'not printing anything will cause problems for other VB programs 
        'unless there is another file to flag which years have CNR fisheries
        'and therefore inform other VB programs when to read or skip fields.

        'only total mortality is in AEQ and the csv file name will indicate that:  totalMort_AEQ.csv
        'everything else is in nominal units
        '--------------------------------------------------------------

        Dim age As Integer
        Dim nameCCFile As String
        Dim CNRIndx As Integer
        Dim CNR_legal_release_MortRate As Single, CNR_sublegal_release_MortRate As Single
        Dim CSVfileID As Integer
        Dim fish As Integer
        Dim number As Single
        Dim stk As Integer
        Dim shakersWithOutDropOffs As Single
        Dim subLegalDropoffs(1) As Single
        Dim str7, str8, str9, str10 As String

        If yr = 0 Then
            Dim codesFileID As Integer

            codesFileID = FreeFile()
            FileOpen(codesFileID, "fisheryCodes.csv", OpenMode.Output)
            WriteLine(codesFileID, "fishery", "fisheryShortName")
            For fish = 1 To numFisheries
                WriteLine(codesFileID, fish, FisheryName(fish))
            Next fish
            FileClose(codesFileID)

            codesFileID = FreeFile()
            FileOpen(codesFileID, "stockCodes.csv", OpenMode.Output)
            WriteLine(codesFileID, "stock", "stockShortName", "stockLongName")
            For stk = 1 To numStocks
                WriteLine(codesFileID, stk, stockAbbreviation(stk), stock_longName(stk))
            Next stk
            FileClose(codesFileID)

            CCCfish_ISBMFileID = FreeFile()
            nameCCFile = saveFilePrefix & "_fish_ISBM_CCC.csv"
            FileOpen(CCCfish_ISBMFileID, nameCCFile, OpenMode.Output)

            CCCfish_AABMFileID = FreeFile()
            nameCCFile = saveFilePrefix & "_fish_AABM_CCC.csv"
            FileOpen(CCCfish_AABMFileID, nameCCFile, OpenMode.Output)

            CCCstkFileID = FreeFile()
            nameCCFile = saveFilePrefix & "_stk_CCC.csv"
            FileOpen(CCCstkFileID, nameCCFile, OpenMode.Output)
        End If

        str7 = New String(" ", 7)
        str8 = New String(" ", 8)
        str9 = New String(" ", 9)
        str10 = New String(" ", 10)

        For stk = 1 To numStocks
            '.....ccc_stock header record
            If Stm = 4 Then
                If yr = 0 And stk = 1 Then
                    WriteLine(CCCstkFileID, "year", "stock", "age-2 AEQ", "age-2 cohort", "age-3 AEQ", "age-3 cohort", "age-4 AEQ", "age-4 cohort", "age-5 AEQ", "age-5 cohort", "age-2 term run", "age-2 escape", "age-3 term run", "age-3 escape", "age-4 term run", "age-4 escape", "age-5 term run", "age-5 escape")
                End If

                'begin ccc_stock record, note new ccc_stock file has 4 digit year
                Write(CCCstkFileID, yr + 1979)

                '..... stock
                'write(CCCstkFileID, stockAbbreviation(stk) )
                Write(CCCstkFileID, stk)

                For age = 2 To NumAges
                    '..... AEQ
                    str7 = RSet(VB6.Format(AdltEqv(age, stk), "0.####0"), Len(str7))
                    Write(CCCstkFileID, Val(str7))

                    '..... cohort
                    'oceanCohort(stk, age) = cohort_(age, stk) / survivalRate(age)
                    oceanCohort(stk, age) = saveCohort(age, stk) / survivalRate(age)
                    str8 = RSet(VB6.Format(oceanCohort(stk, age), "#######0"), Len(str8))
                    Write(CCCstkFileID, Val(str8))
                Next age
                Write(CCCstkFileID, " ")
            End If

            For age = 2 To NumAges
                If Stm = 4 Then
                    '..... save true terminal run
                    str10 = RSet(VB6.Format(EscapeT(age, stk), "######0.#0 "), Len(str10))
                    Write(CCCstkFileID, Val(str10))

                    '..... save spawning escapement
                    str10 = RSet(VB6.Format(Escape(age, stk) * IDL(stk, yr), "######0.#0"), Len(str10))
                    Write(CCCstkFileID, Val(str10))
                End If
            Next age

            If Stm = 4 Then
                '..... end ccc_stock record
                WriteLine(CCCstkFileID)

                '..... begin ccc_total mortality record
                If yr = 0 And stk = 1 Then
                    WriteLine(CCCfish_ISBMFileID, "year", "stock", "fishery", "age", "preTerm catch", "Term catch", "preterm shakers", "term shakers", "preterm legal dropoffs", "term legal dropoffs", "preterm sublegal dropoffs", "term sublegal dropoffs", "preterm CNRlegal", "term CNRlegal", "preterm CNRlegal dropoffs", "term CNRlegal dropoffs", "preterm CNRsublegal", "term CNRsublegal", "preterm CNRsublegal dropoffs", "term CNRsublegal dropoffs")
                    WriteLine(CCCfish_AABMFileID, "year", "stock", "fishery", "age", "preTerm catch", "Term catch", "preterm shakers", "term shakers", "preterm legal dropoffs", "term legal dropoffs", "preterm sublegal dropoffs", "term sublegal dropoffs", "preterm CNRlegal", "term CNRlegal", "preterm CNRlegal dropoffs", "term CNRlegal dropoffs", "preterm CNRsublegal", "term CNRsublegal", "preterm CNRsublegal dropoffs", "term CNRsublegal dropoffs")
                End If
                For fish = 1 To numFisheries

                    Select Case fish
                        Case 1
                            CSVfileID = CCCfish_AABMFileID
                        Case 2
                            CSVfileID = CCCfish_AABMFileID
                        Case 4
                            CSVfileID = CCCfish_AABMFileID
                        Case 7
                            CSVfileID = CCCfish_AABMFileID
                        Case 18
                            CSVfileID = CCCfish_AABMFileID
                        Case 19
                            CSVfileID = CCCfish_AABMFileID
                        Case 20
                            CSVfileID = CCCfish_AABMFileID
                        Case Else
                            CSVfileID = CCCfish_ISBMFileID
                    End Select

                    For age = 2 To NumAges
                        '..... year
                        Write(CSVfileID, yr + 1979)

                        '..... stock 
                        'Write(CSVfileID, stockAbbreviation(stk))
                        Write(CSVfileID, stk)

                        '..... fishery
                        'Write(CSVfileID, FisheryName(fish))
                        Write(CSVfileID, fish)

                        '..... age
                        Write(CSVfileID, age)

                        '..... perterminal and terminal catch 
                        str9 = RSet(VB6.Format(legalCatch(PreTerm, fish, stk, age), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)
                        str9 = RSet(VB6.Format(legalCatch(Term, fish, stk, age), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)

                        '..... perterminal sublegal size shakers without sublegal or legal size dropoffs see eq 44, 45 and 49
                        'subtract legal size dropoffs
                        shakersWithOutDropOffs = stkShakCat(PreTerm, fish, stk, age) - legalDropOffs(PreTerm, fish, stk, age)
                        'find the proportion in the remaining sublegal size fish that is dropoff
                        subLegalDropoffs(0) = shakersWithOutDropOffs * dropoffRate(fish, yr) / ShakMortRate(fish)
                        'subtract sublegal size dropoffs
                        shakersWithOutDropOffs = shakersWithOutDropOffs - subLegalDropoffs(0)
                        str9 = RSet(VB6.Format(shakersWithOutDropOffs, "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)

                        '..... terminal sublegal size shakers without sublegal or legal size dropoffs see eq 44, 45 and 49
                        'subtract legal size dropoffs
                        shakersWithOutDropOffs = stkShakCat(Term, fish, stk, age) - legalDropOffs(Term, fish, stk, age)
                        'find the proportion in the remaining sublegal size fish that is dropoff
                        subLegalDropoffs(1) = shakersWithOutDropOffs * dropoffRate(fish, yr) / ShakMortRate(fish)
                        'subtract sublegal size dropoffs
                        shakersWithOutDropOffs = shakersWithOutDropOffs - subLegalDropoffs(1)
                        str9 = RSet(VB6.Format(stkShakCat(Term, fish, stk, age) - legalDropOffs(Term, fish, stk, age), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)

                        '..... perterminal and terminal legal size dropoffs see eq 51
                        str9 = RSet(VB6.Format(legalDropOffs(PreTerm, fish, stk, age), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)
                        str9 = RSet(VB6.Format(legalDropOffs(Term, fish, stk, age), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)

                        '..... perterminal and terminal sublegal size dropoffs 
                        str9 = RSet(VB6.Format(subLegalDropoffs(0), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)
                        str9 = RSet(VB6.Format(subLegalDropoffs(1), "#####0.#0"), Len(str9))
                        number = Val(str9)
                        Write(CSVfileID, number)

                        If CNRFlag(yr, fish) = 1 Then
                            CNRIndx = pointerCNR(fish)

                            'split mortality into release mortality and dropoff 
                            'see eq 83
                            CNR_legal_release_MortRate = legalReleaseMortRate(fish, yr) / (legalReleaseMortRate(fish, yr) + dropoffRate(fish, yr))
                            'see eq 44
                            CNR_sublegal_release_MortRate = 1 - (dropoffRate(fish, yr) / ShakMortRate(fish))


                            '..... perterminal CNR legal size release mortality 
                            str9 = RSet(VB6.Format(CNRLegal(PreTerm, CNRIndx, stk, age) * CNR_legal_release_MortRate, "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)

                            '..... terminal CNR legal size release mortality 
                            str9 = RSet(VB6.Format(CNRLegal(Term, CNRIndx, stk, age) * CNR_legal_release_MortRate, "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)

                            '..... perterminal CNR legal size dropoff mortality 
                            str9 = RSet(VB6.Format(CNRLegal(PreTerm, CNRIndx, stk, age) * (1 - CNR_legal_release_MortRate), "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)

                            '..... terminal CNR legal size dropoff mortality 
                            str9 = RSet(VB6.Format(CNRLegal(Term, CNRIndx, stk, age) * (1 - CNR_legal_release_MortRate), "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)

                            '..... perterminal CNR sublegal size release mortality 
                            str9 = RSet(VB6.Format(CNRShakCat(PreTerm, CNRIndx, stk, age) * CNR_sublegal_release_MortRate, "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)

                            '..... terminal CNR sublegal size release mortality 
                            str9 = RSet(VB6.Format(CNRShakCat(Term, CNRIndx, stk, age) * CNR_sublegal_release_MortRate, "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)


                            '..... perterminal CNR sublegal size dropoffs
                            str9 = RSet(VB6.Format(CNRShakCat(PreTerm, CNRIndx, stk, age) * (1 - CNR_sublegal_release_MortRate), "#####0.#0"), Len(str9))
                            number = Val(str9)
                            Write(CSVfileID, number)

                            '..... terminal CNR sublegal size dropoffs
                            str9 = RSet(VB6.Format(CNRShakCat(Term, CNRIndx, stk, age) * (1 - CNR_sublegal_release_MortRate), "#####0.#0"), Len(str9))
                            number = Val(str9)
                            WriteLine(CSVfileID, number)

                        Else 'no CNR fishery (write zeros in CNR columns)
                            WriteLine(CSVfileID, 0, 0, 0, 0, 0, 0, 0, 0)
                        End If
                    Next age
                Next fish
            End If

        Next stk

    End Sub

End Module