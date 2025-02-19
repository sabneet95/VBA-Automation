Option Explicit

'*******************************************************************************
' Module:       Climate Data Import Macro
' Author:       Sabneet Bains
' Description:  Imports and processes climate data for multiple U.S. cities 
'               using web queries. The macro loops through a set of states, 
'               cities, and corresponding URL text files to dynamically create 
'               worksheets, fetch data, and format it for analysis.
'
' Usage:        Run the Climate macro from the VBA editor or assign it to a 
'               button in Excel. Ensure that VBA 7 or higher is installed.
'
' Requirements: Microsoft Excel 2016 or later, VBA 7 or higher.
'
' License:      MIT License
'*******************************************************************************
Sub Climate()
    ' Variable declarations with explicit types for clarity and maintenance.
    Dim States As Variant                  ' Array of state names
    Dim States_abbrev As Variant            ' Array of state abbreviations
    Dim Cities As Variant                   ' 2D array of cities for each state
    Dim URL_text_file As Variant            ' 2D array of URL text file names for each city
    Dim Main_URL As String                  ' Base URL for data retrieval
    Dim i As Long, j As Long                ' Loop counters for states and cities
    Dim totalStates As Long                 ' Total number of states
    Dim totalCities As Long                 ' Total number of cities for current state
    Dim stateAbbrev As String               ' Abbreviation of the current state
    Dim stateName As String                 ' Name of the current state
    Dim cityArray As Variant                ' Array of cities for current state
    Dim urlArray As Variant                 ' Array of URL text files for current state
    Dim currentSheet As Worksheet           ' Worksheet reference for data import
    Dim ConnectionString As String          ' Connection string for the web query
    Dim destRange As Range                  ' Destination range for data import
    Dim sheetName As String                 ' Desired worksheet name

    '===============================================================================
    ' INITIAL SETUP
    '===============================================================================
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '-----------------------------------------------------------------------------
    ' Initialize arrays containing state names, abbreviations, cities, and URL files.
    '-----------------------------------------------------------------------------
    States = Array("Alabama", "Alaska", "Arizona", "Arkansas", "California", _
                   "Colorado", "Connecticut", "Delaware", "District of Columbia", "Florida", "Georgia", _
                   "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", _
                   "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", _
                   "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", _
                   "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", _
                   "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", _
                   "West Virginia", "Wisconsin", "Wyoming", "Puerto Rico")
    
    States_abbrev = Array("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", _
                          "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", _
                          "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", _
                          "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "PR")
    
    ' 2D array: Each element corresponds to an array of cities for that state.
    Cities = Array( _
        Array("Birmingham", "Huntsville", "Mobile", "Montgomery"), _
        Array("Anchorage", "Fairbanks", "Juneau"), _
        Array("Flagstaff", "Phoenix", "Tucson", "Yuma"), _
        Array("Fort Smith", "Little Rock"), _
        Array("Fresno", "Los Angeles", "Sacramento", "San Diego", "San Francisco"), _
        Array("Colorado Springs", "Denver", "Grand Junction", "Pueblo"), _
        Array("Bridgeport", "Hartford", "Springfield"), _
        Array("Wilmington"), _
        Array("Washington DC"), _
        Array("Daytona Beach", "Jacksonville", "Miami Beach", "Orlando", "Tallahassee", "Tampa", "St. Petersburg", "West Palm Beach"), _
        Array("Atlanta", "Columbus", "Macon", "Savannah"), _
        Array("Honolulu"), _
        Array("Boise", "Pocatello"), _
        Array("Chicago", "Peoria", "Rockford", "Springfield"), _
        Array("Evansville", "Fort Wayne", "Indianapolis", "South Bend"), _
        Array("Des Moines", "Sioux City"), _
        Array("Goodland", "Topeka", "Wichita"), _
        Array("Lexington", "Louisville", "Paducah"), _
        Array("Baton Rouge", "Lake Charles", "New Orleans", "Shreveport"), _
        Array("Caribou", "Portland"), _
        Array("Baltimore", "Washington DC"), _
        Array("Boston"), _
        Array("Detroit", "Flint", "Grand Rapids", "Lansing", "Sault Ste. Marie"), _
        Array("Duluth", "Minneapolis", "St. Paul"), _
        Array("Jackson", "Tupelo"), _
        Array("Kansas City", "Springfield", "St. Louis"), _
        Array("Billings", "Great Falls", "Helena"), _
        Array("Lincoln", "North Platte", "Omaha"), _
        Array("Reno", "Las Vegas"), _
        Array("Concord"), _
        Array("Atlantic City", "Newark"), _
        Array("Albuquerque"), _
        Array("Albany", "Buffalo", "New York City", "Rochester", "Syracuse"), _
        Array("Asheville", "Charlotte", "Greensboro", "Raleigh", "Durham"), _
        Array("Bismarck", "Fargo"), _
        Array("Akron", "Canton", "Cincinnati", "Cleveland", "Columbus", "Dayton", "Toledo", "Youngstown"), _
        Array("Oklahoma City", "Tulsa"), _
        Array("Eugene", "Medford", "Portland", "Salem"), _
        Array("Allentown", "Erie", "Harrisburg", "Philadelphia", "Pittsburgh", "Wilkes-Barre"), _
        Array("Providence"), _
        Array("Charleston", "Columbia"), _
        Array("Rapid City", "Sioux Falls"), _
        Array("Chattanooga", "Knoxville", "Memphis", "Nashville"), _
        Array("Abilene", "Amarillo", "Austin", "Brownsville", "Corpus Christi", "Dallas", "Fort Worth", "El Paso", "Houston", "Lubbock", "Midland", "Odessa", "San Angelo", "San Antonio", "Waco", "Wichita Falls"), _
        Array("Salt Lake City"), _
        Array("Burlington"), _
        Array("Norfolk", "Richmond", "Roanoke"), _
        Array("Seattle", "Spokane", "Yakima"), _
        Array("Charleston", "Elkins"), _
        Array("Green Bay", "Madison", "Milwaukee"), _
        Array("Casper", "Cheyenne"), _
        Array("San Juan") _
    )
    
    ' 2D array: Each element corresponds to an array of URL text file names for that state.
    URL_text_file = Array( _
        Array("ALBIRMIN.txt", "ALHUNTSV.txt", "ALMOBILE.txt", "ALMONTGO.txt"), _
        Array("AKANCHOR.txt", "AKFAIRBA.txt", "AKJUNEAU.txt"), _
        Array("AZFLAGST.txt", "AZPHOENI.txt", "AZTUCSON.txt", "AZYUMA.txt"), _
        Array("ARFTSMIT.txt", "ARLIROCK.txt"), _
        Array("CAFRESNO.txt", "CALOSANG.txt", "CASACRAM.txt", "CASANDIE.txt", "CASANFRA.txt"), _
        Array("COCOSPGS.txt", "CODENVER.txt", "COGRNDJU.txt", "COPUEBLO.txt"), _
        Array("CTBRIDGE.txt", "CTHARTFO.txt", "CTSPRING.txt"), _
        Array("DEWILMIN.txt"), _
        Array("MDWASHDC.txt"), _
        Array("FLDAYTNA.txt", "FLJACKSV.txt", "FLMIAMIB.txt", "FLORLAND.txt", "FLTALLAH.txt", "FLTAMPA.txt", "FLSTPETR.txt", "FLWPALMB.txt"), _
        Array("GAATLANT.txt", "GACOLMBS.txt", "GAMACON.txt", "GASAVANN.txt"), _
        Array("HIHONOLU.txt"), _
        Array("IDBOISE.txt", "IDPOCATE.txt"), _
        Array("ILCHICAG.txt", "ILPEORIA.txt", "ILROCKFO.txt", "ILSPRING.txt"), _
        Array("INEVANSV.txt", "INFTWAYN.txt", "ININDIAN.txt", "INSOBEND.txt"), _
        Array("IADESMOI.txt", "IASIOCTY.txt"), _
        Array("KSGOODLA.txt", "KSTOPEKA.txt", "KSWICHIT.txt"), _
        Array("KYLEXING.txt", "KYLOUISV.txt", "KYPADUCA.txt"), _
        Array("LABATONR.txt", "LALAKECH.txt", "LANEWORL.txt", "LASHREVE.txt"), _
        Array("MECARIBO.txt", "MEPORTLA.txt"), _
        Array("MDBALTIM.txt", "MDWASHDC.txt"), _
        Array("MABOSTON.txt"), _
        Array("MIDETROI.txt", "MIFLINT.txt", "MIGRNDRA.txt", "MILANSIN.txt", "MISTEMAR.txt"), _
        Array("MNDULUTH.txt", "MNMINPLS.txt"), _
        Array("MSJACKSO.txt", "MSTUPELO.txt"), _
        Array("MOKANCTY.txt", "MOSPRING.txt", "MOSTLOUI.txt"), _
        Array("MTBILLIN.txt", "MTGRFALL.txt", "MTHELENA.txt"), _
        Array("NELINCOL.txt", "NENPLATT.txt", "NEOMAHA.txt"), _
        Array("NVRENO.txt", "NVLASVEG.txt"), _
        Array("NHCONCOR.txt"), _
        Array("NJATLCTY.txt", "NJNEWARK.txt"), _
        Array("NMALBUQU.txt"), _
        Array("NYALBANY.txt", "NYBUFFAL.txt", "NYNEWYOR.txt", "NYROCHES.txt", "NYSYRACU.txt"), _
        Array("NCASHEVI.txt", "NCCHARLO.txt", "NCGRNSBO.txt", "NCRALEIG.txt"), _
        Array("NDBISMAR.txt", "NDFARGO.txt"), _
        Array("OHAKRON.txt", "OHCINCIN.txt", "OHCLEVEL.txt", "OHCOLMBS.txt", "OHDAYTON.txt", "OHTOLEDO.txt", "OHYOUNGS.txt"), _
        Array("OKOKLCTY.txt", "OKTULSA.txt"), _
        Array("OREUGENE.txt", "ORMEDFOR.txt", "ORPORTLA.txt", "ORSALEM.txt"), _
        Array("PAALLENT.txt", "PAERIE.txt", "PAHARRIS.txt", "PAPHILAD.txt", "PAPITTSB.txt", "PAWILKES.txt"), _
        Array("RIPROVID.txt"), _
        Array("SCCHARLE.txt", "SCCOLMBA.txt"), _
        Array("SDRAPCTY.txt", "SDSIOFAL.txt"), _
        Array("TNCHATTA.txt", "TNKNOXVI.txt", "TNMEMPHI.txt", "TNNASHVI.txt"), _
        Array("TXABILEN.txt", "TXAMARIL.txt", "TXAUSTIN.txt", "TXBROWNS.txt", "TXCORPUS.txt", "TXDALLAS.txt", "TXELPASO.txt", "TXHOUSTO.txt", "TXLUBBOC.txt", "TXMIDLAN.txt", "TXSANANG.txt", "TXSANANT.txt", "TXWACO.txt", "TXWICHFA.txt"), _
        Array("UTSALTLK.txt"), _
        Array("VTBURLIN.txt"), _
        Array("VANORFOL.txt", "VARICHMO.txt", "VAROANOK.txt"), _
        Array("WASEATTL.txt", "WASPOKAN.txt", "WAYAKIMA.txt"), _
        Array("WVCHARLE.txt", "WVELKINS.txt"), _
        Array("WIGREBAY.txt", "WIMADISO.txt", "WIMILWAU.txt"), _
        Array("WYCASPER.txt", "WYCHEYEN.txt"), _
        Array("PRSANJUA.txt") _
    )
    
    ' Base URL for data retrieval
    Main_URL = "http://academic.udayton.edu/kissock/http/Weather/gsod95-current/"
    
    totalStates = UBound(States)
    
    '===============================================================================
    ' MAIN PROCESSING LOOP
    '===============================================================================
    ' Loop through each state and then each corresponding city.
    For i = 0 To totalStates
        stateAbbrev = States_abbrev(i)
        stateName = States(i)
        cityArray = Cities(i)
        urlArray = URL_text_file(i)
        totalCities = UBound(cityArray)
        
        For j = 0 To totalCities
            ' Determine the worksheet to use:
            ' Use the current active sheet for the first dataset; create new sheets thereafter.
            If Not (i = 0 And j = 0) Then
                Set currentSheet = ThisWorkbook.Worksheets.Add(After:= _
                                  ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            Else
                Set currentSheet = ActiveSheet
            End If
            
            ' Generate a valid worksheet name and handle naming errors gracefully.
            sheetName = cityArray(j) & "_" & stateAbbrev
            On Error Resume Next
            currentSheet.Name = sheetName
            If Err.Number <> 0 Then
                currentSheet.Name = "Sheet_" & i & "_" & j
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            '-----------------------------------------------------------------------------
            ' Data Import via Web Query
            '-----------------------------------------------------------------------------
            ' Build the connection string for the QueryTable.
            ConnectionString = "URL;" & Main_URL & urlArray(j)
            Set destRange = currentSheet.Range("A2")
            
            ' Create and refresh the QueryTable to import web data.
            With currentSheet.QueryTables.Add(Connection:=ConnectionString, Destination:=destRange)
                .AdjustColumnWidth = False
                .WebSelectionType = xlEntirePage
                .WebFormatting = xlWebFormattingNone
                .WebPreFormattedTextToColumns = True
                .WebConsecutiveDelimitersAsOne = True
                .WebSingleBlockTextImport = False
                .WebDisableDateRecognition = False
                .WebDisableRedirections = False
                .Refresh BackgroundQuery:=False
            End With
            
            '-----------------------------------------------------------------------------
            ' Data Processing and Formatting
            '-----------------------------------------------------------------------------
            ' Convert text data to columns, delete unnecessary columns, and set header labels.
            With currentSheet
                .Columns("A").TextToColumns Destination:=.Range("A2"), DataType:=xlDelimited, _
                    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                    Semicolon:=False, Comma:=False, Space:=True, Other:=False, _
                    FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
                    TrailingMinusNumbers:=True
                .Columns("A").Delete Shift:=xlToLeft
                
                ' Set header row labels
                .Range("A1").Value = "Month"
                .Range("B1").Value = "Day"
                .Range("C1").Value = "Year"
                .Range("D1").Value = "Average Daily Temperature (" & Chr(176) & "F)"
                .Columns("A:D").AutoFit
            End With
        Next j
    Next i

Cleanup:
    ' Restore Excel settings to default
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    ' Display an error message and perform cleanup actions.
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Climate Macro Error"
    Resume Cleanup
End Sub
