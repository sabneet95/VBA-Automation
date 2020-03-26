Sub Climate()

Dim States As Variant
Dim States_abbrev As Variant
Dim Cities As Variant
Dim URL_text_file As Variant
Dim Dynamic_URL As String

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
    
    Cities = Array(Array("Birmingham", "Huntsville", "Mobile", "Montgomery"), Array("Anchorage", "Fairbanks", "Juneau"), _
    Array("Flagstaff", "Phoenix", "Tucson", "Yuma"), Array("Fort Smith", "Little Rock"), _
    Array("Fresno", "Los Angeles", "Sacramento", "San Diego", "San Francisco"), _
    Array("Colorado Springs", "Denver", "Grand Junction", "Pueblo"), Array("Bridgeport", "Hartford Springfield"), _
    Array("Wilmington"), Array("Washington DC"), Array("Daytona Beach", "Jacksonville", "Miami Beach", "Orlando", "Tallahassee", "Tampa St. Petersburg", "West Palm Beach"), _
    Array("Atlanta", "Columbus", "Macon", "Savannah"), Array("Honolulu"), Array("Boise", "Pocatello"), Array("Chicago", "Peoria", "Rockford", "Springfield"), _
    Array("Evansville", "Fort Wayne", "Indianapolis", "South Bend"), Array("Des Moines", "Sioux City"), Array("Goodland", "Topeka", "Wichita"), _
    Array("Lexington", "Louisville", "Paducah"), Array("Baton Rouge", "Lake Charles", "New Orleans", "Shreveport"), Array("Caribou", "Portland"), _
    Array("Baltimore", "Washington DC"), Array("Boston"), Array("Detroit", "Flint", "Grand Rapids", "Lansing", "SaultSte Marie"), _
    Array("Duluth", "Minneapolis St. Paul"), Array("Jackson", "Tupelo"), Array("Kansas City", "Springfield", "St Louis"), _
    Array("Billings", "Great Falls", "Helena"), Array("Lincoln", "North Platte", "Omaha"), Array("Reno", "Las Vegas"), Array("Concord"), _
    Array("Atlantic City", "Newark"), Array("Albuquerque"), Array("Albany", "Buffalo", "New York City", "Rochester", "Syracuse"), _
    Array("Asheville", "Charlotte", "Greensboro", "Raleigh Durham"), Array("Bismarck", "Fargo"), Array("Akron Canton", "Cincinnati", "Cleveland", "Columbus", "Dayton", "Toledo", "Youngstown"), _
    Array("Oklahoma City", "Tulsa"), Array("Eugene", "Medford", "Portland", "Salem"), Array("Allentown", "Erie", "Harrisburg", "Philadelphia", "Pittsburgh", "Wilkes Barre"), _
    Array("Rhode Island"), Array("Charleston", "Columbia"), Array("Rapid City", "Sioux Falls"), Array("Chattanooga", "Knoxville", "Memphis", "Nashville"), _
    Array("Abilene", "Amarillo", "Austin", "Brownsville", "Corpus Christi", "Dallas Ft Worth", "El Paso", "Houston", "Lubbock", "Midland Odessa", "San Angelo", "San Antonio", "Waco", "Wichita Falls"), _
    Array("Salt Lake City"), Array("Burlington"), Array("Norfolk", "Richmond", "Roanoke"), Array("Seattle", "Spokane", "Yakima"), Array("Charleston", "Elkins"), Array("Green Bay", "Madison", "Milwaukee"), _
    Array("Casper", "Cheyenne"), Array("San Juan Puerto Rico"))
    
    URL_text_file = Array(Array("ALBIRMIN.txt", "ALHUNTSV.txt", "ALMOBILE.txt", "ALMONTGO.txt"), Array("AKANCHOR.txt", "AKFAIRBA.txt", "AKJUNEAU.txt"), Array("AZFLAGST.txt", "AZPHOENI.txt", "AZTUCSON.txt", "AZYUMA.txt"), _
    Array("ARFTSMIT.txt", "ARLIROCK.txt"), Array("CAFRESNO.txt", "CALOSANG.txt", "CASACRAM.txt", "CASANDIE.txt", "CASANFRA.txt"), Array("COCOSPGS.txt", "CODENVER.txt", "COGRNDJU.txt", "COPUEBLO.txt"), _
    Array("CTBRIDGE.txt", "CTHARTFO.txt"), Array("DEWILMIN.txt"), Array("MDWASHDC.txt"), Array("FLDAYTNA.txt", "FLJACKSV.txt", "FLMIAMIB.txt", "FLORLAND.txt", "FLTALLAH.txt", "FLTAMPA.txt", "FLWPALMB.txt"), _
    Array("GAATLANT.txt", "GACOLMBS.txt", "GAMACON.txt", "GASAVANN.txt"), Array("HIHONOLU.txt"), Array("IDBOISE.txt", "IDPOCATE.txt"), Array("ILCHICAG.txt", "ILPEORIA.txt", "ILROCKFO.txt", "ILSPRING.txt"), _
    Array("INEVANSV.txt", "INFTWAYN.txt", "ININDIAN.txt", "INSOBEND.txt"), Array("IADESMOI.txt", "IASIOCTY.txt"), Array("KSGOODLA.txt", "KSTOPEKA.txt", "KSWICHIT.txt"), Array("KYLEXING.txt", "KYLOUISV.txt", "KYPADUCA.txt"), _
    Array("LABATONR.txt", "LALAKECH.txt", "LANEWORL.txt", "LASHREVE.txt"), Array("MECARIBO.txt", "MEPORTLA.txt"), Array("MDBALTIM.txt", "MDWASHDC.txt"), Array("MABOSTON.txt"), Array("MIDETROI.txt", "MIFLINT.txt", "MIGRNDRA.txt", "MILANSIN.txt", "MISTEMAR.txt"), _
    Array("MNDULUTH.txt", "MNMINPLS.txt"), Array("MSJACKSO.txt", "MSTUPELO.txt"), Array("MOKANCTY.txt", "MOSPRING.txt", "MOSTLOUI.txt"), Array("MTBILLIN.txt", "MTGRFALL.txt", "MTHELENA.txt"), Array("NELINCOL.txt", "NENPLATT.txt", "NEOMAHA.txt"), _
    Array("NVRENO.txt", "NVLASVEG.txt"), Array("NHCONCOR.txt"), Array("NJATLCTY.txt", "NJNEWARK.txt"), Array("NMALBUQU.txt"), Array("NYALBANY.txt", "NYBUFFAL.txt", "NYNEWYOR.txt", "NYROCHES.txt", "NYSYRACU.txt"), Array("NCASHEVI.txt", "NCCHARLO.txt", "NCGRNSBO.txt", "NCRALEIG.txt"), _
    Array("NDBISMAR.txt", "NDFARGO.txt"), Array("OHAKRON.txt", "OHCINCIN.txt", "OHCLEVEL.txt", "OHCOLMBS.txt", "OHDAYTON.txt", "OHTOLEDO.txt", "OHYOUNGS.txt"), Array("OKOKLCTY.txt", "OKTULSA.txt"), Array("OREUGENE.txt", "ORMEDFOR.txt", "ORPORTLA.txt", "ORSALEM.txt"), _
    Array("PAALLENT.txt", "PAERIE.txt", "PAHARRIS.txt", "PAPHILAD.txt", "PAPITTSB.txt", "PAWILKES.txt"), Array("RIPROVID.txt"), Array("SCCHARLE.txt", "SCCOLMBA.txt"), Array("SDRAPCTY.txt", "SDSIOFAL.txt"), Array("TNCHATTA.txt", "TNKNOXVI.txt", "TNMEMPHI.txt", "TNNASHVI.txt"), _
    Array("TXABILEN.txt", "TXAMARIL.txt", "TXAUSTIN.txt", "TXBROWNS.txt", "TXCORPUS.txt", "TXDALLAS.txt", "TXELPASO.txt", "TXHOUSTO.txt", "TXLUBBOC.txt", "TXMIDLAN.txt", "TXSANANG.txt", "TXSANANT.txt", "TXWACO.txt", "TXWICHFA.txt"), Array("UTSALTLK.txt"), Array("VTBURLIN.txt"), _
    Array("VANORFOL.txt", "VARICHMO.txt", "VAROANOK.txt"), Array("WASEATTL.txt", "WASPOKAN.txt", "WAYAKIMA.txt"), Array("WVCHARLE.txt", "WVELKINS.txt"), Array("WIGREBAY.txt", "WIMADISO.txt", "WIMILWAU.txt"), Array("WYCASPER.txt", "WYCHEYEN.txt"), Array("PRSANJUA.txt"))

    Main_URL = "http://academic.udayton.edu/kissock/http/Weather/gsod95-current/"

    For i = 0 To UBound(States)
        For j = 0 To UBound(URL_text_file(i))
            ActiveSheet.Name = Cities(i)(j) & "_" & States_abbrev(i)
            Dynamic_URL = "URL;" & Main_URL & URL_text_file(i)(j)
            With ActiveSheet.QueryTables.Add(Connection:=Dynamic_URL, Destination:=Range("$A$2"))
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
            Columns("A:A").TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
                Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
                :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
                TrailingMinusNumbers:=True
            Columns("A:A").Delete Shift:=xlToLeft
            Range("A1") = "Month"
            Range("B1") = "Day"
            Range("C1") = "Year"
            Range("D1") = "Average Daily Temperature (" & Chr(176) & "F)"
            Range("A1").Select
            Columns("A:D").EntireColumn.AutoFit
            If i < UBound(States) Then
                Sheets.Add After:=ActiveSheet
            End If
        Next j
    Next i
End Sub
