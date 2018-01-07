REM Originala makro por LibreOffice da João X.M.Santos joaoxms@uol.com.br

REM  *****  BASIC  *****

sub Main
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

dim I as integer

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim args2(18) as new com.sun.star.beans.PropertyValue
args2(0).Name = "SearchItem.StyleFamily"
args2(0).Value = 2
args2(1).Name = "SearchItem.CellType"
args2(1).Value = 0
args2(2).Name = "SearchItem.RowDirection"
args2(2).Value = true
args2(3).Name = "SearchItem.AllTables"
args2(3).Value = false
args2(4).Name = "SearchItem.Backward"
args2(4).Value = false
args2(5).Name = "SearchItem.Pattern"
args2(5).Value = false
args2(6).Name = "SearchItem.Content"
args2(6).Value = false
args2(7).Name = "SearchItem.AsianOptions"
args2(7).Value = false
args2(8).Name = "SearchItem.AlgorithmType"
args2(8).Value = 0
args2(9).Name = "SearchItem.SearchFlags"
args2(9).Value = 65536
args2(10).Name = "SearchItem.SearchString"
args2(10).Value = "referendumo"
args2(11).Name = "SearchItem.ReplaceString"
args2(11).Value = "plebicito"
args2(12).Name = "SearchItem.Locale"
args2(12).Value = 255
args2(13).Name = "SearchItem.ChangedChars"
args2(13).Value = 2
args2(14).Name = "SearchItem.DeletedChars"
args2(14).Value = 2
args2(15).Name = "SearchItem.InsertedChars"
args2(15).Value = 2
args2(16).Name = "SearchItem.TransliterateFlags"
args2(16).Value = 1280
args2(17).Name = "SearchItem.Command"
args2(17).Value = 3
args2(18).Name = "Quiet"
args2(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2())

rem ----------------------------------------------------------------------
dim args3(18) as new com.sun.star.beans.PropertyValue
args3(0).Name = "SearchItem.StyleFamily"
args3(0).Value = 2
args3(1).Name = "SearchItem.CellType"
args3(1).Value = 0
args3(2).Name = "SearchItem.RowDirection"
args3(2).Value = true
args3(3).Name = "SearchItem.AllTables"
args3(3).Value = false
args3(4).Name = "SearchItem.Backward"
args3(4).Value = false
args3(5).Name = "SearchItem.Pattern"
args3(5).Value = false
args3(6).Name = "SearchItem.Content"
args3(6).Value = false
args3(7).Name = "SearchItem.AsianOptions"
args3(7).Value = false
args3(8).Name = "SearchItem.AlgorithmType"
args3(8).Value = 0
args3(9).Name = "SearchItem.SearchFlags"
args3(9).Value = 65536
args3(10).Name = "SearchItem.SearchString"
args3(10).Value = "Granda Nordeyo-Milito"
args3(11).Name = "SearchItem.ReplaceString"
args3(11).Value = "Granda Nordala Milito"
args3(12).Name = "SearchItem.Locale"
args3(12).Value = 255
args3(13).Name = "SearchItem.ChangedChars"
args3(13).Value = 2
args3(14).Name = "SearchItem.DeletedChars"
args3(14).Value = 2
args3(15).Name = "SearchItem.InsertedChars"
args3(15).Value = 2
args3(16).Name = "SearchItem.TransliterateFlags"
args3(16).Value = 1280
args3(17).Name = "SearchItem.Command"
args3(17).Value = 3
args3(18).Name = "Quiet"
args3(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args3())

rem ----------------------------------------------------------------------
dim args4(18) as new com.sun.star.beans.PropertyValue
args4(0).Name = "SearchItem.StyleFamily"
args4(0).Value = 2
args4(1).Name = "SearchItem.CellType"
args4(1).Value = 0
args4(2).Name = "SearchItem.RowDirection"
args4(2).Value = true
args4(3).Name = "SearchItem.AllTables"
args4(3).Value = false
args4(4).Name = "SearchItem.Backward"
args4(4).Value = false
args4(5).Name = "SearchItem.Pattern"
args4(5).Value = false
args4(6).Name = "SearchItem.Content"
args4(6).Value = false
args4(7).Name = "SearchItem.AsianOptions"
args4(7).Value = false
args4(8).Name = "SearchItem.AlgorithmType"
args4(8).Value = 0
args4(9).Name = "SearchItem.SearchFlags"
args4(9).Value = 65536
args4(10).Name = "SearchItem.SearchString"
args4(10).Value = "mez-valora"
args4(11).Name = "SearchItem.ReplaceString"
args4(11).Value = "mezavalora"
args4(12).Name = "SearchItem.Locale"
args4(12).Value = 255
args4(13).Name = "SearchItem.ChangedChars"
args4(13).Value = 2
args4(14).Name = "SearchItem.DeletedChars"
args4(14).Value = 2
args4(15).Name = "SearchItem.InsertedChars"
args4(15).Value = 2
args4(16).Name = "SearchItem.TransliterateFlags"
args4(16).Value = 1280
args4(17).Name = "SearchItem.Command"
args4(17).Value = 3
args4(18).Name = "Quiet"
args4(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args4())

rem ----------------------------------------------------------------------
dim args5(18) as new com.sun.star.beans.PropertyValue
args5(0).Name = "SearchItem.StyleFamily"
args5(0).Value = 2
args5(1).Name = "SearchItem.CellType"
args5(1).Value = 0
args5(2).Name = "SearchItem.RowDirection"
args5(2).Value = true
args5(3).Name = "SearchItem.AllTables"
args5(3).Value = false
args5(4).Name = "SearchItem.Backward"
args5(4).Value = false
args5(5).Name = "SearchItem.Pattern"
args5(5).Value = false
args5(6).Name = "SearchItem.Content"
args5(6).Value = false
args5(7).Name = "SearchItem.AsianOptions"
args5(7).Value = false
args5(8).Name = "SearchItem.AlgorithmType"
args5(8).Value = 0
args5(9).Name = "SearchItem.SearchFlags"
args5(9).Value = 65536
args5(10).Name = "SearchItem.SearchString"
args5(10).Value = "Ica esas "
args5(11).Name = "SearchItem.ReplaceString"
args5(11).Value = "Yen "
args5(12).Name = "SearchItem.Locale"
args5(12).Value = 255
args5(13).Name = "SearchItem.ChangedChars"
args5(13).Value = 2
args5(14).Name = "SearchItem.DeletedChars"
args5(14).Value = 2
args5(15).Name = "SearchItem.InsertedChars"
args5(15).Value = 2
args5(16).Name = "SearchItem.TransliterateFlags"
args5(16).Value = 1280
args5(17).Name = "SearchItem.Command"
args5(17).Value = 3
args5(18).Name = "Quiet"
args5(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(18) as new com.sun.star.beans.PropertyValue
args6(0).Name = "SearchItem.StyleFamily"
args6(0).Value = 2
args6(1).Name = "SearchItem.CellType"
args6(1).Value = 0
args6(2).Name = "SearchItem.RowDirection"
args6(2).Value = true
args6(3).Name = "SearchItem.AllTables"
args6(3).Value = false
args6(4).Name = "SearchItem.Backward"
args6(4).Value = false
args6(5).Name = "SearchItem.Pattern"
args6(5).Value = false
args6(6).Name = "SearchItem.Content"
args6(6).Value = false
args6(7).Name = "SearchItem.AsianOptions"
args6(7).Value = false
args6(8).Name = "SearchItem.AlgorithmType"
args6(8).Value = 0
args6(9).Name = "SearchItem.SearchFlags"
args6(9).Value = 65536
args6(10).Name = "SearchItem.SearchString"
args6(10).Value = "Sud-Afrika"
args6(11).Name = "SearchItem.ReplaceString"
args6(11).Value = "Sudafrika"
args6(12).Name = "SearchItem.Locale"
args6(12).Value = 255
args6(13).Name = "SearchItem.ChangedChars"
args6(13).Value = 2
args6(14).Name = "SearchItem.DeletedChars"
args6(14).Value = 2
args6(15).Name = "SearchItem.InsertedChars"
args6(15).Value = 2
args6(16).Name = "SearchItem.TransliterateFlags"
args6(16).Value = 1280
args6(17).Name = "SearchItem.Command"
args6(17).Value = 3
args6(18).Name = "Quiet"
args6(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(18) as new com.sun.star.beans.PropertyValue

Dim Origi(6), Substi(6) as string
Origi(0) = "uzita" : Origi(1)  = "facita" : Origi(2)  = "donita" : Origi(3)  = "skribita" : Origi(4)  = "uzata" : Origi(5)  = "donata" : Origi(6)  = "kreita"
Substi(0) ="uzesas": Substi(1) = "facesas": Substi(2) = "donesas": Substi(3) = "skribesas": Substi(4) = "uzesas": Substi(5) = "donesas": Substi(6) = "kreesas"

For I = 0 to 6

args7(0).Name = "SearchItem.StyleFamily"
args7(0).Value = 2
args7(1).Name = "SearchItem.CellType"
args7(1).Value = 0
args7(2).Name = "SearchItem.RowDirection"
args7(2).Value = true
args7(3).Name = "SearchItem.AllTables"
args7(3).Value = false
args7(4).Name = "SearchItem.Backward"
args7(4).Value = false
args7(5).Name = "SearchItem.Pattern"
args7(5).Value = false
args7(6).Name = "SearchItem.Content"
args7(6).Value = false
args7(7).Name = "SearchItem.AsianOptions"
args7(7).Value = false
args7(8).Name = "SearchItem.AlgorithmType"
args7(8).Value = 0
args7(9).Name = "SearchItem.SearchFlags"
args7(9).Value = 65536
args7(10).Name = "SearchItem.SearchString"
args7(10).Value = "esas " + Origi(I)
args7(11).Name = "SearchItem.ReplaceString"
args7(11).Value = Substi(I)
args7(12).Name = "SearchItem.Locale"
args7(12).Value = 255
args7(13).Name = "SearchItem.ChangedChars"
args7(13).Value = 2
args7(14).Name = "SearchItem.DeletedChars"
args7(14).Value = 2
args7(15).Name = "SearchItem.InsertedChars"
args7(15).Value = 2
args7(16).Name = "SearchItem.TransliterateFlags"
args7(16).Value = 1280
args7(17).Name = "SearchItem.Command"
args7(17).Value = 3
args7(18).Name = "Quiet"
args7(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args7())

Next I

rem ----------------------------------------------------------------------
dim args8(18) as new com.sun.star.beans.PropertyValue
args8(0).Name = "SearchItem.StyleFamily"
args8(0).Value = 2
args8(1).Name = "SearchItem.CellType"
args8(1).Value = 0
args8(2).Name = "SearchItem.RowDirection"
args8(2).Value = true
args8(3).Name = "SearchItem.AllTables"
args8(3).Value = false
args8(4).Name = "SearchItem.Backward"
args8(4).Value = false
args8(5).Name = "SearchItem.Pattern"
args8(5).Value = false
args8(6).Name = "SearchItem.Content"
args8(6).Value = false
args8(7).Name = "SearchItem.AsianOptions"
args8(7).Value = false
args8(8).Name = "SearchItem.AlgorithmType"
args8(8).Value = 0
args8(9).Name = "SearchItem.SearchFlags"
args8(9).Value = 65536
args8(10).Name = "SearchItem.SearchString"
args8(10).Value = "autorulo"
args8(11).Name = "SearchItem.ReplaceString"
args8(11).Value = "skriptisto"
args8(12).Name = "SearchItem.Locale"
args8(12).Value = 255
args8(13).Name = "SearchItem.ChangedChars"
args8(13).Value = 2
args8(14).Name = "SearchItem.DeletedChars"
args8(14).Value = 2
args8(15).Name = "SearchItem.InsertedChars"
args8(15).Value = 2
args8(16).Name = "SearchItem.TransliterateFlags"
args8(16).Value = 1280
args8(17).Name = "SearchItem.Command"
args8(17).Value = 3
args8(18).Name = "Quiet"
args8(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args8())

rem ----------------------------------------------------------------------
dim args9(18) as new com.sun.star.beans.PropertyValue
args9(0).Name = "SearchItem.StyleFamily"
args9(0).Value = 2
args9(1).Name = "SearchItem.CellType"
args9(1).Value = 0
args9(2).Name = "SearchItem.RowDirection"
args9(2).Value = true
args9(3).Name = "SearchItem.AllTables"
args9(3).Value = false
args9(4).Name = "SearchItem.Backward"
args9(4).Value = false
args9(5).Name = "SearchItem.Pattern"
args9(5).Value = false
args9(6).Name = "SearchItem.Content"
args9(6).Value = false
args9(7).Name = "SearchItem.AsianOptions"
args9(7).Value = false
args9(8).Name = "SearchItem.AlgorithmType"
args9(8).Value = 0
args9(9).Name = "SearchItem.SearchFlags"
args9(9).Value = 65536
args9(10).Name = "SearchItem.SearchString"
args9(10).Value = "generale"
args9(11).Name = "SearchItem.ReplaceString"
args9(11).Value = "ordinare"
args9(12).Name = "SearchItem.Locale"
args9(12).Value = 255
args9(13).Name = "SearchItem.ChangedChars"
args9(13).Value = 2
args9(14).Name = "SearchItem.DeletedChars"
args9(14).Value = 2
args9(15).Name = "SearchItem.InsertedChars"
args9(15).Value = 2
args9(16).Name = "SearchItem.TransliterateFlags"
args9(16).Value = 1280
args9(17).Name = "SearchItem.Command"
args9(17).Value = 3
args9(18).Name = "Quiet"
args9(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args9())

rem ----------------------------------------------------------------------
dim args10(18) as new com.sun.star.beans.PropertyValue
args10(0).Name = "SearchItem.StyleFamily"
args10(0).Value = 2
args10(1).Name = "SearchItem.CellType"
args10(1).Value = 0
args10(2).Name = "SearchItem.RowDirection"
args10(2).Value = true
args10(3).Name = "SearchItem.AllTables"
args10(3).Value = false
args10(4).Name = "SearchItem.Backward"
args10(4).Value = false
args10(5).Name = "SearchItem.Pattern"
args10(5).Value = false
args10(6).Name = "SearchItem.Content"
args10(6).Value = false
args10(7).Name = "SearchItem.AsianOptions"
args10(7).Value = false
args10(8).Name = "SearchItem.AlgorithmType"
args10(8).Value = 1
args10(9).Name = "SearchItem.SearchFlags"
args10(9).Value = 65536
args10(10).Name = "SearchItem.SearchString"
args10(10).Value = "–"
args10(11).Name = "SearchItem.ReplaceString"
args10(11).Value = "-"
args10(12).Name = "SearchItem.Locale"
args10(12).Value = 255
args10(13).Name = "SearchItem.ChangedChars"
args10(13).Value = 2
args10(14).Name = "SearchItem.DeletedChars"
args10(14).Value = 2
args10(15).Name = "SearchItem.InsertedChars"
args10(15).Value = 2
args10(16).Name = "SearchItem.TransliterateFlags"
args10(16).Value = 1280
args10(17).Name = "SearchItem.Command"
args10(17).Value = 3
args10(18).Name = "Quiet"
args10(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args10())

rem ----------------------------------------------------------------------
dim args11(18) as new com.sun.star.beans.PropertyValue
args11(0).Name = "SearchItem.StyleFamily"
args11(0).Value = 2
args11(1).Name = "SearchItem.CellType"
args11(1).Value = 0
args11(2).Name = "SearchItem.RowDirection"
args11(2).Value = true
args11(3).Name = "SearchItem.AllTables"
args11(3).Value = false
args11(4).Name = "SearchItem.Backward"
args11(4).Value = false
args11(5).Name = "SearchItem.Pattern"
args11(5).Value = false
args11(6).Name = "SearchItem.Content"
args11(6).Value = false
args11(7).Name = "SearchItem.AsianOptions"
args11(7).Value = false
args11(8).Name = "SearchItem.AlgorithmType"
args11(8).Value = 0
args11(9).Name = "SearchItem.SearchFlags"
args11(9).Value = 65536
args11(10).Name = "SearchItem.SearchString"
args11(10).Value = "chefministro di Peru"
args11(11).Name = "SearchItem.ReplaceString"
args11(11).Value = "chefa ministro di Peru"
args11(12).Name = "SearchItem.Locale"
args11(12).Value = 255
args11(13).Name = "SearchItem.ChangedChars"
args11(13).Value = 2
args11(14).Name = "SearchItem.DeletedChars"
args11(14).Value = 2
args11(15).Name = "SearchItem.InsertedChars"
args11(15).Value = 2
args11(16).Name = "SearchItem.TransliterateFlags"
args11(16).Value = 1280
args11(17).Name = "SearchItem.Command"
args11(17).Value = 3
args11(18).Name = "Quiet"
args11(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args11())

rem ----------------------------------------------------------------------
dim args12(18) as new com.sun.star.beans.PropertyValue

Dim Original(6), Subst(6) as string
Original(0)="uzita": Original(1)="facita": Original(2)="donita": Original(3)="skribita": Original(4)="uzata": Original(5)="donata": Original(6)="kreita"
Subst(0) = "uzesis": Subst(1) = "facesis": Subst(2) = "donesis": Subst(3) = "skribesis": Subst(4) = "uzesis": Subst(5) = "donesis": Subst(6) = "kreesis"

For I = 0 to 6

args12(0).Name = "SearchItem.StyleFamily"
args12(0).Value = 2
args12(1).Name = "SearchItem.CellType"
args12(1).Value = 0
args12(2).Name = "SearchItem.RowDirection"
args12(2).Value = true
args12(3).Name = "SearchItem.AllTables"
args12(3).Value = false
args12(4).Name = "SearchItem.Backward"
args12(4).Value = false
args12(5).Name = "SearchItem.Pattern"
args12(5).Value = false
args12(6).Name = "SearchItem.Content"
args12(6).Value = false
args12(7).Name = "SearchItem.AsianOptions"
args12(7).Value = false
args12(8).Name = "SearchItem.AlgorithmType"
args12(8).Value = 0
args12(9).Name = "SearchItem.SearchFlags"
args12(9).Value = 65536
args12(10).Name = "SearchItem.SearchString"
args12(10).Value = "esis "+Original(I)
args12(11).Name = "SearchItem.ReplaceString"
args12(11).Value = Subst(I)
args12(12).Name = "SearchItem.Locale"
args12(12).Value = 255
args12(13).Name = "SearchItem.ChangedChars"
args12(13).Value = 2
args12(14).Name = "SearchItem.DeletedChars"
args12(14).Value = 2
args12(15).Name = "SearchItem.InsertedChars"
args12(15).Value = 2
args12(16).Name = "SearchItem.TransliterateFlags"
args12(16).Value = 1280
args12(17).Name = "SearchItem.Command"
args12(17).Value = 3
args12(18).Name = "Quiet"
args12(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args12())

Next I

rem ----------------------------------------------------------------------
dim args13(18) as new com.sun.star.beans.PropertyValue
args13(0).Name = "SearchItem.StyleFamily"
args13(0).Value = 2
args13(1).Name = "SearchItem.CellType"
args13(1).Value = 0
args13(2).Name = "SearchItem.RowDirection"
args13(2).Value = true
args13(3).Name = "SearchItem.AllTables"
args13(3).Value = false
args13(4).Name = "SearchItem.Backward"
args13(4).Value = false
args13(5).Name = "SearchItem.Pattern"
args13(5).Value = false
args13(6).Name = "SearchItem.Content"
args13(6).Value = false
args13(7).Name = "SearchItem.AsianOptions"
args13(7).Value = false
args13(8).Name = "SearchItem.AlgorithmType"
args13(8).Value = 0
args13(9).Name = "SearchItem.SearchFlags"
args13(9).Value = 65536
args13(10).Name = "SearchItem.SearchString"
args13(10).Value = "la [[Usana"
args13(11).Name = "SearchItem.ReplaceString"
args13(11).Value = "l'[[Usana"
args13(12).Name = "SearchItem.Locale"
args13(12).Value = 255
args13(13).Name = "SearchItem.ChangedChars"
args13(13).Value = 2
args13(14).Name = "SearchItem.DeletedChars"
args13(14).Value = 2
args13(15).Name = "SearchItem.InsertedChars"
args13(15).Value = 2
args13(16).Name = "SearchItem.TransliterateFlags"
args13(16).Value = 1280
args13(17).Name = "SearchItem.Command"
args13(17).Value = 3
args13(18).Name = "Quiet"
args13(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args13())

rem ----------------------------------------------------------------------
dim args14(18) as new com.sun.star.beans.PropertyValue
args14(0).Name = "SearchItem.StyleFamily"
args14(0).Value = 2
args14(1).Name = "SearchItem.CellType"
args14(1).Value = 0
args14(2).Name = "SearchItem.RowDirection"
args14(2).Value = true
args14(3).Name = "SearchItem.AllTables"
args14(3).Value = false
args14(4).Name = "SearchItem.Backward"
args14(4).Value = false
args14(5).Name = "SearchItem.Pattern"
args14(5).Value = false
args14(6).Name = "SearchItem.Content"
args14(6).Value = false
args14(7).Name = "SearchItem.AsianOptions"
args14(7).Value = false
args14(8).Name = "SearchItem.AlgorithmType"
args14(8).Value = 1
args14(9).Name = "SearchItem.SearchFlags"
args14(9).Value = 65536
args14(10).Name = "SearchItem.SearchString"
args14(10).Value = "fiziologisto"
args14(11).Name = "SearchItem.ReplaceString"
args14(11).Value = "fiziologiisto"
args14(12).Name = "SearchItem.Locale"
args14(12).Value = 255
args14(13).Name = "SearchItem.ChangedChars"
args14(13).Value = 2
args14(14).Name = "SearchItem.DeletedChars"
args14(14).Value = 2
args14(15).Name = "SearchItem.InsertedChars"
args14(15).Value = 2
args14(16).Name = "SearchItem.TransliterateFlags"
args14(16).Value = 1280
args14(17).Name = "SearchItem.Command"
args14(17).Value = 3
args14(18).Name = "Quiet"
args14(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args14())

rem ----------------------------------------------------------------------
dim args15(18) as new com.sun.star.beans.PropertyValue
args15(0).Name = "SearchItem.StyleFamily"
args15(0).Value = 2
args15(1).Name = "SearchItem.CellType"
args15(1).Value = 0
args15(2).Name = "SearchItem.RowDirection"
args15(2).Value = true
args15(3).Name = "SearchItem.AllTables"
args15(3).Value = false
args15(4).Name = "SearchItem.Backward"
args15(4).Value = false
args15(5).Name = "SearchItem.Pattern"
args15(5).Value = false
args15(6).Name = "SearchItem.Content"
args15(6).Value = false
args15(7).Name = "SearchItem.AsianOptions"
args15(7).Value = false
args15(8).Name = "SearchItem.AlgorithmType"
args15(8).Value = 0
args15(9).Name = "SearchItem.SearchFlags"
args15(9).Value = 65536
args15(10).Name = "SearchItem.SearchString"
args15(10).Value = "ne-employeso"
args15(11).Name = "SearchItem.ReplaceString"
args15(11).Value = "chomeso"
args15(12).Name = "SearchItem.Locale"
args15(12).Value = 255
args15(13).Name = "SearchItem.ChangedChars"
args15(13).Value = 2
args15(14).Name = "SearchItem.DeletedChars"
args15(14).Value = 2
args15(15).Name = "SearchItem.InsertedChars"
args15(15).Value = 2
args15(16).Name = "SearchItem.TransliterateFlags"
args15(16).Value = 1280
args15(17).Name = "SearchItem.Command"
args15(17).Value = 3
args15(18).Name = "Quiet"
args15(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args15())

rem ----------------------------------------------------------------------
dim args16(18) as new com.sun.star.beans.PropertyValue
args16(0).Name = "SearchItem.StyleFamily"
args16(0).Value = 2
args16(1).Name = "SearchItem.CellType"
args16(1).Value = 0
args16(2).Name = "SearchItem.RowDirection"
args16(2).Value = true
args16(3).Name = "SearchItem.AllTables"
args16(3).Value = false
args16(4).Name = "SearchItem.Backward"
args16(4).Value = false
args16(5).Name = "SearchItem.Pattern"
args16(5).Value = false
args16(6).Name = "SearchItem.Content"
args16(6).Value = false
args16(7).Name = "SearchItem.AsianOptions"
args16(7).Value = false
args16(8).Name = "SearchItem.AlgorithmType"
args16(8).Value = 0
args16(9).Name = "SearchItem.SearchFlags"
args16(9).Value = 65536
args16(10).Name = "SearchItem.SearchString"
args16(10).Value = "Islama Stato"
args16(11).Name = "SearchItem.ReplaceString"
args16(11).Value = "Islamana Stato"
args16(12).Name = "SearchItem.Locale"
args16(12).Value = 255
args16(13).Name = "SearchItem.ChangedChars"
args16(13).Value = 2
args16(14).Name = "SearchItem.DeletedChars"
args16(14).Value = 2
args16(15).Name = "SearchItem.InsertedChars"
args16(15).Value = 2
args16(16).Name = "SearchItem.TransliterateFlags"
args16(16).Value = 1280
args16(17).Name = "SearchItem.Command"
args16(17).Value = 3
args16(18).Name = "Quiet"
args16(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args16())

rem ----------------------------------------------------------------------
dim args17(18) as new com.sun.star.beans.PropertyValue
args17(0).Name = "SearchItem.StyleFamily"
args17(0).Value = 2
args17(1).Name = "SearchItem.CellType"
args17(1).Value = 0
args17(2).Name = "SearchItem.RowDirection"
args17(2).Value = true
args17(3).Name = "SearchItem.AllTables"
args17(3).Value = false
args17(4).Name = "SearchItem.Backward"
args17(4).Value = false
args17(5).Name = "SearchItem.Pattern"
args17(5).Value = false
args17(6).Name = "SearchItem.Content"
args17(6).Value = false
args17(7).Name = "SearchItem.AsianOptions"
args17(7).Value = false
args17(8).Name = "SearchItem.AlgorithmType"
args17(8).Value = 0
args17(9).Name = "SearchItem.SearchFlags"
args17(9).Value = 65536
args17(10).Name = "SearchItem.SearchString"
args17(10).Value = "Inflacio "
args17(11).Name = "SearchItem.ReplaceString"
args17(11).Value = "Inflaciono "
args17(12).Name = "SearchItem.Locale"
args17(12).Value = 255
args17(13).Name = "SearchItem.ChangedChars"
args17(13).Value = 2
args17(14).Name = "SearchItem.DeletedChars"
args17(14).Value = 2
args17(15).Name = "SearchItem.InsertedChars"
args17(15).Value = 2
args17(16).Name = "SearchItem.TransliterateFlags"
args17(16).Value = 1280
args17(17).Name = "SearchItem.Command"
args17(17).Value = 3
args17(18).Name = "Quiet"
args17(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args17())

rem ----------------------------------------------------------------------
dim args18(18) as new com.sun.star.beans.PropertyValue
args18(0).Name = "SearchItem.StyleFamily"
args18(0).Value = 2
args18(1).Name = "SearchItem.CellType"
args18(1).Value = 0
args18(2).Name = "SearchItem.RowDirection"
args18(2).Value = true
args18(3).Name = "SearchItem.AllTables"
args18(3).Value = false
args18(4).Name = "SearchItem.Backward"
args18(4).Value = false
args18(5).Name = "SearchItem.Pattern"
args18(5).Value = false
args18(6).Name = "SearchItem.Content"
args18(6).Value = false
args18(7).Name = "SearchItem.AsianOptions"
args18(7).Value = false
args18(8).Name = "SearchItem.AlgorithmType"
args18(8).Value = 0
args18(9).Name = "SearchItem.SearchFlags"
args18(9).Value = 65536
args18(10).Name = "SearchItem.SearchString"
args18(10).Value = "populaciono"
args18(11).Name = "SearchItem.ReplaceString"
args18(11).Value = "habitantaro"
args18(12).Name = "SearchItem.Locale"
args18(12).Value = 255
args18(13).Name = "SearchItem.ChangedChars"
args18(13).Value = 2
args18(14).Name = "SearchItem.DeletedChars"
args18(14).Value = 2
args18(15).Name = "SearchItem.InsertedChars"
args18(15).Value = 2
args18(16).Name = "SearchItem.TransliterateFlags"
args18(16).Value = 1280
args18(17).Name = "SearchItem.Command"
args18(17).Value = 3
args18(18).Name = "Quiet"
args18(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args18())

rem ----------------------------------------------------------------------
dim args19(18) as new com.sun.star.beans.PropertyValue

dim DECBO1, DECBO2 as string
For I = 1 to 31
DECBO1 = Ltrim(Str$(I))+" di decembro"
DECBO2 = Ltrim(Str$(I))+"ma di decembro"

args19(0).Name = "SearchItem.StyleFamily"
args19(0).Value = 2
args19(1).Name = "SearchItem.CellType"
args19(1).Value = 0
args19(2).Name = "SearchItem.RowDirection"
args19(2).Value = true
args19(3).Name = "SearchItem.AllTables"
args19(3).Value = false
args19(4).Name = "SearchItem.Backward"
args19(4).Value = false
args19(5).Name = "SearchItem.Pattern"
args19(5).Value = false
args19(6).Name = "SearchItem.Content"
args19(6).Value = false
args19(7).Name = "SearchItem.AsianOptions"
args19(7).Value = false
args19(8).Name = "SearchItem.AlgorithmType"
args19(8).Value = 0
args19(9).Name = "SearchItem.SearchFlags"
args19(9).Value = 65536
args19(10).Name = "SearchItem.SearchString"
args19(10).Value = DECBO1
args19(11).Name = "SearchItem.ReplaceString"
args19(11).Value = DECBO2
args19(12).Name = "SearchItem.Locale"
args19(12).Value = 255
args19(13).Name = "SearchItem.ChangedChars"
args19(13).Value = 2
args19(14).Name = "SearchItem.DeletedChars"
args19(14).Value = 2
args19(15).Name = "SearchItem.InsertedChars"
args19(15).Value = 2
args19(16).Name = "SearchItem.TransliterateFlags"
args19(16).Value = 1280
args19(17).Name = "SearchItem.Command"
args19(17).Value = 3
args19(18).Name = "Quiet"
args19(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args19())

Next I

rem ----------------------------------------------------------------------
dim args20(18) as new com.sun.star.beans.PropertyValue
args20(0).Name = "SearchItem.StyleFamily"
args20(0).Value = 2
args20(1).Name = "SearchItem.CellType"
args20(1).Value = 0
args20(2).Name = "SearchItem.RowDirection"
args20(2).Value = true
args20(3).Name = "SearchItem.AllTables"
args20(3).Value = false
args20(4).Name = "SearchItem.Backward"
args20(4).Value = false
args20(5).Name = "SearchItem.Pattern"
args20(5).Value = false
args20(6).Name = "SearchItem.Content"
args20(6).Value = false
args20(7).Name = "SearchItem.AsianOptions"
args20(7).Value = false
args20(8).Name = "SearchItem.AlgorithmType"
args20(8).Value = 0
args20(9).Name = "SearchItem.SearchFlags"
args20(9).Value = 65536
args20(10).Name = "SearchItem.SearchString"
args20(10).Value = "kontrakto"
args20(11).Name = "SearchItem.ReplaceString"
args20(11).Value = "kontrato"
args20(12).Name = "SearchItem.Locale"
args20(12).Value = 255
args20(13).Name = "SearchItem.ChangedChars"
args20(13).Value = 2
args20(14).Name = "SearchItem.DeletedChars"
args20(14).Value = 2
args20(15).Name = "SearchItem.InsertedChars"
args20(15).Value = 2
args20(16).Name = "SearchItem.TransliterateFlags"
args20(16).Value = 1280
args20(17).Name = "SearchItem.Command"
args20(17).Value = 3
args20(18).Name = "Quiet"
args20(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args20())

rem ----------------------------------------------------------------------
dim args21(18) as new com.sun.star.beans.PropertyValue
args21(0).Name = "SearchItem.StyleFamily"
args21(0).Value = 2
args21(1).Name = "SearchItem.CellType"
args21(1).Value = 0
args21(2).Name = "SearchItem.RowDirection"
args21(2).Value = true
args21(3).Name = "SearchItem.AllTables"
args21(3).Value = false
args21(4).Name = "SearchItem.Backward"
args21(4).Value = false
args21(5).Name = "SearchItem.Pattern"
args21(5).Value = false
args21(6).Name = "SearchItem.Content"
args21(6).Value = false
args21(7).Name = "SearchItem.AsianOptions"
args21(7).Value = false
args21(8).Name = "SearchItem.AlgorithmType"
args21(8).Value = 0
args21(9).Name = "SearchItem.SearchFlags"
args21(9).Value = 65536
args21(10).Name = "SearchItem.SearchString"
args21(10).Value = "mixuri"
args21(11).Name = "SearchItem.ReplaceString"
args21(11).Value = "mixi"
args21(12).Name = "SearchItem.Locale"
args21(12).Value = 255
args21(13).Name = "SearchItem.ChangedChars"
args21(13).Value = 2
args21(14).Name = "SearchItem.DeletedChars"
args21(14).Value = 2
args21(15).Name = "SearchItem.InsertedChars"
args21(15).Value = 2
args21(16).Name = "SearchItem.TransliterateFlags"
args21(16).Value = 1280
args21(17).Name = "SearchItem.Command"
args21(17).Value = 3
args21(18).Name = "Quiet"
args21(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args21())

rem ----------------------------------------------------------------------
dim args22(18) as new com.sun.star.beans.PropertyValue
args22(0).Name = "SearchItem.StyleFamily"
args22(0).Value = 2
args22(1).Name = "SearchItem.CellType"
args22(1).Value = 0
args22(2).Name = "SearchItem.RowDirection"
args22(2).Value = true
args22(3).Name = "SearchItem.AllTables"
args22(3).Value = false
args22(4).Name = "SearchItem.Backward"
args22(4).Value = false
args22(5).Name = "SearchItem.Pattern"
args22(5).Value = false
args22(6).Name = "SearchItem.Content"
args22(6).Value = false
args22(7).Name = "SearchItem.AsianOptions"
args22(7).Value = false
args22(8).Name = "SearchItem.AlgorithmType"
args22(8).Value = 0
args22(9).Name = "SearchItem.SearchFlags"
args22(9).Value = 65536
args22(10).Name = "SearchItem.SearchString"
args22(10).Value = " formas "
args22(11).Name = "SearchItem.ReplaceString"
args22(11).Value = " formacas "
args22(12).Name = "SearchItem.Locale"
args22(12).Value = 255
args22(13).Name = "SearchItem.ChangedChars"
args22(13).Value = 2
args22(14).Name = "SearchItem.DeletedChars"
args22(14).Value = 2
args22(15).Name = "SearchItem.InsertedChars"
args22(15).Value = 2
args22(16).Name = "SearchItem.TransliterateFlags"
args22(16).Value = 1280
args22(17).Name = "SearchItem.Command"
args22(17).Value = 3
args22(18).Name = "Quiet"
args22(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args22())

rem ----------------------------------------------------------------------
dim args23(18) as new com.sun.star.beans.PropertyValue
args23(0).Name = "SearchItem.StyleFamily"
args23(0).Value = 2
args23(1).Name = "SearchItem.CellType"
args23(1).Value = 0
args23(2).Name = "SearchItem.RowDirection"
args23(2).Value = true
args23(3).Name = "SearchItem.AllTables"
args23(3).Value = false
args23(4).Name = "SearchItem.Backward"
args23(4).Value = false
args23(5).Name = "SearchItem.Pattern"
args23(5).Value = false
args23(6).Name = "SearchItem.Content"
args23(6).Value = false
args23(7).Name = "SearchItem.AsianOptions"
args23(7).Value = false
args23(8).Name = "SearchItem.AlgorithmType"
args23(8).Value = 0
args23(9).Name = "SearchItem.SearchFlags"
args23(9).Value = 65536
args23(10).Name = "SearchItem.SearchString"
args23(10).Value = "fondisto"
args23(11).Name = "SearchItem.ReplaceString"
args23(11).Value = "fondinto"
args23(12).Name = "SearchItem.Locale"
args23(12).Value = 255
args23(13).Name = "SearchItem.ChangedChars"
args23(13).Value = 2
args23(14).Name = "SearchItem.DeletedChars"
args23(14).Value = 2
args23(15).Name = "SearchItem.InsertedChars"
args23(15).Value = 2
args23(16).Name = "SearchItem.TransliterateFlags"
args23(16).Value = 1280
args23(17).Name = "SearchItem.Command"
args23(17).Value = 3
args23(18).Name = "Quiet"
args23(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args23())

rem ----------------------------------------------------------------------
dim args24(18) as new com.sun.star.beans.PropertyValue
args24(0).Name = "SearchItem.StyleFamily"
args24(0).Value = 2
args24(1).Name = "SearchItem.CellType"
args24(1).Value = 0
args24(2).Name = "SearchItem.RowDirection"
args24(2).Value = true
args24(3).Name = "SearchItem.AllTables"
args24(3).Value = false
args24(4).Name = "SearchItem.Backward"
args24(4).Value = false
args24(5).Name = "SearchItem.Pattern"
args24(5).Value = false
args24(6).Name = "SearchItem.Content"
args24(6).Value = false
args24(7).Name = "SearchItem.AsianOptions"
args24(7).Value = false
args24(8).Name = "SearchItem.AlgorithmType"
args24(8).Value = 0
args24(9).Name = "SearchItem.SearchFlags"
args24(9).Value = 65536
args24(10).Name = "SearchItem.SearchString"
args24(10).Value = "elektrikala"
args24(11).Name = "SearchItem.ReplaceString"
args24(11).Value = "elektrala"
args24(12).Name = "SearchItem.Locale"
args24(12).Value = 255
args24(13).Name = "SearchItem.ChangedChars"
args24(13).Value = 2
args24(14).Name = "SearchItem.DeletedChars"
args24(14).Value = 2
args24(15).Name = "SearchItem.InsertedChars"
args24(15).Value = 2
args24(16).Name = "SearchItem.TransliterateFlags"
args24(16).Value = 1280
args24(17).Name = "SearchItem.Command"
args24(17).Value = 3
args24(18).Name = "Quiet"
args24(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args24())

rem ----------------------------------------------------------------------
dim args25(18) as new com.sun.star.beans.PropertyValue
args25(0).Name = "SearchItem.StyleFamily"
args25(0).Value = 2
args25(1).Name = "SearchItem.CellType"
args25(1).Value = 0
args25(2).Name = "SearchItem.RowDirection"
args25(2).Value = true
args25(3).Name = "SearchItem.AllTables"
args25(3).Value = false
args25(4).Name = "SearchItem.Backward"
args25(4).Value = false
args25(5).Name = "SearchItem.Pattern"
args25(5).Value = false
args25(6).Name = "SearchItem.Content"
args25(6).Value = false
args25(7).Name = "SearchItem.AsianOptions"
args25(7).Value = false
args25(8).Name = "SearchItem.AlgorithmType"
args25(8).Value = 0
args25(9).Name = "SearchItem.SearchFlags"
args25(9).Value = 65536
args25(10).Name = "SearchItem.SearchString"
args25(10).Value = "di lua populo"
args25(11).Name = "SearchItem.ReplaceString"
args25(11).Value = "di lua habitantaro"
args25(12).Name = "SearchItem.Locale"
args25(12).Value = 255
args25(13).Name = "SearchItem.ChangedChars"
args25(13).Value = 2
args25(14).Name = "SearchItem.DeletedChars"
args25(14).Value = 2
args25(15).Name = "SearchItem.InsertedChars"
args25(15).Value = 2
args25(16).Name = "SearchItem.TransliterateFlags"
args25(16).Value = 1280
args25(17).Name = "SearchItem.Command"
args25(17).Value = 3
args25(18).Name = "Quiet"
args25(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args25())

rem ----------------------------------------------------------------------
dim args26(18) as new com.sun.star.beans.PropertyValue
args26(0).Name = "SearchItem.StyleFamily"
args26(0).Value = 2
args26(1).Name = "SearchItem.CellType"
args26(1).Value = 0
args26(2).Name = "SearchItem.RowDirection"
args26(2).Value = true
args26(3).Name = "SearchItem.AllTables"
args26(3).Value = false
args26(4).Name = "SearchItem.Backward"
args26(4).Value = false
args26(5).Name = "SearchItem.Pattern"
args26(5).Value = false
args26(6).Name = "SearchItem.Content"
args26(6).Value = false
args26(7).Name = "SearchItem.AsianOptions"
args26(7).Value = false
args26(8).Name = "SearchItem.AlgorithmType"
args26(8).Value = 0
args26(9).Name = "SearchItem.SearchFlags"
args26(9).Value = 65536
args26(10).Name = "SearchItem.SearchString"
args26(10).Value = "cinemo-direktisto"
args26(11).Name = "SearchItem.ReplaceString"
args26(11).Value = "filmifisto"
args26(12).Name = "SearchItem.Locale"
args26(12).Value = 255
args26(13).Name = "SearchItem.ChangedChars"
args26(13).Value = 2
args26(14).Name = "SearchItem.DeletedChars"
args26(14).Value = 2
args26(15).Name = "SearchItem.InsertedChars"
args26(15).Value = 2
args26(16).Name = "SearchItem.TransliterateFlags"
args26(16).Value = 1280
args26(17).Name = "SearchItem.Command"
args26(17).Value = 3
args26(18).Name = "Quiet"
args26(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args26())

rem ----------------------------------------------------------------------
dim args27(18) as new com.sun.star.beans.PropertyValue
args27(0).Name = "SearchItem.StyleFamily"
args27(0).Value = 2
args27(1).Name = "SearchItem.CellType"
args27(1).Value = 0
args27(2).Name = "SearchItem.RowDirection"
args27(2).Value = true
args27(3).Name = "SearchItem.AllTables"
args27(3).Value = false
args27(4).Name = "SearchItem.Backward"
args27(4).Value = false
args27(5).Name = "SearchItem.Pattern"
args27(5).Value = false
args27(6).Name = "SearchItem.Content"
args27(6).Value = false
args27(7).Name = "SearchItem.AsianOptions"
args27(7).Value = false
args27(8).Name = "SearchItem.AlgorithmType"
args27(8).Value = 0
args27(9).Name = "SearchItem.SearchFlags"
args27(9).Value = 65536
args27(10).Name = "SearchItem.SearchString"
args27(10).Value = "direktisto di cinemo"
args27(11).Name = "SearchItem.ReplaceString"
args27(11).Value = "filmifisto"
args27(12).Name = "SearchItem.Locale"
args27(12).Value = 255
args27(13).Name = "SearchItem.ChangedChars"
args27(13).Value = 2
args27(14).Name = "SearchItem.DeletedChars"
args27(14).Value = 2
args27(15).Name = "SearchItem.InsertedChars"
args27(15).Value = 2
args27(16).Name = "SearchItem.TransliterateFlags"
args27(16).Value = 1280
args27(17).Name = "SearchItem.Command"
args27(17).Value = 3
args27(18).Name = "Quiet"
args27(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args27())

rem ----------------------------------------------------------------------
dim args28(18) as new com.sun.star.beans.PropertyValue
args28(0).Name = "SearchItem.StyleFamily"
args28(0).Value = 2
args28(1).Name = "SearchItem.CellType"
args28(1).Value = 0
args28(2).Name = "SearchItem.RowDirection"
args28(2).Value = true
args28(3).Name = "SearchItem.AllTables"
args28(3).Value = false
args28(4).Name = "SearchItem.Backward"
args28(4).Value = false
args28(5).Name = "SearchItem.Pattern"
args28(5).Value = false
args28(6).Name = "SearchItem.Content"
args28(6).Value = false
args28(7).Name = "SearchItem.AsianOptions"
args28(7).Value = false
args28(8).Name = "SearchItem.AlgorithmType"
args28(8).Value = 0
args28(9).Name = "SearchItem.SearchFlags"
args28(9).Value = 65536
args28(10).Name = "SearchItem.SearchString"
args28(10).Value = "Imajo:"
args28(11).Name = "SearchItem.ReplaceString"
args28(11).Value = "Arkivo:"
args28(12).Name = "SearchItem.Locale"
args28(12).Value = 255
args28(13).Name = "SearchItem.ChangedChars"
args28(13).Value = 2
args28(14).Name = "SearchItem.DeletedChars"
args28(14).Value = 2
args28(15).Name = "SearchItem.InsertedChars"
args28(15).Value = 2
args28(16).Name = "SearchItem.TransliterateFlags"
args28(16).Value = 1280
args28(17).Name = "SearchItem.Command"
args28(17).Value = 3
args28(18).Name = "Quiet"
args28(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args28())

rem ----------------------------------------------------------------------
dim args29(18) as new com.sun.star.beans.PropertyValue
args29(0).Name = "SearchItem.StyleFamily"
args29(0).Value = 2
args29(1).Name = "SearchItem.CellType"
args29(1).Value = 0
args29(2).Name = "SearchItem.RowDirection"
args29(2).Value = true
args29(3).Name = "SearchItem.AllTables"
args29(3).Value = false
args29(4).Name = "SearchItem.Backward"
args29(4).Value = false
args29(5).Name = "SearchItem.Pattern"
args29(5).Value = false
args29(6).Name = "SearchItem.Content"
args29(6).Value = false
args29(7).Name = "SearchItem.AsianOptions"
args29(7).Value = false
args29(8).Name = "SearchItem.AlgorithmType"
args29(8).Value = 0
args29(9).Name = "SearchItem.SearchFlags"
args29(9).Value = 65536
args29(10).Name = "SearchItem.SearchString"
args29(10).Value = "Category:"
args29(11).Name = "SearchItem.ReplaceString"
args29(11).Value = "Kategorio:"
args29(12).Name = "SearchItem.Locale"
args29(12).Value = 255
args29(13).Name = "SearchItem.ChangedChars"
args29(13).Value = 2
args29(14).Name = "SearchItem.DeletedChars"
args29(14).Value = 2
args29(15).Name = "SearchItem.InsertedChars"
args29(15).Value = 2
args29(16).Name = "SearchItem.TransliterateFlags"
args29(16).Value = 1280
args29(17).Name = "SearchItem.Command"
args29(17).Value = 3
args29(18).Name = "Quiet"
args29(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args29())

rem ----------------------------------------------------------------------
dim args31(18) as new com.sun.star.beans.PropertyValue
args31(0).Name = "SearchItem.StyleFamily"
args31(0).Value = 2
args31(1).Name = "SearchItem.CellType"
args31(1).Value = 0
args31(2).Name = "SearchItem.RowDirection"
args31(2).Value = true
args31(3).Name = "SearchItem.AllTables"
args31(3).Value = false
args31(4).Name = "SearchItem.Backward"
args31(4).Value = false
args31(5).Name = "SearchItem.Pattern"
args31(5).Value = false
args31(6).Name = "SearchItem.Content"
args31(6).Value = false
args31(7).Name = "SearchItem.AsianOptions"
args31(7).Value = false
args31(8).Name = "SearchItem.AlgorithmType"
args31(8).Value = 0
args31(9).Name = "SearchItem.SearchFlags"
args31(9).Value = 65536
args31(10).Name = "SearchItem.SearchString"
args31(10).Value = "rebeliono"
args31(11).Name = "SearchItem.ReplaceString"
args31(11).Value = "rebeleso"
args31(12).Name = "SearchItem.Locale"
args31(12).Value = 255
args31(13).Name = "SearchItem.ChangedChars"
args31(13).Value = 2
args31(14).Name = "SearchItem.DeletedChars"
args31(14).Value = 2
args31(15).Name = "SearchItem.InsertedChars"
args31(15).Value = 2
args31(16).Name = "SearchItem.TransliterateFlags"
args31(16).Value = 1280
args31(17).Name = "SearchItem.Command"
args31(17).Value = 3
args31(18).Name = "Quiet"
args31(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args31())

rem ----------------------------------------------------------------------
dim args32(18) as new com.sun.star.beans.PropertyValue
args32(0).Name = "SearchItem.StyleFamily"
args32(0).Value = 2
args32(1).Name = "SearchItem.CellType"
args32(1).Value = 0
args32(2).Name = "SearchItem.RowDirection"
args32(2).Value = true
args32(3).Name = "SearchItem.AllTables"
args32(3).Value = false
args32(4).Name = "SearchItem.Backward"
args32(4).Value = false
args32(5).Name = "SearchItem.Pattern"
args32(5).Value = false
args32(6).Name = "SearchItem.Content"
args32(6).Value = false
args32(7).Name = "SearchItem.AsianOptions"
args32(7).Value = false
args32(8).Name = "SearchItem.AlgorithmType"
args32(8).Value = 0
args32(9).Name = "SearchItem.SearchFlags"
args32(9).Value = 65536
args32(10).Name = "SearchItem.SearchString"
args32(10).Value = "rebelionis"
args32(11).Name = "SearchItem.ReplaceString"
args32(11).Value = "revoltis"
args32(12).Name = "SearchItem.Locale"
args32(12).Value = 255
args32(13).Name = "SearchItem.ChangedChars"
args32(13).Value = 2
args32(14).Name = "SearchItem.DeletedChars"
args32(14).Value = 2
args32(15).Name = "SearchItem.InsertedChars"
args32(15).Value = 2
args32(16).Name = "SearchItem.TransliterateFlags"
args32(16).Value = 1280
args32(17).Name = "SearchItem.Command"
args32(17).Value = 3
args32(18).Name = "Quiet"
args32(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args32())

rem ----------------------------------------------------------------------
dim args33(18) as new com.sun.star.beans.PropertyValue
args33(0).Name = "SearchItem.StyleFamily"
args33(0).Value = 2
args33(1).Name = "SearchItem.CellType"
args33(1).Value = 0
args33(2).Name = "SearchItem.RowDirection"
args33(2).Value = true
args33(3).Name = "SearchItem.AllTables"
args33(3).Value = false
args33(4).Name = "SearchItem.Backward"
args33(4).Value = false
args33(5).Name = "SearchItem.Pattern"
args33(5).Value = false
args33(6).Name = "SearchItem.Content"
args33(6).Value = false
args33(7).Name = "SearchItem.AsianOptions"
args33(7).Value = false
args33(8).Name = "SearchItem.AlgorithmType"
args33(8).Value = 0
args33(9).Name = "SearchItem.SearchFlags"
args33(9).Value = 65536
args33(10).Name = "SearchItem.SearchString"
args33(10).Value = "Unionita Povi"
args33(11).Name = "SearchItem.ReplaceString"
args33(11).Value = "[[Federiti de la duesma mondomilito|la Federiti]]"
args33(12).Name = "SearchItem.Locale"
args33(12).Value = 255
args33(13).Name = "SearchItem.ChangedChars"
args33(13).Value = 2
args33(14).Name = "SearchItem.DeletedChars"
args33(14).Value = 2
args33(15).Name = "SearchItem.InsertedChars"
args33(15).Value = 2
args33(16).Name = "SearchItem.TransliterateFlags"
args33(16).Value = 1280
args33(17).Name = "SearchItem.Command"
args33(17).Value = 3
args33(18).Name = "Quiet"
args33(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args33())

rem ----------------------------------------------------------------------
dim args34(18) as new com.sun.star.beans.PropertyValue
args34(0).Name = "SearchItem.StyleFamily"
args34(0).Value = 2
args34(1).Name = "SearchItem.CellType"
args34(1).Value = 0
args34(2).Name = "SearchItem.RowDirection"
args34(2).Value = true
args34(3).Name = "SearchItem.AllTables"
args34(3).Value = false
args34(4).Name = "SearchItem.Backward"
args34(4).Value = false
args34(5).Name = "SearchItem.Pattern"
args34(5).Value = false
args34(6).Name = "SearchItem.Content"
args34(6).Value = false
args34(7).Name = "SearchItem.AsianOptions"
args34(7).Value = false
args34(8).Name = "SearchItem.AlgorithmType"
args34(8).Value = 0
args34(9).Name = "SearchItem.SearchFlags"
args34(9).Value = 65536
args34(10).Name = "SearchItem.SearchString"
args34(10).Value = "chefa ministri"
args34(11).Name = "SearchItem.ReplaceString"
args34(11).Value = "chefministri"
args34(12).Name = "SearchItem.Locale"
args34(12).Value = 255
args34(13).Name = "SearchItem.ChangedChars"
args34(13).Value = 2
args34(14).Name = "SearchItem.DeletedChars"
args34(14).Value = 2
args34(15).Name = "SearchItem.InsertedChars"
args34(15).Value = 2
args34(16).Name = "SearchItem.TransliterateFlags"
args34(16).Value = 1280
args34(17).Name = "SearchItem.Command"
args34(17).Value = 3
args34(18).Name = "Quiet"
args34(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args34())

rem ----------------------------------------------------------------------
dim args35(18) as new com.sun.star.beans.PropertyValue
args35(0).Name = "SearchItem.StyleFamily"
args35(0).Value = 2
args35(1).Name = "SearchItem.CellType"
args35(1).Value = 0
args35(2).Name = "SearchItem.RowDirection"
args35(2).Value = true
args35(3).Name = "SearchItem.AllTables"
args35(3).Value = false
args35(4).Name = "SearchItem.Backward"
args35(4).Value = false
args35(5).Name = "SearchItem.Pattern"
args35(5).Value = false
args35(6).Name = "SearchItem.Content"
args35(6).Value = false
args35(7).Name = "SearchItem.AsianOptions"
args35(7).Value = false
args35(8).Name = "SearchItem.AlgorithmType"
args35(8).Value = 0
args35(9).Name = "SearchItem.SearchFlags"
args35(9).Value = 65536
args35(10).Name = "SearchItem.SearchString"
args35(10).Value = "Medicino e Fiziologio"
args35(11).Name = "SearchItem.ReplaceString"
args35(11).Value = "Fiziologio o Medicino"
args35(12).Name = "SearchItem.Locale"
args35(12).Value = 255
args35(13).Name = "SearchItem.ChangedChars"
args35(13).Value = 2
args35(14).Name = "SearchItem.DeletedChars"
args35(14).Value = 2
args35(15).Name = "SearchItem.InsertedChars"
args35(15).Value = 2
args35(16).Name = "SearchItem.TransliterateFlags"
args35(16).Value = 1280
args35(17).Name = "SearchItem.Command"
args35(17).Value = 3
args35(18).Name = "Quiet"
args35(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args35())

rem ----------------------------------------------------------------------
dim args36(18) as new com.sun.star.beans.PropertyValue
args36(0).Name = "SearchItem.StyleFamily"
args36(0).Value = 2
args36(1).Name = "SearchItem.CellType"
args36(1).Value = 0
args36(2).Name = "SearchItem.RowDirection"
args36(2).Value = true
args36(3).Name = "SearchItem.AllTables"
args36(3).Value = false
args36(4).Name = "SearchItem.Backward"
args36(4).Value = false
args36(5).Name = "SearchItem.Pattern"
args36(5).Value = false
args36(6).Name = "SearchItem.Content"
args36(6).Value = false
args36(7).Name = "SearchItem.AsianOptions"
args36(7).Value = false
args36(8).Name = "SearchItem.AlgorithmType"
args36(8).Value = 0
args36(9).Name = "SearchItem.SearchFlags"
args36(9).Value = 65536
args36(10).Name = "SearchItem.SearchString"
args36(10).Value = "ico esas"
args36(11).Name = "SearchItem.ReplaceString"
args36(11).Value = "yen"
args36(12).Name = "SearchItem.Locale"
args36(12).Value = 255
args36(13).Name = "SearchItem.ChangedChars"
args36(13).Value = 2
args36(14).Name = "SearchItem.DeletedChars"
args36(14).Value = 2
args36(15).Name = "SearchItem.InsertedChars"
args36(15).Value = 2
args36(16).Name = "SearchItem.TransliterateFlags"
args36(16).Value = 1280
args36(17).Name = "SearchItem.Command"
args36(17).Value = 3
args36(18).Name = "Quiet"
args36(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args36()
rem ----------------------------------------------------------------------
dim args37(18) as new com.sun.star.beans.PropertyValue
args37(0).Name = "SearchItem.StyleFamily"
args37(0).Value = 2
args37(1).Name = "SearchItem.CellType"
args37(1).Value = 0
args37(2).Name = "SearchItem.RowDirection"
args37(2).Value = true
args37(3).Name = "SearchItem.AllTables"
args37(3).Value = false
args37(4).Name = "SearchItem.Backward"
args37(4).Value = false
args37(5).Name = "SearchItem.Pattern"
args37(5).Value = false
args37(6).Name = "SearchItem.Content"
args37(6).Value = false
args37(7).Name = "SearchItem.AsianOptions"
args37(7).Value = false
args37(8).Name = "SearchItem.AlgorithmType"
args37(8).Value = 0
args37(9).Name = "SearchItem.SearchFlags"
args37(9).Value = 65536
args37(10).Name = "SearchItem.SearchString"
args37(10).Value = "=2 |[[Arkivo:"
args37(11).Name = "SearchItem.ReplaceString"
args37(11).Value = "| Imajo = [[Arkivo:"
args37(12).Name = "SearchItem.Locale"
args37(12).Value = 255
args37(13).Name = "SearchItem.ChangedChars"
args37(13).Value = 2
args37(14).Name = "SearchItem.DeletedChars"
args37(14).Value = 2
args37(15).Name = "SearchItem.InsertedChars"
args37(15).Value = 2
args37(16).Name = "SearchItem.TransliterateFlags"
args37(16).Value = 1280
args37(17).Name = "SearchItem.Command"
args37(17).Value = 3
args37(18).Name = "Quiet"
args37(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args37()
rem ----------------------------------------------------------------------
dim args38(18) as new com.sun.star.beans.PropertyValue
args38(0).Name = "SearchItem.StyleFamily"
args38(0).Value = 2
args38(1).Name = "SearchItem.CellType"
args38(1).Value = 0
args38(2).Name = "SearchItem.RowDirection"
args38(2).Value = true
args38(3).Name = "SearchItem.AllTables"
args38(3).Value = false
args38(4).Name = "SearchItem.Backward"
args38(4).Value = false
args38(5).Name = "SearchItem.Pattern"
args38(5).Value = false
args38(6).Name = "SearchItem.Content"
args38(6).Value = false
args38(7).Name = "SearchItem.AsianOptions"
args38(7).Value = false
args38(8).Name = "SearchItem.AlgorithmType"
args38(8).Value = 0
args38(9).Name = "SearchItem.SearchFlags"
args38(9).Value = 65536
args38(10).Name = "SearchItem.SearchString"
args38(10).Value = "!align=right|Guvernisteso:"
args38(11).Name = "SearchItem.ReplaceString"
args38(11).Value = "| Guvernisteso = "
args38(12).Name = "SearchItem.Locale"
args38(12).Value = 255
args38(13).Name = "SearchItem.ChangedChars"
args38(13).Value = 2
args38(14).Name = "SearchItem.DeletedChars"
args38(14).Value = 2
args38(15).Name = "SearchItem.InsertedChars"
args38(15).Value = 2
args38(16).Name = "SearchItem.TransliterateFlags"
args38(16).Value = 1280
args38(17).Name = "SearchItem.Command"
args38(17).Value = 3
args38(18).Name = "Quiet"
args38(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args38()
rem ----------------------------------------------------------------------
dim args39(18) as new com.sun.star.beans.PropertyValue
args39(0).Name = "SearchItem.StyleFamily"
args39(0).Value = 2
args39(1).Name = "SearchItem.CellType"
args39(1).Value = 0
args39(2).Name = "SearchItem.RowDirection"
args39(2).Value = true
args39(3).Name = "SearchItem.AllTables"
args39(3).Value = false
args39(4).Name = "SearchItem.Backward"
args39(4).Value = false
args39(5).Name = "SearchItem.Pattern"
args39(5).Value = false
args39(6).Name = "SearchItem.Content"
args39(6).Value = false
args39(7).Name = "SearchItem.AsianOptions"
args39(7).Value = false
args39(8).Name = "SearchItem.AlgorithmType"
args39(8).Value = 0
args39(9).Name = "SearchItem.SearchFlags"
args39(9).Value = 65536
args39(10).Name = "SearchItem.SearchString"
args39(10).Value = "!align=right|Precedanto:"
args39(11).Name = "SearchItem.ReplaceString"
args39(11).Value = "| Precedanto1 = "
args39(12).Name = "SearchItem.Locale"
args39(12).Value = 255
args39(13).Name = "SearchItem.ChangedChars"
args39(13).Value = 2
args39(14).Name = "SearchItem.DeletedChars"
args39(14).Value = 2
args39(15).Name = "SearchItem.InsertedChars"
args39(15).Value = 2
args39(16).Name = "SearchItem.TransliterateFlags"
args39(16).Value = 1280
args39(17).Name = "SearchItem.Command"
args39(17).Value = 3
args39(18).Name = "Quiet"
args39(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args39()
rem ----------------------------------------------------------------------
dim args40(18) as new com.sun.star.beans.PropertyValue
args40(0).Name = "SearchItem.StyleFamily"
args40(0).Value = 2
args40(1).Name = "SearchItem.CellType"
args40(1).Value = 0
args40(2).Name = "SearchItem.RowDirection"
args40(2).Value = true
args40(3).Name = "SearchItem.AllTables"
args40(3).Value = false
args40(4).Name = "SearchItem.Backward"
args40(4).Value = false
args40(5).Name = "SearchItem.Pattern"
args40(5).Value = false
args40(6).Name = "SearchItem.Content"
args40(6).Value = false
args40(7).Name = "SearchItem.AsianOptions"
args40(7).Value = false
args40(8).Name = "SearchItem.AlgorithmType"
args40(8).Value = 0
args40(9).Name = "SearchItem.SearchFlags"
args40(9).Value = 65536
args40(10).Name = "SearchItem.SearchString"
args40(10).Value = "asistis"
args40(11).Name = "SearchItem.ReplaceString"
args40(11).Value = "frequentis"
args40(12).Name = "SearchItem.Locale"
args40(12).Value = 255
args40(13).Name = "SearchItem.ChangedChars"
args40(13).Value = 2
args40(14).Name = "SearchItem.DeletedChars"
args40(14).Value = 2
args40(15).Name = "SearchItem.InsertedChars"
args40(15).Value = 2
args40(16).Name = "SearchItem.TransliterateFlags"
args40(16).Value = 1280
args40(17).Name = "SearchItem.Command"
args40(17).Value = 3
args40(18).Name = "Quiet"
args40(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args40()

rem ----------------------------------------------------------------------
dim args41(18) as new com.sun.star.beans.PropertyValue
args41(0).Name = "SearchItem.StyleFamily"
args41(0).Value = 2
args41(1).Name = "SearchItem.CellType"
args41(1).Value = 0
args41(2).Name = "SearchItem.RowDirection"
args41(2).Value = true
args41(3).Name = "SearchItem.AllTables"
args41(3).Value = false
args41(4).Name = "SearchItem.Backward"
args41(4).Value = false
args41(5).Name = "SearchItem.Pattern"
args41(5).Value = false
args41(6).Name = "SearchItem.Content"
args41(6).Value = false
args41(7).Name = "SearchItem.AsianOptions"
args41(7).Value = false
args41(8).Name = "SearchItem.AlgorithmType"
args41(8).Value = 0
args41(9).Name = "SearchItem.SearchFlags"
args41(9).Value = 65536
args41(10).Name = "SearchItem.SearchString"
args41(10).Value = "!align=right|Sucedanto:"
args41(11).Name = "SearchItem.ReplaceString"
args41(11).Value = "| Sucedanto = "
args41(12).Name = "SearchItem.Locale"
args41(12).Value = 255
args41(13).Name = "SearchItem.ChangedChars"
args41(13).Value = 2
args41(14).Name = "SearchItem.DeletedChars"
args41(14).Value = 2
args41(15).Name = "SearchItem.InsertedChars"
args41(15).Value = 2
args41(16).Name = "SearchItem.TransliterateFlags"
args41(16).Value = 1280
args41(17).Name = "SearchItem.Command"
args41(17).Value = 3
args41(18).Name = "Quiet"
args41(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args41()

rem ----------------------------------------------------------------------
dim args42(18) as new com.sun.star.beans.PropertyValue
args42(0).Name = "SearchItem.StyleFamily"
args42(0).Value = 2
args42(1).Name = "SearchItem.CellType"
args42(1).Value = 0
args42(2).Name = "SearchItem.RowDirection"
args42(2).Value = true
args42(3).Name = "SearchItem.AllTables"
args42(3).Value = false
args42(4).Name = "SearchItem.Backward"
args42(4).Value = false
args42(5).Name = "SearchItem.Pattern"
args42(5).Value = false
args42(6).Name = "SearchItem.Content"
args42(6).Value = false
args42(7).Name = "SearchItem.AsianOptions"
args42(7).Value = false
args42(8).Name = "SearchItem.AlgorithmType"
args42(8).Value = 0
args42(9).Name = "SearchItem.SearchFlags"
args42(9).Value = 65536
args42(10).Name = "SearchItem.SearchString"
args42(10).Value = "{Ciencisto|"
args42(11).Name = "SearchItem.ReplaceString"
args42(11).Value = "{Biografio"
args42(12).Name = "SearchItem.Locale"
args42(12).Value = 255
args42(13).Name = "SearchItem.ChangedChars"
args42(13).Value = 2
args42(14).Name = "SearchItem.DeletedChars"
args42(14).Value = 2
args42(15).Name = "SearchItem.InsertedChars"
args42(15).Value = 2
args42(16).Name = "SearchItem.TransliterateFlags"
args42(16).Value = 1280
args42(17).Name = "SearchItem.Command"
args42(17).Value = 3
args42(18).Name = "Quiet"
args42(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args42()

rem ----------------------------------------------------------------------
dim args43(18) as new com.sun.star.beans.PropertyValue
args43(0).Name = "SearchItem.StyleFamily"
args43(0).Value = 2
args43(1).Name = "SearchItem.CellType"
args43(1).Value = 0
args43(2).Name = "SearchItem.RowDirection"
args43(2).Value = true
args43(3).Name = "SearchItem.AllTables"
args43(3).Value = false
args43(4).Name = "SearchItem.Backward"
args43(4).Value = false
args43(5).Name = "SearchItem.Pattern"
args43(5).Value = false
args43(6).Name = "SearchItem.Content"
args43(6).Value = false
args43(7).Name = "SearchItem.AsianOptions"
args43(7).Value = false
args43(8).Name = "SearchItem.AlgorithmType"
args43(8).Value = 0
args43(9).Name = "SearchItem.SearchFlags"
args43(9).Value = 65536
args43(10).Name = "SearchItem.SearchString"
args43(10).Value = "autorino"
args43(11).Name = "SearchItem.ReplaceString"
args43(11).Value = "skriptisto"
args43(12).Name = "SearchItem.Locale"
args43(12).Value = 255
args43(13).Name = "SearchItem.ChangedChars"
args43(13).Value = 2
args43(14).Name = "SearchItem.DeletedChars"
args43(14).Value = 2
args43(15).Name = "SearchItem.InsertedChars"
args43(15).Value = 2
args43(16).Name = "SearchItem.TransliterateFlags"
args43(16).Value = 1280
args43(17).Name = "SearchItem.Command"
args43(17).Value = 3
args43(18).Name = "Quiet"
args43(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args43()

rem ----------------------------------------------------------------------
dim args44(18) as new com.sun.star.beans.PropertyValue
args44(0).Name = "SearchItem.StyleFamily"
args44(0).Value = 2
args44(1).Name = "SearchItem.CellType"
args44(1).Value = 0
args44(2).Name = "SearchItem.RowDirection"
args44(2).Value = true
args44(3).Name = "SearchItem.AllTables"
args44(3).Value = false
args44(4).Name = "SearchItem.Backward"
args44(4).Value = false
args44(5).Name = "SearchItem.Pattern"
args44(5).Value = false
args44(6).Name = "SearchItem.Content"
args44(6).Value = false
args44(7).Name = "SearchItem.AsianOptions"
args44(7).Value = false
args44(8).Name = "SearchItem.AlgorithmType"
args44(8).Value = 0
args44(9).Name = "SearchItem.SearchFlags"
args44(9).Value = 65536
args44(10).Name = "SearchItem.SearchString"
args44(10).Value = "Formulo 1ma"
args44(11).Name = "SearchItem.ReplaceString"
args44(11).Value = "Formulo 1"
args44(12).Name = "SearchItem.Locale"
args44(12).Value = 255
args44(13).Name = "SearchItem.ChangedChars"
args44(13).Value = 2
args44(14).Name = "SearchItem.DeletedChars"
args44(14).Value = 2
args44(15).Name = "SearchItem.InsertedChars"
args44(15).Value = 2
args44(16).Name = "SearchItem.TransliterateFlags"
args44(16).Value = 1280
args44(17).Name = "SearchItem.Command"
args44(17).Value = 3
args44(18).Name = "Quiet"
args44(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args44()

rem ----------------------------------------------------------------------
dim args45(18) as new com.sun.star.beans.PropertyValue
args45(0).Name = "SearchItem.StyleFamily"
args45(0).Value = 2
args45(1).Name = "SearchItem.CellType"
args45(1).Value = 0
args45(2).Name = "SearchItem.RowDirection"
args45(2).Value = true
args45(3).Name = "SearchItem.AllTables"
args45(3).Value = false
args45(4).Name = "SearchItem.Backward"
args45(4).Value = false
args45(5).Name = "SearchItem.Pattern"
args45(5).Value = false
args45(6).Name = "SearchItem.Content"
args45(6).Value = false
args45(7).Name = "SearchItem.AsianOptions"
args45(7).Value = false
args45(8).Name = "SearchItem.AlgorithmType"
args45(8).Value = 0
args45(9).Name = "SearchItem.SearchFlags"
args45(9).Value = 65536
args45(10).Name = "SearchItem.SearchString"
args45(10).Value = "autoro"
args45(11).Name = "SearchItem.ReplaceString"
args45(11).Value = "skriptisto"
args45(12).Name = "SearchItem.Locale"
args45(12).Value = 255
args45(13).Name = "SearchItem.ChangedChars"
args45(13).Value = 2
args45(14).Name = "SearchItem.DeletedChars"
args45(14).Value = 2
args45(15).Name = "SearchItem.InsertedChars"
args45(15).Value = 2
args45(16).Name = "SearchItem.TransliterateFlags"
args45(16).Value = 1280
args45(17).Name = "SearchItem.Command"
args45(17).Value = 3
args45(18).Name = "Quiet"
args45(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args45()

rem ----------------------------------------------------------------------
dim args46(18) as new com.sun.star.beans.PropertyValue
args46(0).Name = "SearchItem.StyleFamily"
args46(0).Value = 2
args46(1).Name = "SearchItem.CellType"
args46(1).Value = 0
args46(2).Name = "SearchItem.RowDirection"
args46(2).Value = true
args46(3).Name = "SearchItem.AllTables"
args46(3).Value = false
args46(4).Name = "SearchItem.Backward"
args46(4).Value = false
args46(5).Name = "SearchItem.Pattern"
args46(5).Value = false
args46(6).Name = "SearchItem.Content"
args46(6).Value = false
args46(7).Name = "SearchItem.AsianOptions"
args46(7).Value = false
args46(8).Name = "SearchItem.AlgorithmType"
args46(8).Value = 0
args46(9).Name = "SearchItem.SearchFlags"
args46(9).Value = 65536
args46(10).Name = "SearchItem.SearchString"
args46(10).Value = "Gregoriana"
args46(11).Name = "SearchItem.ReplaceString"
args46(11).Value = "Gregoriala"
args46(12).Name = "SearchItem.Locale"
args46(12).Value = 255
args46(13).Name = "SearchItem.ChangedChars"
args46(13).Value = 2
args46(14).Name = "SearchItem.DeletedChars"
args46(14).Value = 2
args46(15).Name = "SearchItem.InsertedChars"
args46(15).Value = 2
args46(16).Name = "SearchItem.TransliterateFlags"
args46(16).Value = 1280
args46(17).Name = "SearchItem.Command"
args46(17).Value = 3
args46(18).Name = "Quiet"
args46(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args46()

rem ----------------------------------------------------------------------
dim args47(18) as new com.sun.star.beans.PropertyValue
args47(0).Name = "SearchItem.StyleFamily"
args47(0).Value = 2
args47(1).Name = "SearchItem.CellType"
args47(1).Value = 0
args47(2).Name = "SearchItem.RowDirection"
args47(2).Value = true
args47(3).Name = "SearchItem.AllTables"
args47(3).Value = false
args47(4).Name = "SearchItem.Backward"
args47(4).Value = false
args47(5).Name = "SearchItem.Pattern"
args47(5).Value = false
args47(6).Name = "SearchItem.Content"
args47(6).Value = false
args47(7).Name = "SearchItem.AsianOptions"
args47(7).Value = false
args47(8).Name = "SearchItem.AlgorithmType"
args47(8).Value = 0
args47(9).Name = "SearchItem.SearchFlags"
args47(9).Value = 65536
args47(10).Name = "SearchItem.SearchString"
args47(10).Value = "Juliana"
args47(11).Name = "SearchItem.ReplaceString"
args47(11).Value = "Juliala"
args47(12).Name = "SearchItem.Locale"
args47(12).Value = 255
args47(13).Name = "SearchItem.ChangedChars"
args47(13).Value = 2
args47(14).Name = "SearchItem.DeletedChars"
args47(14).Value = 2
args47(15).Name = "SearchItem.InsertedChars"
args47(15).Value = 2
args47(16).Name = "SearchItem.TransliterateFlags"
args47(16).Value = 1280
args47(17).Name = "SearchItem.Command"
args47(17).Value = 3
args47(18).Name = "Quiet"
args47(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args47()

rem ----------------------------------------------------------------------
dim args48(18) as new com.sun.star.beans.PropertyValue
args48(0).Name = "SearchItem.StyleFamily"
args48(0).Value = 2
args48(1).Name = "SearchItem.CellType"
args48(1).Value = 0
args48(2).Name = "SearchItem.RowDirection"
args48(2).Value = true
args48(3).Name = "SearchItem.AllTables"
args48(3).Value = false
args48(4).Name = "SearchItem.Backward"
args48(4).Value = false
args48(5).Name = "SearchItem.Pattern"
args48(5).Value = false
args48(6).Name = "SearchItem.Content"
args48(6).Value = false
args48(7).Name = "SearchItem.AsianOptions"
args48(7).Value = false
args48(8).Name = "SearchItem.AlgorithmType"
args48(8).Value = 0
args48(9).Name = "SearchItem.SearchFlags"
args48(9).Value = 65536
args48(10).Name = "SearchItem.SearchString"
args48(10).Value = "sociologiisto"
args48(11).Name = "SearchItem.ReplaceString"
args48(11).Value = "sociologo"
args48(12).Name = "SearchItem.Locale"
args48(12).Value = 255
args48(13).Name = "SearchItem.ChangedChars"
args48(13).Value = 2
args48(14).Name = "SearchItem.DeletedChars"
args48(14).Value = 2
args48(15).Name = "SearchItem.InsertedChars"
args48(15).Value = 2
args48(16).Name = "SearchItem.TransliterateFlags"
args48(16).Value = 1280
args48(17).Name = "SearchItem.Command"
args48(17).Value = 3
args48(18).Name = "Quiet"
args48(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args48()

rem ----------------------------------------------------------------------
dim args49(18) as new com.sun.star.beans.PropertyValue
args49(0).Name = "SearchItem.StyleFamily"
args49(0).Value = 2
args49(1).Name = "SearchItem.CellType"
args49(1).Value = 0
args49(2).Name = "SearchItem.RowDirection"
args49(2).Value = true
args49(3).Name = "SearchItem.AllTables"
args49(3).Value = false
args49(4).Name = "SearchItem.Backward"
args49(4).Value = false
args49(5).Name = "SearchItem.Pattern"
args49(5).Value = false
args49(6).Name = "SearchItem.Content"
args49(6).Value = false
args49(7).Name = "SearchItem.AsianOptions"
args49(7).Value = false
args49(8).Name = "SearchItem.AlgorithmType"
args49(8).Value = 0
args49(9).Name = "SearchItem.SearchFlags"
args49(9).Value = 65536
args49(10).Name = "SearchItem.SearchString"
args49(10).Value = "kantistulo"
args49(11).Name = "SearchItem.ReplaceString"
args49(11).Value = "kantisto"
args49(12).Name = "SearchItem.Locale"
args49(12).Value = 255
args49(13).Name = "SearchItem.ChangedChars"
args49(13).Value = 2
args49(14).Name = "SearchItem.DeletedChars"
args49(14).Value = 2
args49(15).Name = "SearchItem.InsertedChars"
args49(15).Value = 2
args49(16).Name = "SearchItem.TransliterateFlags"
args49(16).Value = 1280
args49(17).Name = "SearchItem.Command"
args49(17).Value = 3
args49(18).Name = "Quiet"
args49(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args49()

rem ----------------------------------------------------------------------
dim args50(18) as new com.sun.star.beans.PropertyValue
args50(0).Name = "SearchItem.StyleFamily"
args50(0).Value = 2
args50(1).Name = "SearchItem.CellType"
args50(1).Value = 0
args50(2).Name = "SearchItem.RowDirection"
args50(2).Value = true
args50(3).Name = "SearchItem.AllTables"
args50(3).Value = false
args50(4).Name = "SearchItem.Backward"
args50(4).Value = false
args50(5).Name = "SearchItem.Pattern"
args50(5).Value = false
args50(6).Name = "SearchItem.Content"
args50(6).Value = false
args50(7).Name = "SearchItem.AsianOptions"
args50(7).Value = false
args50(8).Name = "SearchItem.AlgorithmType"
args50(8).Value = 0
args50(9).Name = "SearchItem.SearchFlags"
args50(9).Value = 65536
args50(10).Name = "SearchItem.SearchString"
args50(10).Value = "psikiatro"
args50(11).Name = "SearchItem.ReplaceString"
args50(11).Value = "psikiatriisto"
args50(12).Name = "SearchItem.Locale"
args50(12).Value = 255
args50(13).Name = "SearchItem.ChangedChars"
args50(13).Value = 2
args50(14).Name = "SearchItem.DeletedChars"
args50(14).Value = 2
args50(15).Name = "SearchItem.InsertedChars"
args50(15).Value = 2
args50(16).Name = "SearchItem.TransliterateFlags"
args50(16).Value = 1280
args50(17).Name = "SearchItem.Command"
args50(17).Value = 3
args50(18).Name = "Quiet"
args50(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args50()

rem ---------------------------------------------------------------------
dim args51(18) as new com.sun.star.beans.PropertyValue

dim VARIAVEL1, VARIAVEL2 as string
For I = 1000 to 2010 step 10
VARIAVEL1 = Ltrim(Str$(I))+"a yari"
VARIAVEL2 = "yari "+Ltrim(Str$(I))+"ma"

args51(0).Name = "SearchItem.StyleFamily"
args51(0).Value = 2
args51(1).Name = "SearchItem.CellType"
args51(1).Value = 0
args51(2).Name = "SearchItem.RowDirection"
args51(2).Value = true
args51(3).Name = "SearchItem.AllTables"
args51(3).Value = false
args51(4).Name = "SearchItem.Backward"
args51(4).Value = false
args51(5).Name = "SearchItem.Pattern"
args51(5).Value = false
args51(6).Name = "SearchItem.Content"
args51(6).Value = false
args51(7).Name = "SearchItem.AsianOptions"
args51(7).Value = false
args51(8).Name = "SearchItem.AlgorithmType"
args51(8).Value = 0
args51(9).Name = "SearchItem.SearchFlags"
args51(9).Value = 65536
args51(10).Name = "SearchItem.SearchString"
args51(10).Value = VARIAVEL1
args51(11).Name = "SearchItem.ReplaceString"
args51(11).Value = VARIAVEL2
args51(12).Name = "SearchItem.Locale"
args51(12).Value = 255
args51(13).Name = "SearchItem.ChangedChars"
args51(13).Value = 2
args51(14).Name = "SearchItem.DeletedChars"
args51(14).Value = 2
args51(15).Name = "SearchItem.InsertedChars"
args51(15).Value = 2
args51(16).Name = "SearchItem.TransliterateFlags"
args51(16).Value = 1280
args51(17).Name = "SearchItem.Command"
args51(17).Value = 3
args51(18).Name = "Quiet"
args51(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args51())

Next I

rem ----------------------------------------------------------------------
dim args52(18) as new com.sun.star.beans.PropertyValue
args52(0).Name = "SearchItem.StyleFamily"
args52(0).Value = 2
args52(1).Name = "SearchItem.CellType"
args52(1).Value = 0
args52(2).Name = "SearchItem.RowDirection"
args52(2).Value = true
args52(3).Name = "SearchItem.AllTables"
args52(3).Value = false
args52(4).Name = "SearchItem.Backward"
args52(4).Value = false
args52(5).Name = "SearchItem.Pattern"
args52(5).Value = false
args52(6).Name = "SearchItem.Content"
args52(6).Value = false
args52(7).Name = "SearchItem.AsianOptions"
args52(7).Value = false
args52(8).Name = "SearchItem.AlgorithmType"
args52(8).Value = 0
args52(9).Name = "SearchItem.SearchFlags"
args52(9).Value = 65536
args52(10).Name = "SearchItem.SearchString"
args52(10).Value = "Islama]"
args52(11).Name = "SearchItem.ReplaceString"
args52(11).Value = "Islamana]"
args52(12).Name = "SearchItem.Locale"
args52(12).Value = 255
args52(13).Name = "SearchItem.ChangedChars"
args52(13).Value = 2
args52(14).Name = "SearchItem.DeletedChars"
args52(14).Value = 2
args52(15).Name = "SearchItem.InsertedChars"
args52(15).Value = 2
args52(16).Name = "SearchItem.TransliterateFlags"
args52(16).Value = 1280
args52(17).Name = "SearchItem.Command"
args52(17).Value = 3
args52(18).Name = "Quiet"
args52(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args52()

rem ----------------------------------------------------------------------
dim args53(18) as new com.sun.star.beans.PropertyValue
args53(0).Name = "SearchItem.StyleFamily"
args53(0).Value = 2
args53(1).Name = "SearchItem.CellType"
args53(1).Value = 0
args53(2).Name = "SearchItem.RowDirection"
args53(2).Value = true
args53(3).Name = "SearchItem.AllTables"
args53(3).Value = false
args53(4).Name = "SearchItem.Backward"
args53(4).Value = false
args53(5).Name = "SearchItem.Pattern"
args53(5).Value = false
args53(6).Name = "SearchItem.Content"
args53(6).Value = false
args53(7).Name = "SearchItem.AsianOptions"
args53(7).Value = false
args53(8).Name = "SearchItem.AlgorithmType"
args53(8).Value = 0
args53(9).Name = "SearchItem.SearchFlags"
args53(9).Value = 65536
args53(10).Name = "SearchItem.SearchString"
args53(10).Value = "automobilo-kontestisto"
args53(11).Name = "SearchItem.ReplaceString"
args53(11).Value = "automobilisto"
args53(12).Name = "SearchItem.Locale"
args53(12).Value = 255
args53(13).Name = "SearchItem.ChangedChars"
args53(13).Value = 2
args53(14).Name = "SearchItem.DeletedChars"
args53(14).Value = 2
args53(15).Name = "SearchItem.InsertedChars"
args53(15).Value = 2
args53(16).Name = "SearchItem.TransliterateFlags"
args53(16).Value = 1280
args53(17).Name = "SearchItem.Command"
args53(17).Value = 3
args53(18).Name = "Quiet"
args53(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args53()

rem ----------------------------------------------------------------------
dim args54(18) as new com.sun.star.beans.PropertyValue
args54(0).Name = "SearchItem.StyleFamily"
args54(0).Value = 2
args54(1).Name = "SearchItem.CellType"
args54(1).Value = 0
args54(2).Name = "SearchItem.RowDirection"
args54(2).Value = true
args54(3).Name = "SearchItem.AllTables"
args54(3).Value = false
args54(4).Name = "SearchItem.Backward"
args54(4).Value = false
args54(5).Name = "SearchItem.Pattern"
args54(5).Value = false
args54(6).Name = "SearchItem.Content"
args54(6).Value = false
args54(7).Name = "SearchItem.AsianOptions"
args54(7).Value = false
args54(8).Name = "SearchItem.AlgorithmType"
args54(8).Value = 0
args54(9).Name = "SearchItem.SearchFlags"
args54(9).Value = 65536
args54(10).Name = "SearchItem.SearchString"
args54(10).Value = "aktorulo"
args54(11).Name = "SearchItem.ReplaceString"
args54(11).Value = "aktoro"
args54(12).Name = "SearchItem.Locale"
args54(12).Value = 255
args54(13).Name = "SearchItem.ChangedChars"
args54(13).Value = 2
args54(14).Name = "SearchItem.DeletedChars"
args54(14).Value = 2
args54(15).Name = "SearchItem.InsertedChars"
args54(15).Value = 2
args54(16).Name = "SearchItem.TransliterateFlags"
args54(16).Value = 1280
args54(17).Name = "SearchItem.Command"
args54(17).Value = 3
args54(18).Name = "Quiet"
args54(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args54()

rem ----------------------------------------------------------------------
dim args55(18) as new com.sun.star.beans.PropertyValue
args55(0).Name = "SearchItem.StyleFamily"
args55(0).Value = 2
args55(1).Name = "SearchItem.CellType"
args55(1).Value = 0
args55(2).Name = "SearchItem.RowDirection"
args55(2).Value = true
args55(3).Name = "SearchItem.AllTables"
args55(3).Value = false
args55(4).Name = "SearchItem.Backward"
args55(4).Value = false
args55(5).Name = "SearchItem.Pattern"
args55(5).Value = false
args55(6).Name = "SearchItem.Content"
args55(6).Value = false
args55(7).Name = "SearchItem.AsianOptions"
args55(7).Value = false
args55(8).Name = "SearchItem.AlgorithmType"
args55(8).Value = 0
args55(9).Name = "SearchItem.SearchFlags"
args55(9).Value = 65536
args55(10).Name = "SearchItem.SearchString"
args55(10).Value = "Unioninta Nacioni"
args55(11).Name = "SearchItem.ReplaceString"
args55(11).Value = "Unionita Nacioni"
args55(12).Name = "SearchItem.Locale"
args55(12).Value = 255
args55(13).Name = "SearchItem.ChangedChars"
args55(13).Value = 2
args55(14).Name = "SearchItem.DeletedChars"
args55(14).Value = 2
args55(15).Name = "SearchItem.InsertedChars"
args55(15).Value = 2
args55(16).Name = "SearchItem.TransliterateFlags"
args55(16).Value = 1280
args55(17).Name = "SearchItem.Command"
args55(17).Value = 3
args55(18).Name = "Quiet"
args55(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args55()

rem ----------------------------------------------------------------------
dim args56(18) as new com.sun.star.beans.PropertyValue
args56(0).Name = "SearchItem.StyleFamily"
args56(0).Value = 2
args56(1).Name = "SearchItem.CellType"
args56(1).Value = 0
args56(2).Name = "SearchItem.RowDirection"
args56(2).Value = true
args56(3).Name = "SearchItem.AllTables"
args56(3).Value = false
args56(4).Name = "SearchItem.Backward"
args56(4).Value = false
args56(5).Name = "SearchItem.Pattern"
args56(5).Value = false
args56(6).Name = "SearchItem.Content"
args56(6).Value = false
args56(7).Name = "SearchItem.AsianOptions"
args56(7).Value = false
args56(8).Name = "SearchItem.AlgorithmType"
args56(8).Value = 0
args56(9).Name = "SearchItem.SearchFlags"
args56(9).Value = 65536
args56(10).Name = "SearchItem.SearchString"
args56(10).Value = "graviteso"
args56(11).Name = "SearchItem.ReplaceString"
args56(11).Value = "gravitado"
args56(12).Name = "SearchItem.Locale"
args56(12).Value = 255
args56(13).Name = "SearchItem.ChangedChars"
args56(13).Value = 2
args56(14).Name = "SearchItem.DeletedChars"
args56(14).Value = 2
args56(15).Name = "SearchItem.InsertedChars"
args56(15).Value = 2
args56(16).Name = "SearchItem.TransliterateFlags"
args56(16).Value = 1280
args56(17).Name = "SearchItem.Command"
args56(17).Value = 3
args56(18).Name = "Quiet"
args56(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args56()

rem ----------------------------------------------------------------------
dim args57(18) as new com.sun.star.beans.PropertyValue
args57(0).Name = "SearchItem.StyleFamily"
args57(0).Value = 2
args57(1).Name = "SearchItem.CellType"
args57(1).Value = 0
args57(2).Name = "SearchItem.RowDirection"
args57(2).Value = true
args57(3).Name = "SearchItem.AllTables"
args57(3).Value = false
args57(4).Name = "SearchItem.Backward"
args57(4).Value = false
args57(5).Name = "SearchItem.Pattern"
args57(5).Value = false
args57(6).Name = "SearchItem.Content"
args57(6).Value = false
args57(7).Name = "SearchItem.AsianOptions"
args57(7).Value = false
args57(8).Name = "SearchItem.AlgorithmType"
args57(8).Value = 0
args57(9).Name = "SearchItem.SearchFlags"
args57(9).Value = 65536
args57(10).Name = "SearchItem.SearchString"
args57(10).Value = "prezenti"
args57(11).Name = "SearchItem.ReplaceString"
args57(11).Value = "prizenti"
args57(12).Name = "SearchItem.Locale"
args57(12).Value = 255
args57(13).Name = "SearchItem.ChangedChars"
args57(13).Value = 2
args57(14).Name = "SearchItem.DeletedChars"
args57(14).Value = 2
args57(15).Name = "SearchItem.InsertedChars"
args57(15).Value = 2
args57(16).Name = "SearchItem.TransliterateFlags"
args57(16).Value = 1280
args57(17).Name = "SearchItem.Command"
args57(17).Value = 3
args57(18).Name = "Quiet"
args57(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args57()

rem ----------------------------------------------------------------------
dim args58(18) as new com.sun.star.beans.PropertyValue
args58(0).Name = "SearchItem.StyleFamily"
args58(0).Value = 2
args58(1).Name = "SearchItem.CellType"
args58(1).Value = 0
args58(2).Name = "SearchItem.RowDirection"
args58(2).Value = true
args58(3).Name = "SearchItem.AllTables"
args58(3).Value = false
args58(4).Name = "SearchItem.Backward"
args58(4).Value = false
args58(5).Name = "SearchItem.Pattern"
args58(5).Value = false
args58(6).Name = "SearchItem.Content"
args58(6).Value = false
args58(7).Name = "SearchItem.AsianOptions"
args58(7).Value = false
args58(8).Name = "SearchItem.AlgorithmType"
args58(8).Value = 0
args58(9).Name = "SearchItem.SearchFlags"
args58(9).Value = 65536
args58(10).Name = "SearchItem.SearchString"
args58(10).Value = "eligita"
args58(11).Name = "SearchItem.ReplaceString"
args58(11).Value = "elektita"
args58(12).Name = "SearchItem.Locale"
args58(12).Value = 255
args58(13).Name = "SearchItem.ChangedChars"
args58(13).Value = 2
args58(14).Name = "SearchItem.DeletedChars"
args58(14).Value = 2
args58(15).Name = "SearchItem.InsertedChars"
args58(15).Value = 2
args58(16).Name = "SearchItem.TransliterateFlags"
args58(16).Value = 1280
args58(17).Name = "SearchItem.Command"
args58(17).Value = 3
args58(18).Name = "Quiet"
args58(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args58()

rem ----------------------------------------------------------------------
dim args60(18) as new com.sun.star.beans.PropertyValue
args60(0).Name = "SearchItem.StyleFamily"
args60(0).Value = 2
args60(1).Name = "SearchItem.CellType"
args60(1).Value = 0
args60(2).Name = "SearchItem.RowDirection"
args60(2).Value = true
args60(3).Name = "SearchItem.AllTables"
args60(3).Value = false
args60(4).Name = "SearchItem.Backward"
args60(4).Value = false
args60(5).Name = "SearchItem.Pattern"
args60(5).Value = false
args60(6).Name = "SearchItem.Content"
args60(6).Value = false
args60(7).Name = "SearchItem.AsianOptions"
args60(7).Value = false
args60(8).Name = "SearchItem.AlgorithmType"
args60(8).Value = 0
args60(9).Name = "SearchItem.SearchFlags"
args60(9).Value = 65536
args60(10).Name = "SearchItem.SearchString"
args60(10).Value = "guverno"
args60(11).Name = "SearchItem.ReplaceString"
args60(11).Value = "guvernerio"
args60(12).Name = "SearchItem.Locale"
args60(12).Value = 255
args60(13).Name = "SearchItem.ChangedChars"
args60(13).Value = 2
args60(14).Name = "SearchItem.DeletedChars"
args60(14).Value = 2
args60(15).Name = "SearchItem.InsertedChars"
args60(15).Value = 2
args60(16).Name = "SearchItem.TransliterateFlags"
args60(16).Value = 1280
args60(17).Name = "SearchItem.Command"
args60(17).Value = 3
args60(18).Name = "Quiet"
args60(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args60()

rem ----------------------------------------------------------------------
dim args61(18) as new com.sun.star.beans.PropertyValue
args61(0).Name = "SearchItem.StyleFamily"
args61(0).Value = 2
args61(1).Name = "SearchItem.CellType"
args61(1).Value = 0
args61(2).Name = "SearchItem.RowDirection"
args61(2).Value = true
args61(3).Name = "SearchItem.AllTables"
args61(3).Value = false
args61(4).Name = "SearchItem.Backward"
args61(4).Value = false
args61(5).Name = "SearchItem.Pattern"
args61(5).Value = false
args61(6).Name = "SearchItem.Content"
args61(6).Value = false
args61(7).Name = "SearchItem.AsianOptions"
args61(7).Value = false
args61(8).Name = "SearchItem.AlgorithmType"
args61(8).Value = 0
args61(9).Name = "SearchItem.SearchFlags"
args61(9).Value = 65536
args61(10).Name = "SearchItem.SearchString"
args61(10).Value = "omna yari"
args61(11).Name = "SearchItem.ReplaceString"
args61(11).Value = "omnayare"
args61(12).Name = "SearchItem.Locale"
args61(12).Value = 255
args61(13).Name = "SearchItem.ChangedChars"
args61(13).Value = 2
args61(14).Name = "SearchItem.DeletedChars"
args61(14).Value = 2
args61(15).Name = "SearchItem.InsertedChars"
args61(15).Value = 2
args61(16).Name = "SearchItem.TransliterateFlags"
args61(16).Value = 1280
args61(17).Name = "SearchItem.Command"
args61(17).Value = 3
args61(18).Name = "Quiet"
args61(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args61()

rem ----------------------------------------------------------------------
dim args62(18) as new com.sun.star.beans.PropertyValue
args62(0).Name = "SearchItem.StyleFamily"
args62(0).Value = 2
args62(1).Name = "SearchItem.CellType"
args62(1).Value = 0
args62(2).Name = "SearchItem.RowDirection"
args62(2).Value = true
args62(3).Name = "SearchItem.AllTables"
args62(3).Value = false
args62(4).Name = "SearchItem.Backward"
args62(4).Value = false
args62(5).Name = "SearchItem.Pattern"
args62(5).Value = false
args62(6).Name = "SearchItem.Content"
args62(6).Value = false
args62(7).Name = "SearchItem.AsianOptions"
args62(7).Value = false
args62(8).Name = "SearchItem.AlgorithmType"
args62(8).Value = 0
args62(9).Name = "SearchItem.SearchFlags"
args62(9).Value = 65536
args62(10).Name = "SearchItem.SearchString"
args62(10).Value = "nexta"
args62(11).Name = "SearchItem.ReplaceString"
args62(11).Value = "sequanta"
args62(12).Name = "SearchItem.Locale"
args62(12).Value = 255
args62(13).Name = "SearchItem.ChangedChars"
args62(13).Value = 2
args62(14).Name = "SearchItem.DeletedChars"
args62(14).Value = 2
args62(15).Name = "SearchItem.InsertedChars"
args62(15).Value = 2
args62(16).Name = "SearchItem.TransliterateFlags"
args62(16).Value = 1024
args62(17).Name = "SearchItem.Command"
args62(17).Value = 3
args62(18).Name = "Quiet"
args62(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args62())

rem ----------------------------------------------------------------------
dim args63(18) as new com.sun.star.beans.PropertyValue
args63(0).Name = "SearchItem.StyleFamily"
args63(0).Value = 2
args63(1).Name = "SearchItem.CellType"
args63(1).Value = 0
args63(2).Name = "SearchItem.RowDirection"
args63(2).Value = true
args63(3).Name = "SearchItem.AllTables"
args63(3).Value = false
args63(4).Name = "SearchItem.Backward"
args63(4).Value = false
args63(5).Name = "SearchItem.Pattern"
args63(5).Value = false
args63(6).Name = "SearchItem.Content"
args63(6).Value = false
args63(7).Name = "SearchItem.AsianOptions"
args63(7).Value = false
args63(8).Name = "SearchItem.AlgorithmType"
args63(8).Value = 0
args63(9).Name = "SearchItem.SearchFlags"
args63(9).Value = 65536
args63(10).Name = "SearchItem.SearchString"
args63(10).Value = "reprizenta"
args63(11).Name = "SearchItem.ReplaceString"
args63(11).Value = "reprezenta"
args63(12).Name = "SearchItem.Locale"
args63(12).Value = 255
args63(13).Name = "SearchItem.ChangedChars"
args63(13).Value = 2
args63(14).Name = "SearchItem.DeletedChars"
args63(14).Value = 2
args63(15).Name = "SearchItem.InsertedChars"
args63(15).Value = 2
args63(16).Name = "SearchItem.TransliterateFlags"
args63(16).Value = 1280
args63(17).Name = "SearchItem.Command"
args63(17).Value = 3
args63(18).Name = "Quiet"
args63(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args63()

rem ----------------------------------------------------------------------
dim args64(18) as new com.sun.star.beans.PropertyValue
args64(0).Name = "SearchItem.StyleFamily"
args64(0).Value = 2
args64(1).Name = "SearchItem.CellType"
args64(1).Value = 0
args64(2).Name = "SearchItem.RowDirection"
args64(2).Value = true
args64(3).Name = "SearchItem.AllTables"
args64(3).Value = false
args64(4).Name = "SearchItem.Backward"
args64(4).Value = false
args64(5).Name = "SearchItem.Pattern"
args64(5).Value = false
args64(6).Name = "SearchItem.Content"
args64(6).Value = false
args64(7).Name = "SearchItem.AsianOptions"
args64(7).Value = false
args64(8).Name = "SearchItem.AlgorithmType"
args64(8).Value = 0
args64(9).Name = "SearchItem.SearchFlags"
args64(9).Value = 65536
args64(10).Name = "SearchItem.SearchString"
args64(10).Value = "Unioninta Rejio"
args64(11).Name = "SearchItem.ReplaceString"
args64(11).Value = "Unionita Rejio"
args64(12).Name = "SearchItem.Locale"
args64(12).Value = 255
args64(13).Name = "SearchItem.ChangedChars"
args64(13).Value = 2
args64(14).Name = "SearchItem.DeletedChars"
args64(14).Value = 2
args64(15).Name = "SearchItem.InsertedChars"
args64(15).Value = 2
args64(16).Name = "SearchItem.TransliterateFlags"
args64(16).Value = 1280
args64(17).Name = "SearchItem.Command"
args64(17).Value = 3
args64(18).Name = "Quiet"
args64(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args64()

rem ----------------------------------------------------------------------
dim args65(18) as new com.sun.star.beans.PropertyValue
args65(0).Name = "SearchItem.StyleFamily"
args65(0).Value = 2
args65(1).Name = "SearchItem.CellType"
args65(1).Value = 0
args65(2).Name = "SearchItem.RowDirection"
args65(2).Value = true
args65(3).Name = "SearchItem.AllTables"
args65(3).Value = false
args65(4).Name = "SearchItem.Backward"
args65(4).Value = false
args65(5).Name = "SearchItem.Pattern"
args65(5).Value = false
args65(6).Name = "SearchItem.Content"
args65(6).Value = false
args65(7).Name = "SearchItem.AsianOptions"
args65(7).Value = false
args65(8).Name = "SearchItem.AlgorithmType"
args65(8).Value = 0
args65(9).Name = "SearchItem.SearchFlags"
args65(9).Value = 65536
args65(10).Name = "SearchItem.SearchString"
args65(10).Value = "predominanta"
args65(11).Name = "SearchItem.ReplaceString"
args65(11).Value = "dominacanta"
args65(12).Name = "SearchItem.Locale"
args65(12).Value = 255
args65(13).Name = "SearchItem.ChangedChars"
args65(13).Value = 2
args65(14).Name = "SearchItem.DeletedChars"
args65(14).Value = 2
args65(15).Name = "SearchItem.InsertedChars"
args65(15).Value = 2
args65(16).Name = "SearchItem.TransliterateFlags"
args65(16).Value = 1280
args65(17).Name = "SearchItem.Command"
args65(17).Value = 3
args65(18).Name = "Quiet"
args65(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args65()

rem ----------------------------------------------------------------------
dim args66(18) as new com.sun.star.beans.PropertyValue

dim JANUARO1, JANUARO2 as string
For I = 1 to 31
JANUARO1 = Ltrim(Str$(I))+" di januaro"
JANUARO2 = Ltrim(Str$(I))+"ma di januaro"

args66(0).Name = "SearchItem.StyleFamily"
args66(0).Value = 2
args66(1).Name = "SearchItem.CellType"
args66(1).Value = 0
args66(2).Name = "SearchItem.RowDirection"
args66(2).Value = true
args66(3).Name = "SearchItem.AllTables"
args66(3).Value = false
args66(4).Name = "SearchItem.Backward"
args66(4).Value = false
args66(5).Name = "SearchItem.Pattern"
args66(5).Value = false
args66(6).Name = "SearchItem.Content"
args66(6).Value = false
args66(7).Name = "SearchItem.AsianOptions"
args66(7).Value = false
args66(8).Name = "SearchItem.AlgorithmType"
args66(8).Value = 0
args66(9).Name = "SearchItem.SearchFlags"
args66(9).Value = 65536
args66(10).Name = "SearchItem.SearchString"
args66(10).Value = JANUARO1
args66(11).Name = "SearchItem.ReplaceString"
args66(11).Value = JANUARO2
args66(12).Name = "SearchItem.Locale"
args66(12).Value = 255
args66(13).Name = "SearchItem.ChangedChars"
args66(13).Value = 2
args66(14).Name = "SearchItem.DeletedChars"
args66(14).Value = 2
args66(15).Name = "SearchItem.InsertedChars"
args66(15).Value = 2
args66(16).Name = "SearchItem.TransliterateFlags"
args66(16).Value = 1280
args66(17).Name = "SearchItem.Command"
args66(17).Value = 3
args66(18).Name = "Quiet"
args66(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args66()

Next I

rem ----------------------------------------------------------------------
dim args67(18) as new com.sun.star.beans.PropertyValue
args67(0).Name = "SearchItem.StyleFamily"
args67(0).Value = 2
args67(1).Name = "SearchItem.CellType"
args67(1).Value = 0
args67(2).Name = "SearchItem.RowDirection"
args67(2).Value = true
args67(3).Name = "SearchItem.AllTables"
args67(3).Value = false
args67(4).Name = "SearchItem.Backward"
args67(4).Value = false
args67(5).Name = "SearchItem.Pattern"
args67(5).Value = false
args67(6).Name = "SearchItem.Content"
args67(6).Value = false
args67(7).Name = "SearchItem.AsianOptions"
args67(7).Value = false
args67(8).Name = "SearchItem.AlgorithmType"
args67(8).Value = 0
args67(9).Name = "SearchItem.SearchFlags"
args67(9).Value = 65536
args67(10).Name = "SearchItem.SearchString"
args67(10).Value = "rezigna"
args67(11).Name = "SearchItem.ReplaceString"
args67(11).Value = "renunca"
args67(12).Name = "SearchItem.Locale"
args67(12).Value = 255
args67(13).Name = "SearchItem.ChangedChars"
args67(13).Value = 2
args67(14).Name = "SearchItem.DeletedChars"
args67(14).Value = 2
args67(15).Name = "SearchItem.InsertedChars"
args67(15).Value = 2
args67(16).Name = "SearchItem.TransliterateFlags"
args67(16).Value = 1280
args67(17).Name = "SearchItem.Command"
args67(17).Value = 3
args67(18).Name = "Quiet"
args67(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args67()

rem ----------------------------------------------------------------------
dim args68(18) as new com.sun.star.beans.PropertyValue
args68(0).Name = "SearchItem.StyleFamily"
args68(0).Value = 2
args68(1).Name = "SearchItem.CellType"
args68(1).Value = 0
args68(2).Name = "SearchItem.RowDirection"
args68(2).Value = true
args68(3).Name = "SearchItem.AllTables"
args68(3).Value = false
args68(4).Name = "SearchItem.Backward"
args68(4).Value = false
args68(5).Name = "SearchItem.Pattern"
args68(5).Value = false
args68(6).Name = "SearchItem.Content"
args68(6).Value = false
args68(7).Name = "SearchItem.AsianOptions"
args68(7).Value = false
args68(8).Name = "SearchItem.AlgorithmType"
args68(8).Value = 0
args68(9).Name = "SearchItem.SearchFlags"
args68(9).Value = 65536
args68(10).Name = "SearchItem.SearchString"
args68(10).Value = "Santo Paulus"
args68(11).Name = "SearchItem.ReplaceString"
args68(11).Value = "Santa Paulus"
args68(12).Name = "SearchItem.Locale"
args68(12).Value = 255
args68(13).Name = "SearchItem.ChangedChars"
args68(13).Value = 2
args68(14).Name = "SearchItem.DeletedChars"
args68(14).Value = 2
args68(15).Name = "SearchItem.InsertedChars"
args68(15).Value = 2
args68(16).Name = "SearchItem.TransliterateFlags"
args68(16).Value = 1280
args68(17).Name = "SearchItem.Command"
args68(17).Value = 3
args68(18).Name = "Quiet"
args68(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args68()

rem ----------------------------------------------------------------------
dim args69(18) as new com.sun.star.beans.PropertyValue
args69(0).Name = "SearchItem.StyleFamily"
args69(0).Value = 2
args69(1).Name = "SearchItem.CellType"
args69(1).Value = 0
args69(2).Name = "SearchItem.RowDirection"
args69(2).Value = true
args69(3).Name = "SearchItem.AllTables"
args69(3).Value = false
args69(4).Name = "SearchItem.Backward"
args69(4).Value = false
args69(5).Name = "SearchItem.Pattern"
args69(5).Value = false
args69(6).Name = "SearchItem.Content"
args69(6).Value = false
args69(7).Name = "SearchItem.AsianOptions"
args69(7).Value = false
args69(8).Name = "SearchItem.AlgorithmType"
args69(8).Value = 0
args69(9).Name = "SearchItem.SearchFlags"
args69(9).Value = 65536
args69(10).Name = "SearchItem.SearchString"
args69(10).Value = "qui evas"
args69(11).Name = "SearchItem.ReplaceString"
args69(11).Value = "evante"
args69(12).Name = "SearchItem.Locale"
args69(12).Value = 255
args69(13).Name = "SearchItem.ChangedChars"
args69(13).Value = 2
args69(14).Name = "SearchItem.DeletedChars"
args69(14).Value = 2
args69(15).Name = "SearchItem.InsertedChars"
args69(15).Value = 2
args69(16).Name = "SearchItem.TransliterateFlags"
args69(16).Value = 1280
args69(17).Name = "SearchItem.Command"
args69(17).Value = 3
args69(18).Name = "Quiet"
args69(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args69()

rem ----------------------------------------------------------------------
dim args70(18) as new com.sun.star.beans.PropertyValue
args70(0).Name = "SearchItem.StyleFamily"
args70(0).Value = 2
args70(1).Name = "SearchItem.CellType"
args70(1).Value = 0
args70(2).Name = "SearchItem.RowDirection"
args70(2).Value = true
args70(3).Name = "SearchItem.AllTables"
args70(3).Value = false
args70(4).Name = "SearchItem.Backward"
args70(4).Value = false
args70(5).Name = "SearchItem.Pattern"
args70(5).Value = false
args70(6).Name = "SearchItem.Content"
args70(6).Value = false
args70(7).Name = "SearchItem.AsianOptions"
args70(7).Value = false
args70(8).Name = "SearchItem.AlgorithmType"
args70(8).Value = 0
args70(9).Name = "SearchItem.SearchFlags"
args70(9).Value = 65536
args70(10).Name = "SearchItem.SearchString"
args70(10).Value = "aktorino di cinemo"
args70(11).Name = "SearchItem.ReplaceString"
args70(11).Value = "cinem-aktorino"
args70(12).Name = "SearchItem.Locale"
args70(12).Value = 255
args70(13).Name = "SearchItem.ChangedChars"
args70(13).Value = 2
args70(14).Name = "SearchItem.DeletedChars"
args70(14).Value = 2
args70(15).Name = "SearchItem.InsertedChars"
args70(15).Value = 2
args70(16).Name = "SearchItem.TransliterateFlags"
args70(16).Value = 1280
args70(17).Name = "SearchItem.Command"
args70(17).Value = 3
args70(18).Name = "Quiet"
args70(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args70()

rem ----------------------------------------------------------------------
dim args71(18) as new com.sun.star.beans.PropertyValue
args71(0).Name = "SearchItem.StyleFamily"
args71(0).Value = 2
args71(1).Name = "SearchItem.CellType"
args71(1).Value = 0
args71(2).Name = "SearchItem.RowDirection"
args71(2).Value = true
args71(3).Name = "SearchItem.AllTables"
args71(3).Value = false
args71(4).Name = "SearchItem.Backward"
args71(4).Value = false
args71(5).Name = "SearchItem.Pattern"
args71(5).Value = false
args71(6).Name = "SearchItem.Content"
args71(6).Value = false
args71(7).Name = "SearchItem.AsianOptions"
args71(7).Value = false
args71(8).Name = "SearchItem.AlgorithmType"
args71(8).Value = 0
args71(9).Name = "SearchItem.SearchFlags"
args71(9).Value = 65536
args71(10).Name = "SearchItem.SearchString"
args71(10).Value = "aktorulo di cinemo"
args71(11).Name = "SearchItem.ReplaceString"
args71(11).Value = "cinem-aktorulo"
args71(12).Name = "SearchItem.Locale"
args71(12).Value = 255
args71(13).Name = "SearchItem.ChangedChars"
args71(13).Value = 2
args71(14).Name = "SearchItem.DeletedChars"
args71(14).Value = 2
args71(15).Name = "SearchItem.InsertedChars"
args71(15).Value = 2
args71(16).Name = "SearchItem.TransliterateFlags"
args71(16).Value = 1280
args71(17).Name = "SearchItem.Command"
args71(17).Value = 3
args71(18).Name = "Quiet"
args71(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args71()


rem ----------------------------------------------------------------------
dim args72(18) as new com.sun.star.beans.PropertyValue
args72(0).Name = "SearchItem.StyleFamily"
args72(0).Value = 2
args72(1).Name = "SearchItem.CellType"
args72(1).Value = 0
args72(2).Name = "SearchItem.RowDirection"
args72(2).Value = true
args72(3).Name = "SearchItem.AllTables"
args72(3).Value = false
args72(4).Name = "SearchItem.Backward"
args72(4).Value = false
args72(5).Name = "SearchItem.Pattern"
args72(5).Value = false
args72(6).Name = "SearchItem.Content"
args72(6).Value = false
args72(7).Name = "SearchItem.AsianOptions"
args72(7).Value = false
args72(8).Name = "SearchItem.AlgorithmType"
args72(8).Value = 0
args72(9).Name = "SearchItem.SearchFlags"
args72(9).Value = 65536
args72(10).Name = "SearchItem.SearchString"
args72(10).Value = "Malgre ol"
args72(11).Name = "SearchItem.ReplaceString"
args72(11).Value = "Quankam ol"
args72(12).Name = "SearchItem.Locale"
args72(12).Value = 255
args72(13).Name = "SearchItem.ChangedChars"
args72(13).Value = 2
args72(14).Name = "SearchItem.DeletedChars"
args72(14).Value = 2
args72(15).Name = "SearchItem.InsertedChars"
args72(15).Value = 2
args72(16).Name = "SearchItem.TransliterateFlags"
args72(16).Value = 1280
args72(17).Name = "SearchItem.Command"
args72(17).Value = 3
args72(18).Name = "Quiet"
args72(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args72()

rem ----------------------------------------------------------------------
dim args73(18) as new com.sun.star.beans.PropertyValue
args73(0).Name = "SearchItem.StyleFamily"
args73(0).Value = 2
args73(1).Name = "SearchItem.CellType"
args73(1).Value = 0
args73(2).Name = "SearchItem.RowDirection"
args73(2).Value = true
args73(3).Name = "SearchItem.AllTables"
args73(3).Value = false
args73(4).Name = "SearchItem.Backward"
args73(4).Value = false
args73(5).Name = "SearchItem.Pattern"
args73(5).Value = false
args73(6).Name = "SearchItem.Content"
args73(6).Value = false
args73(7).Name = "SearchItem.AsianOptions"
args73(7).Value = false
args73(8).Name = "SearchItem.AlgorithmType"
args73(8).Value = 0
args73(9).Name = "SearchItem.SearchFlags"
args73(9).Value = 65536
args73(10).Name = "SearchItem.SearchString"
args73(10).Value = "averajal"
args73(11).Name = "SearchItem.ReplaceString"
args73(11).Value = "mezavalor"
args73(12).Name = "SearchItem.Locale"
args73(12).Value = 255
args73(13).Name = "SearchItem.ChangedChars"
args73(13).Value = 2
args73(14).Name = "SearchItem.DeletedChars"
args73(14).Value = 2
args73(15).Name = "SearchItem.InsertedChars"
args73(15).Value = 2
args73(16).Name = "SearchItem.TransliterateFlags"
args73(16).Value = 1280
args73(17).Name = "SearchItem.Command"
args73(17).Value = 3
args73(18).Name = "Quiet"
args73(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args73()

rem ----------------------------------------------------------------------
dim args74(18) as new com.sun.star.beans.PropertyValue
args74(0).Name = "SearchItem.StyleFamily"
args74(0).Value = 2
args74(1).Name = "SearchItem.CellType"
args74(1).Value = 0
args74(2).Name = "SearchItem.RowDirection"
args74(2).Value = true
args74(3).Name = "SearchItem.AllTables"
args74(3).Value = false
args74(4).Name = "SearchItem.Backward"
args74(4).Value = false
args74(5).Name = "SearchItem.Pattern"
args74(5).Value = false
args74(6).Name = "SearchItem.Content"
args74(6).Value = false
args74(7).Name = "SearchItem.AsianOptions"
args74(7).Value = false
args74(8).Name = "SearchItem.AlgorithmType"
args74(8).Value = 0
args74(9).Name = "SearchItem.SearchFlags"
args74(9).Value = 65536
args74(10).Name = "SearchItem.SearchString"
args74(10).Value = "datizas de"
args74(11).Name = "SearchItem.ReplaceString"
args74(11).Value = "evas de"
args74(12).Name = "SearchItem.Locale"
args74(12).Value = 255
args74(13).Name = "SearchItem.ChangedChars"
args74(13).Value = 2
args74(14).Name = "SearchItem.DeletedChars"
args74(14).Value = 2
args74(15).Name = "SearchItem.InsertedChars"
args74(15).Value = 2
args74(16).Name = "SearchItem.TransliterateFlags"
args74(16).Value = 1280
args74(17).Name = "SearchItem.Command"
args74(17).Value = 3
args74(18).Name = "Quiet"
args74(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args74()

rem ----------------------------------------------------------------------
dim args75(18) as new com.sun.star.beans.PropertyValue
args75(0).Name = "SearchItem.StyleFamily"
args75(0).Value = 2
args75(1).Name = "SearchItem.CellType"
args75(1).Value = 0
args75(2).Name = "SearchItem.RowDirection"
args75(2).Value = true
args75(3).Name = "SearchItem.AllTables"
args75(3).Value = false
args75(4).Name = "SearchItem.Backward"
args75(4).Value = false
args75(5).Name = "SearchItem.Pattern"
args75(5).Value = false
args75(6).Name = "SearchItem.Content"
args75(6).Value = false
args75(7).Name = "SearchItem.AsianOptions"
args75(7).Value = false
args75(8).Name = "SearchItem.AlgorithmType"
args75(8).Value = 0
args75(9).Name = "SearchItem.SearchFlags"
args75(9).Value = 65536
args75(10).Name = "SearchItem.SearchString"
args75(10).Value = "esas elektita da "
args75(11).Name = "SearchItem.ReplaceString"
args75(11).Value = "elektesas dal "
args75(12).Name = "SearchItem.Locale"
args75(12).Value = 255
args75(13).Name = "SearchItem.ChangedChars"
args75(13).Value = 2
args75(14).Name = "SearchItem.DeletedChars"
args75(14).Value = 2
args75(15).Name = "SearchItem.InsertedChars"
args75(15).Value = 2
args75(16).Name = "SearchItem.TransliterateFlags"
args75(16).Value = 1280
args75(17).Name = "SearchItem.Command"
args75(17).Value = 3
args75(18).Name = "Quiet"
args75(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args75()

rem ----------------------------------------------------------------------
dim args76(18) as new com.sun.star.beans.PropertyValue
args76(0).Name = "SearchItem.StyleFamily"
args76(0).Value = 2
args76(1).Name = "SearchItem.CellType"
args76(1).Value = 0
args76(2).Name = "SearchItem.RowDirection"
args76(2).Value = true
args76(3).Name = "SearchItem.AllTables"
args76(3).Value = false
args76(4).Name = "SearchItem.Backward"
args76(4).Value = false
args76(5).Name = "SearchItem.Pattern"
args76(5).Value = false
args76(6).Name = "SearchItem.Content"
args76(6).Value = false
args76(7).Name = "SearchItem.AsianOptions"
args76(7).Value = false
args76(8).Name = "SearchItem.AlgorithmType"
args76(8).Value = 0
args76(9).Name = "SearchItem.SearchFlags"
args76(9).Value = 65536
args76(10).Name = "SearchItem.SearchString"
args76(10).Value = "Germaniana"
args76(11).Name = "SearchItem.ReplaceString"
args76(11).Value = "Germana"
args76(12).Name = "SearchItem.Locale"
args76(12).Value = 255
args76(13).Name = "SearchItem.ChangedChars"
args76(13).Value = 2
args76(14).Name = "SearchItem.DeletedChars"
args76(14).Value = 2
args76(15).Name = "SearchItem.InsertedChars"
args76(15).Value = 2
args76(16).Name = "SearchItem.TransliterateFlags"
args76(16).Value = 1280
args76(17).Name = "SearchItem.Command"
args76(17).Value = 3
args76(18).Name = "Quiet"
args76(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args76()

rem ----------------------------------------------------------------------
dim args77(18) as new com.sun.star.beans.PropertyValue
args77(0).Name = "SearchItem.StyleFamily"
args77(0).Value = 2
args77(1).Name = "SearchItem.CellType"
args77(1).Value = 0
args77(2).Name = "SearchItem.RowDirection"
args77(2).Value = true
args77(3).Name = "SearchItem.AllTables"
args77(3).Value = false
args77(4).Name = "SearchItem.Backward"
args77(4).Value = false
args77(5).Name = "SearchItem.Pattern"
args77(5).Value = false
args77(6).Name = "SearchItem.Content"
args77(6).Value = false
args77(7).Name = "SearchItem.AsianOptions"
args77(7).Value = false
args77(8).Name = "SearchItem.AlgorithmType"
args77(8).Value = 0
args77(9).Name = "SearchItem.SearchFlags"
args77(9).Value = 65536
args77(10).Name = "SearchItem.SearchString"
args77(10).Value = "an autori"
args77(11).Name = "SearchItem.ReplaceString"
args77(11).Value = "ana skriptisti"
args77(12).Name = "SearchItem.Locale"
args77(12).Value = 255
args77(13).Name = "SearchItem.ChangedChars"
args77(13).Value = 2
args77(14).Name = "SearchItem.DeletedChars"
args77(14).Value = 2
args77(15).Name = "SearchItem.InsertedChars"
args77(15).Value = 2
args77(16).Name = "SearchItem.TransliterateFlags"
args77(16).Value = 1280
args77(17).Name = "SearchItem.Command"
args77(17).Value = 3
args77(18).Name = "Quiet"
args77(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args77())

rem ----------------------------------------------------------------------
dim args78(18) as new com.sun.star.beans.PropertyValue
args78(0).Name = "SearchItem.StyleFamily"
args78(0).Value = 2
args78(1).Name = "SearchItem.CellType"
args78(1).Value = 0
args78(2).Name = "SearchItem.RowDirection"
args78(2).Value = true
args78(3).Name = "SearchItem.AllTables"
args78(3).Value = false
args78(4).Name = "SearchItem.Backward"
args78(4).Value = false
args78(5).Name = "SearchItem.Pattern"
args78(5).Value = false
args78(6).Name = "SearchItem.Content"
args78(6).Value = false
args78(7).Name = "SearchItem.AsianOptions"
args78(7).Value = false
args78(8).Name = "SearchItem.AlgorithmType"
args78(8).Value = 0
args78(9).Name = "SearchItem.SearchFlags"
args78(9).Value = 65536
args78(10).Name = "SearchItem.SearchString"
args78(10).Value = "esis establisita"
args78(11).Name = "SearchItem.ReplaceString"
args78(11).Value = "establisesis"
args78(12).Name = "SearchItem.Locale"
args78(12).Value = 255
args78(13).Name = "SearchItem.ChangedChars"
args78(13).Value = 2
args78(14).Name = "SearchItem.DeletedChars"
args78(14).Value = 2
args78(15).Name = "SearchItem.InsertedChars"
args78(15).Value = 2
args78(16).Name = "SearchItem.TransliterateFlags"
args78(16).Value = 1280
args78(17).Name = "SearchItem.Command"
args78(17).Value = 3
args78(18).Name = "Quiet"
args78(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args78()

rem ----------------------------------------------------------------------
dim args79(18) as new com.sun.star.beans.PropertyValue
args79(0).Name = "SearchItem.StyleFamily"
args79(0).Value = 2
args79(1).Name = "SearchItem.CellType"
args79(1).Value = 0
args79(2).Name = "SearchItem.RowDirection"
args79(2).Value = true
args79(3).Name = "SearchItem.AllTables"
args79(3).Value = false
args79(4).Name = "SearchItem.Backward"
args79(4).Value = false
args79(5).Name = "SearchItem.Pattern"
args79(5).Value = false
args79(6).Name = "SearchItem.Content"
args79(6).Value = false
args79(7).Name = "SearchItem.AsianOptions"
args79(7).Value = false
args79(8).Name = "SearchItem.AlgorithmType"
args79(8).Value = 0
args79(9).Name = "SearchItem.SearchFlags"
args79(9).Value = 65536
args79(10).Name = "SearchItem.SearchString"
args79(10).Value = "{{Nobel-laureati en "
args79(11).Name = "SearchItem.ReplaceString"
args79(11).Value = "{{Nobel-laureati pri "
args79(12).Name = "SearchItem.Locale"
args79(12).Value = 255
args79(13).Name = "SearchItem.ChangedChars"
args79(13).Value = 2
args79(14).Name = "SearchItem.DeletedChars"
args79(14).Value = 2
args79(15).Name = "SearchItem.InsertedChars"
args79(15).Value = 2
args79(16).Name = "SearchItem.TransliterateFlags"
args79(16).Value = 1280
args79(17).Name = "SearchItem.Command"
args79(17).Value = 3
args79(18).Name = "Quiet"
args79(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args79()

rem ----------------------------------------------------------------------
dim args80(18) as new com.sun.star.beans.PropertyValue
args80(0).Name = "SearchItem.StyleFamily"
args80(0).Value = 2
args80(1).Name = "SearchItem.CellType"
args80(1).Value = 0
args80(2).Name = "SearchItem.RowDirection"
args80(2).Value = true
args80(3).Name = "SearchItem.AllTables"
args80(3).Value = false
args80(4).Name = "SearchItem.Backward"
args80(4).Value = false
args80(5).Name = "SearchItem.Pattern"
args80(5).Value = false
args80(6).Name = "SearchItem.Content"
args80(6).Value = false
args80(7).Name = "SearchItem.AsianOptions"
args80(7).Value = false
args80(8).Name = "SearchItem.AlgorithmType"
args80(8).Value = 0
args80(9).Name = "SearchItem.SearchFlags"
args80(9).Value = 65536
args80(10).Name = "SearchItem.SearchString"
args80(10).Value = "an sk"
args80(11).Name = "SearchItem.ReplaceString"
args80(11).Value = "ana sk"
args80(12).Name = "SearchItem.Locale"
args80(12).Value = 255
args80(13).Name = "SearchItem.ChangedChars"
args80(13).Value = 2
args80(14).Name = "SearchItem.DeletedChars"
args80(14).Value = 2
args80(15).Name = "SearchItem.InsertedChars"
args80(15).Value = 2
args80(16).Name = "SearchItem.TransliterateFlags"
args80(16).Value = 1280
args80(17).Name = "SearchItem.Command"
args80(17).Value = 3
args80(18).Name = "Quiet"
args80(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args80()

rem ----------------------------------------------------------------------
dim args81(18) as new com.sun.star.beans.PropertyValue
args81(0).Name = "SearchItem.StyleFamily"
args81(0).Value = 2
args81(1).Name = "SearchItem.CellType"
args81(1).Value = 0
args81(2).Name = "SearchItem.RowDirection"
args81(2).Value = true
args81(3).Name = "SearchItem.AllTables"
args81(3).Value = false
args81(4).Name = "SearchItem.Backward"
args81(4).Value = false
args81(5).Name = "SearchItem.Pattern"
args81(5).Value = false
args81(6).Name = "SearchItem.Content"
args81(6).Value = false
args81(7).Name = "SearchItem.AsianOptions"
args81(7).Value = false
args81(8).Name = "SearchItem.AlgorithmType"
args81(8).Value = 0
args81(9).Name = "SearchItem.SearchFlags"
args81(9).Value = 65536
args81(10).Name = "SearchItem.SearchString"
args81(10).Value = "Ne havas imajo en Commons"
args81(11).Name = "SearchItem.ReplaceString"
args81(11).Value = "[[Arkivo:Ne havas libera imajo.svg|180px]]"
args81(12).Name = "SearchItem.Locale"
args81(12).Value = 255
args81(13).Name = "SearchItem.ChangedChars"
args81(13).Value = 2
args81(14).Name = "SearchItem.DeletedChars"
args81(14).Value = 2
args81(15).Name = "SearchItem.InsertedChars"
args81(15).Value = 2
args81(16).Name = "SearchItem.TransliterateFlags"
args81(16).Value = 1280
args81(17).Name = "SearchItem.Command"
args81(17).Value = 3
args81(18).Name = "Quiet"
args81(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args81()

rem ----------------------------------------------------------------------
dim args82(18) as new com.sun.star.beans.PropertyValue
args82(0).Name = "SearchItem.StyleFamily"
args82(0).Value = 2
args82(1).Name = "SearchItem.CellType"
args82(1).Value = 0
args82(2).Name = "SearchItem.RowDirection"
args82(2).Value = true
args82(3).Name = "SearchItem.AllTables"
args82(3).Value = false
args82(4).Name = "SearchItem.Backward"
args82(4).Value = false
args82(5).Name = "SearchItem.Pattern"
args82(5).Value = false
args82(6).Name = "SearchItem.Content"
args82(6).Value = false
args82(7).Name = "SearchItem.AsianOptions"
args82(7).Value = false
args82(8).Name = "SearchItem.AlgorithmType"
args82(8).Value = 0
args82(9).Name = "SearchItem.SearchFlags"
args82(9).Value = 65536
args82(10).Name = "SearchItem.SearchString"
args82(10).Value = "havas libera imajo.jpg"
args82(11).Name = "SearchItem.ReplaceString"
args82(11).Value = "havas libera imajo.svg"
args82(12).Name = "SearchItem.Locale"
args82(12).Value = 255
args82(13).Name = "SearchItem.ChangedChars"
args82(13).Value = 2
args82(14).Name = "SearchItem.DeletedChars"
args82(14).Value = 2
args82(15).Name = "SearchItem.InsertedChars"
args82(15).Value = 2
args82(16).Name = "SearchItem.TransliterateFlags"
args82(16).Value = 1280
args82(17).Name = "SearchItem.Command"
args82(17).Value = 3
args82(18).Name = "Quiet"
args82(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args82()

rem ----------------------------------------------------------------------
dim args83(18) as new com.sun.star.beans.PropertyValue
args83(0).Name = "SearchItem.StyleFamily"
args83(0).Value = 2
args83(1).Name = "SearchItem.CellType"
args83(1).Value = 0
args83(2).Name = "SearchItem.RowDirection"
args83(2).Value = true
args83(3).Name = "SearchItem.AllTables"
args83(3).Value = false
args83(4).Name = "SearchItem.Backward"
args83(4).Value = false
args83(5).Name = "SearchItem.Pattern"
args83(5).Value = false
args83(6).Name = "SearchItem.Content"
args83(6).Value = false
args83(7).Name = "SearchItem.AsianOptions"
args83(7).Value = false
args83(8).Name = "SearchItem.AlgorithmType"
args83(8).Value = 0
args83(9).Name = "SearchItem.SearchFlags"
args83(9).Value = 65536
args83(10).Name = "SearchItem.SearchString"
args83(10).Value = "Unioninta Povi"
args83(11).Name = "SearchItem.ReplaceString"
args83(11).Value = "westala federiti"
args83(12).Name = "SearchItem.Locale"
args83(12).Value = 255
args83(13).Name = "SearchItem.ChangedChars"
args83(13).Value = 2
args83(14).Name = "SearchItem.DeletedChars"
args83(14).Value = 2
args83(15).Name = "SearchItem.InsertedChars"
args83(15).Value = 2
args83(16).Name = "SearchItem.TransliterateFlags"
args83(16).Value = 1280
args83(17).Name = "SearchItem.Command"
args83(17).Value = 3
args83(18).Name = "Quiet"
args83(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args83()

rem ----------------------------------------------------------------------
dim args84(18) as new com.sun.star.beans.PropertyValue
args84(0).Name = "SearchItem.StyleFamily"
args84(0).Value = 2
args84(1).Name = "SearchItem.CellType"
args84(1).Value = 0
args84(2).Name = "SearchItem.RowDirection"
args84(2).Value = true
args84(3).Name = "SearchItem.AllTables"
args84(3).Value = false
args84(4).Name = "SearchItem.Backward"
args84(4).Value = false
args84(5).Name = "SearchItem.Pattern"
args84(5).Value = false
args84(6).Name = "SearchItem.Content"
args84(6).Value = false
args84(7).Name = "SearchItem.AsianOptions"
args84(7).Value = false
args84(8).Name = "SearchItem.AlgorithmType"
args84(8).Value = 0
args84(9).Name = "SearchItem.SearchFlags"
args84(9).Value = 65536
args84(10).Name = "SearchItem.SearchString"
args84(10).Value = " di guvernisteso"
args84(11).Name = "SearchItem.ReplaceString"
args84(11).Value = " dil ofico"
args84(12).Name = "SearchItem.Locale"
args84(12).Value = 255
args84(13).Name = "SearchItem.ChangedChars"
args84(13).Value = 2
args84(14).Name = "SearchItem.DeletedChars"
args84(14).Value = 2
args84(15).Name = "SearchItem.InsertedChars"
args84(15).Value = 2
args84(16).Name = "SearchItem.TransliterateFlags"
args84(16).Value = 1280
args84(17).Name = "SearchItem.Command"
args84(17).Value = 3
args84(18).Name = "Quiet"
args84(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args84()

rem ----------------------------------------------------------------------
dim args85(18) as new com.sun.star.beans.PropertyValue
args85(0).Name = "SearchItem.StyleFamily"
args85(0).Value = 2
args85(1).Name = "SearchItem.CellType"
args85(1).Value = 0
args85(2).Name = "SearchItem.RowDirection"
args85(2).Value = true
args85(3).Name = "SearchItem.AllTables"
args85(3).Value = false
args85(4).Name = "SearchItem.Backward"
args85(4).Value = false
args85(5).Name = "SearchItem.Pattern"
args85(5).Value = false
args85(6).Name = "SearchItem.Content"
args85(6).Value = false
args85(7).Name = "SearchItem.AsianOptions"
args85(7).Value = false
args85(8).Name = "SearchItem.AlgorithmType"
args85(8).Value = 0
args85(9).Name = "SearchItem.SearchFlags"
args85(9).Value = 65536
args85(10).Name = "SearchItem.SearchString"
args85(10).Value = "= p"
args85(11).Name = "SearchItem.ReplaceString"
args85(11).Value = "= P"
args85(12).Name = "SearchItem.Locale"
args85(12).Value = 255
args85(13).Name = "SearchItem.ChangedChars"
args85(13).Value = 2
args85(14).Name = "SearchItem.DeletedChars"
args85(14).Value = 2
args85(15).Name = "SearchItem.InsertedChars"
args85(15).Value = 2
args85(16).Name = "SearchItem.TransliterateFlags"
args85(16).Value = 1280
args85(17).Name = "SearchItem.Command"
args85(17).Value = 3
args85(18).Name = "Quiet"
args85(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args85()

rem ----------------------------------------------------------------------
dim args86(18) as new com.sun.star.beans.PropertyValue
args86(0).Name = "SearchItem.StyleFamily"
args86(0).Value = 2
args86(1).Name = "SearchItem.CellType"
args86(1).Value = 0
args86(2).Name = "SearchItem.RowDirection"
args86(2).Value = true
args86(3).Name = "SearchItem.AllTables"
args86(3).Value = false
args86(4).Name = "SearchItem.Backward"
args86(4).Value = false
args86(5).Name = "SearchItem.Pattern"
args86(5).Value = false
args86(6).Name = "SearchItem.Content"
args86(6).Value = false
args86(7).Name = "SearchItem.AsianOptions"
args86(7).Value = false
args86(8).Name = "SearchItem.AlgorithmType"
args86(8).Value = 0
args86(9).Name = "SearchItem.SearchFlags"
args86(9).Value = 65536
args86(10).Name = "SearchItem.SearchString"
args86(10).Value = "io:p"
args86(11).Name = "SearchItem.ReplaceString"
args86(11).Value = "io:P"
args86(12).Name = "SearchItem.Locale"
args86(12).Value = 255
args86(13).Name = "SearchItem.ChangedChars"
args86(13).Value = 2
args86(14).Name = "SearchItem.DeletedChars"
args86(14).Value = 2
args86(15).Name = "SearchItem.InsertedChars"
args86(15).Value = 2
args86(16).Name = "SearchItem.TransliterateFlags"
args86(16).Value = 1280
args86(17).Name = "SearchItem.Command"
args86(17).Value = 3
args86(18).Name = "Quiet"
args86(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args86()

rem ----------------------------------------------------------------------
dim args87(18) as new com.sun.star.beans.PropertyValue
args87(0).Name = "SearchItem.StyleFamily"
args87(0).Value = 2
args87(1).Name = "SearchItem.CellType"
args87(1).Value = 0
args87(2).Name = "SearchItem.RowDirection"
args87(2).Value = true
args87(3).Name = "SearchItem.AllTables"
args87(3).Value = false
args87(4).Name = "SearchItem.Backward"
args87(4).Value = false
args87(5).Name = "SearchItem.Pattern"
args87(5).Value = false
args87(6).Name = "SearchItem.Content"
args87(6).Value = false
args87(7).Name = "SearchItem.AsianOptions"
args87(7).Value = false
args87(8).Name = "SearchItem.AlgorithmType"
args87(8).Value = 0
args87(9).Name = "SearchItem.SearchFlags"
args87(9).Value = 65536
args87(10).Name = "SearchItem.SearchString"
args87(10).Value = "Portuo Riko"
args87(11).Name = "SearchItem.ReplaceString"
args87(11).Value = "Porto-Riko"
args87(12).Name = "SearchItem.Locale"
args87(12).Value = 255
args87(13).Name = "SearchItem.ChangedChars"
args87(13).Value = 2
args87(14).Name = "SearchItem.DeletedChars"
args87(14).Value = 2
args87(15).Name = "SearchItem.InsertedChars"
args87(15).Value = 2
args87(16).Name = "SearchItem.TransliterateFlags"
args87(16).Value = 1280
args87(17).Name = "SearchItem.Command"
args87(17).Value = 3
args87(18).Name = "Quiet"
args87(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args87()

rem ----------------------------------------------------------------------
dim args88(18) as new com.sun.star.beans.PropertyValue
args88(0).Name = "SearchItem.StyleFamily"
args88(0).Value = 2
args88(1).Name = "SearchItem.CellType"
args88(1).Value = 0
args88(2).Name = "SearchItem.RowDirection"
args88(2).Value = true
args88(3).Name = "SearchItem.AllTables"
args88(3).Value = false
args88(4).Name = "SearchItem.Backward"
args88(4).Value = false
args88(5).Name = "SearchItem.Pattern"
args88(5).Value = false
args88(6).Name = "SearchItem.Content"
args88(6).Value = false
args88(7).Name = "SearchItem.AsianOptions"
args88(7).Value = false
args88(8).Name = "SearchItem.AlgorithmType"
args88(8).Value = 0
args88(9).Name = "SearchItem.SearchFlags"
args88(9).Value = 65536
args88(10).Name = "SearchItem.SearchString"
args88(10).Value = "aktoro di cinemo"
args88(11).Name = "SearchItem.ReplaceString"
args88(11).Value = "cinem-aktoro"
args88(12).Name = "SearchItem.Locale"
args88(12).Value = 255
args88(13).Name = "SearchItem.ChangedChars"
args88(13).Value = 2
args88(14).Name = "SearchItem.DeletedChars"
args88(14).Value = 2
args88(15).Name = "SearchItem.InsertedChars"
args88(15).Value = 2
args88(16).Name = "SearchItem.TransliterateFlags"
args88(16).Value = 1280
args88(17).Name = "SearchItem.Command"
args88(17).Value = 3
args88(18).Name = "Quiet"
args88(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args88()

rem ----------------------------------------------------------------------
dim args89(18) as new com.sun.star.beans.PropertyValue
args89(0).Name = "SearchItem.StyleFamily"
args89(0).Value = 2
args89(1).Name = "SearchItem.CellType"
args89(1).Value = 0
args89(2).Name = "SearchItem.RowDirection"
args89(2).Value = true
args89(3).Name = "SearchItem.AllTables"
args89(3).Value = false
args89(4).Name = "SearchItem.Backward"
args89(4).Value = false
args89(5).Name = "SearchItem.Pattern"
args89(5).Value = false
args89(6).Name = "SearchItem.Content"
args89(6).Value = false
args89(7).Name = "SearchItem.AsianOptions"
args89(7).Value = false
args89(8).Name = "SearchItem.AlgorithmType"
args89(8).Value = 0
args89(9).Name = "SearchItem.SearchFlags"
args89(9).Value = 65536
args89(10).Name = "SearchItem.SearchString"
args89(10).Value = "Nijeria"
args89(11).Name = "SearchItem.ReplaceString"
args89(11).Value = "Nigeria"
args89(12).Name = "SearchItem.Locale"
args89(12).Value = 255
args89(13).Name = "SearchItem.ChangedChars"
args89(13).Value = 2
args89(14).Name = "SearchItem.DeletedChars"
args89(14).Value = 2
args89(15).Name = "SearchItem.InsertedChars"
args89(15).Value = 2
args89(16).Name = "SearchItem.TransliterateFlags"
args89(16).Value = 1280
args89(17).Name = "SearchItem.Command"
args89(17).Value = 3
args89(18).Name = "Quiet"
args89(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args89()

rem ----------------------------------------------------------------------
dim args90(18) as new com.sun.star.beans.PropertyValue
args90(0).Name = "SearchItem.StyleFamily"
args90(0).Value = 2
args90(1).Name = "SearchItem.CellType"
args90(1).Value = 0
args90(2).Name = "SearchItem.RowDirection"
args90(2).Value = true
args90(3).Name = "SearchItem.AllTables"
args90(3).Value = false
args90(4).Name = "SearchItem.Backward"
args90(4).Value = false
args90(5).Name = "SearchItem.Pattern"
args90(5).Value = false
args90(6).Name = "SearchItem.Content"
args90(6).Value = false
args90(7).Name = "SearchItem.AsianOptions"
args90(7).Value = false
args90(8).Name = "SearchItem.AlgorithmType"
args90(8).Value = 0
args90(9).Name = "SearchItem.SearchFlags"
args90(9).Value = 65536
args90(10).Name = "SearchItem.SearchString"
args90(10).Value = "Islama "
args90(11).Name = "SearchItem.ReplaceString"
args90(11).Value = "Islamala "
args90(12).Name = "SearchItem.Locale"
args90(12).Value = 255
args90(13).Name = "SearchItem.ChangedChars"
args90(13).Value = 2
args90(14).Name = "SearchItem.DeletedChars"
args90(14).Value = 2
args90(15).Name = "SearchItem.InsertedChars"
args90(15).Value = 2
args90(16).Name = "SearchItem.TransliterateFlags"
args90(16).Value = 1280
args90(17).Name = "SearchItem.Command"
args90(17).Value = 3
args90(18).Name = "Quiet"
args90(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args90()

rem ----------------------------------------------------------------------
dim args91(18) as new com.sun.star.beans.PropertyValue
args91(0).Name = "SearchItem.StyleFamily"
args91(0).Value = 2
args91(1).Name = "SearchItem.CellType"
args91(1).Value = 0
args91(2).Name = "SearchItem.RowDirection"
args91(2).Value = true
args91(3).Name = "SearchItem.AllTables"
args91(3).Value = false
args91(4).Name = "SearchItem.Backward"
args91(4).Value = false
args91(5).Name = "SearchItem.Pattern"
args91(5).Value = false
args91(6).Name = "SearchItem.Content"
args91(6).Value = false
args91(7).Name = "SearchItem.AsianOptions"
args91(7).Value = false
args91(8).Name = "SearchItem.AlgorithmType"
args91(8).Value = 0
args91(9).Name = "SearchItem.SearchFlags"
args91(9).Value = 65536
args91(10).Name = "SearchItem.SearchString"
args91(10).Value = "kontrapapo"
args91(11).Name = "SearchItem.ReplaceString"
args91(11).Value = "kontrepapo"
args91(12).Name = "SearchItem.Locale"
args91(12).Value = 255
args91(13).Name = "SearchItem.ChangedChars"
args91(13).Value = 2
args91(14).Name = "SearchItem.DeletedChars"
args91(14).Value = 2
args91(15).Name = "SearchItem.InsertedChars"
args91(15).Value = 2
args91(16).Name = "SearchItem.TransliterateFlags"
args91(16).Value = 1280
args91(17).Name = "SearchItem.Command"
args91(17).Value = 3
args91(18).Name = "Quiet"
args91(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args91()

rem ----------------------------------------------------------------------
dim args92(18) as new com.sun.star.beans.PropertyValue
args92(0).Name = "SearchItem.StyleFamily"
args92(0).Value = 2
args92(1).Name = "SearchItem.CellType"
args92(1).Value = 0
args92(2).Name = "SearchItem.RowDirection"
args92(2).Value = true
args92(3).Name = "SearchItem.AllTables"
args92(3).Value = false
args92(4).Name = "SearchItem.Backward"
args92(4).Value = false
args92(5).Name = "SearchItem.Pattern"
args92(5).Value = false
args92(6).Name = "SearchItem.Content"
args92(6).Value = false
args92(7).Name = "SearchItem.AsianOptions"
args92(7).Value = false
args92(8).Name = "SearchItem.AlgorithmType"
args92(8).Value = 0
args92(9).Name = "SearchItem.SearchFlags"
args92(9).Value = 65536
args92(10).Name = "SearchItem.SearchString"
args92(10).Value = "lojanti"
args92(11).Name = "SearchItem.ReplaceString"
args92(11).Value = "habitanti"
args92(12).Name = "SearchItem.Locale"
args92(12).Value = 255
args92(13).Name = "SearchItem.ChangedChars"
args92(13).Value = 2
args92(14).Name = "SearchItem.DeletedChars"
args92(14).Value = 2
args92(15).Name = "SearchItem.InsertedChars"
args92(15).Value = 2
args92(16).Name = "SearchItem.TransliterateFlags"
args92(16).Value = 1280
args92(17).Name = "SearchItem.Command"
args92(17).Value = 3
args92(18).Name = "Quiet"
args92(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args92()

rem ----------------------------------------------------------------------
dim args93(18) as new com.sun.star.beans.PropertyValue
args93(0).Name = "SearchItem.StyleFamily"
args93(0).Value = 2
args93(1).Name = "SearchItem.CellType"
args93(1).Value = 0
args93(2).Name = "SearchItem.RowDirection"
args93(2).Value = true
args93(3).Name = "SearchItem.AllTables"
args93(3).Value = false
args93(4).Name = "SearchItem.Backward"
args93(4).Value = false
args93(5).Name = "SearchItem.Pattern"
args93(5).Value = false
args93(6).Name = "SearchItem.Content"
args93(6).Value = false
args93(7).Name = "SearchItem.AsianOptions"
args93(7).Value = false
args93(8).Name = "SearchItem.AlgorithmType"
args93(8).Value = 0
args93(9).Name = "SearchItem.SearchFlags"
args93(9).Value = 65536
args93(10).Name = "SearchItem.SearchString"
args93(10).Value = "esis habitita"
args93(11).Name = "SearchItem.ReplaceString"
args93(11).Value = "habitesis"
args93(12).Name = "SearchItem.Locale"
args93(12).Value = 255
args93(13).Name = "SearchItem.ChangedChars"
args93(13).Value = 2
args93(14).Name = "SearchItem.DeletedChars"
args93(14).Value = 2
args93(15).Name = "SearchItem.InsertedChars"
args93(15).Value = 2
args93(16).Name = "SearchItem.TransliterateFlags"
args93(16).Value = 1280
args93(17).Name = "SearchItem.Command"
args93(17).Value = 3
args93(18).Name = "Quiet"
args93(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args93()

rem ----------------------------------------------------------------------
dim args94(18) as new com.sun.star.beans.PropertyValue
args94(0).Name = "SearchItem.StyleFamily"
args94(0).Value = 2
args94(1).Name = "SearchItem.CellType"
args94(1).Value = 0
args94(2).Name = "SearchItem.RowDirection"
args94(2).Value = true
args94(3).Name = "SearchItem.AllTables"
args94(3).Value = false
args94(4).Name = "SearchItem.Backward"
args94(4).Value = false
args94(5).Name = "SearchItem.Pattern"
args94(5).Value = false
args94(6).Name = "SearchItem.Content"
args94(6).Value = false
args94(7).Name = "SearchItem.AsianOptions"
args94(7).Value = false
args94(8).Name = "SearchItem.AlgorithmType"
args94(8).Value = 0
args94(9).Name = "SearchItem.SearchFlags"
args94(9).Value = 65536
args94(10).Name = "SearchItem.SearchString"
args94(10).Value = "Latin Amerika"
args94(11).Name = "SearchItem.ReplaceString"
args94(11).Value = "Latinida Amerika"
args94(12).Name = "SearchItem.Locale"
args94(12).Value = 255
args94(13).Name = "SearchItem.ChangedChars"
args94(13).Value = 2
args94(14).Name = "SearchItem.DeletedChars"
args94(14).Value = 2
args94(15).Name = "SearchItem.InsertedChars"
args94(15).Value = 2
args94(16).Name = "SearchItem.TransliterateFlags"
args94(16).Value = 1280
args94(17).Name = "SearchItem.Command"
args94(17).Value = 3
args94(18).Name = "Quiet"
args94(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args94()

rem ----------------------------------------------------------------------
dim args95(18) as new com.sun.star.beans.PropertyValue
args95(0).Name = "SearchItem.StyleFamily"
args95(0).Value = 2
args95(1).Name = "SearchItem.CellType"
args95(1).Value = 0
args95(2).Name = "SearchItem.RowDirection"
args95(2).Value = true
args95(3).Name = "SearchItem.AllTables"
args95(3).Value = false
args95(4).Name = "SearchItem.Backward"
args95(4).Value = false
args95(5).Name = "SearchItem.Pattern"
args95(5).Value = false
args95(6).Name = "SearchItem.Content"
args95(6).Value = false
args95(7).Name = "SearchItem.AsianOptions"
args95(7).Value = false
args95(8).Name = "SearchItem.AlgorithmType"
args95(8).Value = 0
args95(9).Name = "SearchItem.SearchFlags"
args95(9).Value = 65536
args95(10).Name = "SearchItem.SearchString"
args95(10).Value = "esas mortigita"
args95(11).Name = "SearchItem.ReplaceString"
args95(11).Value = "mortigesas"
args95(12).Name = "SearchItem.Locale"
args95(12).Value = 255
args95(13).Name = "SearchItem.ChangedChars"
args95(13).Value = 2
args95(14).Name = "SearchItem.DeletedChars"
args95(14).Value = 2
args95(15).Name = "SearchItem.InsertedChars"
args95(15).Value = 2
args95(16).Name = "SearchItem.TransliterateFlags"
args95(16).Value = 1280
args95(17).Name = "SearchItem.Command"
args95(17).Value = 3
args95(18).Name = "Quiet"
args95(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args95()

rem ----------------------------------------------------------------------
dim args96(18) as new com.sun.star.beans.PropertyValue
args96(0).Name = "SearchItem.StyleFamily"
args96(0).Value = 2
args96(1).Name = "SearchItem.CellType"
args96(1).Value = 0
args96(2).Name = "SearchItem.RowDirection"
args96(2).Value = true
args96(3).Name = "SearchItem.AllTables"
args96(3).Value = false
args96(4).Name = "SearchItem.Backward"
args96(4).Value = false
args96(5).Name = "SearchItem.Pattern"
args96(5).Value = false
args96(6).Name = "SearchItem.Content"
args96(6).Value = false
args96(7).Name = "SearchItem.AsianOptions"
args96(7).Value = false
args96(8).Name = "SearchItem.AlgorithmType"
args96(8).Value = 0
args96(9).Name = "SearchItem.SearchFlags"
args96(9).Value = 65536
args96(10).Name = "SearchItem.SearchString"
args96(10).Value = "laureato en"
args96(11).Name = "SearchItem.ReplaceString"
args96(11).Value = "laureato pri"
args96(12).Name = "SearchItem.Locale"
args96(12).Value = 255
args96(13).Name = "SearchItem.ChangedChars"
args96(13).Value = 2
args96(14).Name = "SearchItem.DeletedChars"
args96(14).Value = 2
args96(15).Name = "SearchItem.InsertedChars"
args96(15).Value = 2
args96(16).Name = "SearchItem.TransliterateFlags"
args96(16).Value = 1280
args96(17).Name = "SearchItem.Command"
args96(17).Value = 3
args96(18).Name = "Quiet"
args96(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args96()

rem ----------------------------------------------------------------------
dim args97(18) as new com.sun.star.beans.PropertyValue
args97(0).Name = "SearchItem.StyleFamily"
args97(0).Value = 2
args97(1).Name = "SearchItem.CellType"
args97(1).Value = 0
args97(2).Name = "SearchItem.RowDirection"
args97(2).Value = true
args97(3).Name = "SearchItem.AllTables"
args97(3).Value = false
args97(4).Name = "SearchItem.Backward"
args97(4).Value = false
args97(5).Name = "SearchItem.Pattern"
args97(5).Value = false
args97(6).Name = "SearchItem.Content"
args97(6).Value = false
args97(7).Name = "SearchItem.AsianOptions"
args97(7).Value = false
args97(8).Name = "SearchItem.AlgorithmType"
args97(8).Value = 0
args97(9).Name = "SearchItem.SearchFlags"
args97(9).Value = 65536
args97(10).Name = "SearchItem.SearchString"
args97(10).Value = "esis rielektita"
args97(11).Name = "SearchItem.ReplaceString"
args97(11).Value = "rielektesis"
args97(12).Name = "SearchItem.Locale"
args97(12).Value = 255
args97(13).Name = "SearchItem.ChangedChars"
args97(13).Value = 2
args97(14).Name = "SearchItem.DeletedChars"
args97(14).Value = 2
args97(15).Name = "SearchItem.InsertedChars"
args97(15).Value = 2
args97(16).Name = "SearchItem.TransliterateFlags"
args97(16).Value = 1280
args97(17).Name = "SearchItem.Command"
args97(17).Value = 3
args97(18).Name = "Quiet"
args97(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args97()

rem ----------------------------------------------------------------------
dim args98(18) as new com.sun.star.beans.PropertyValue
args98(0).Name = "SearchItem.StyleFamily"
args98(0).Value = 2
args98(1).Name = "SearchItem.CellType"
args98(1).Value = 0
args98(2).Name = "SearchItem.RowDirection"
args98(2).Value = true
args98(3).Name = "SearchItem.AllTables"
args98(3).Value = false
args98(4).Name = "SearchItem.Backward"
args98(4).Value = false
args98(5).Name = "SearchItem.Pattern"
args98(5).Value = false
args98(6).Name = "SearchItem.Content"
args98(6).Value = false
args98(7).Name = "SearchItem.AsianOptions"
args98(7).Value = false
args98(8).Name = "SearchItem.AlgorithmType"
args98(8).Value = 0
args98(9).Name = "SearchItem.SearchFlags"
args98(9).Value = 65536
args98(10).Name = "SearchItem.SearchString"
args98(10).Value = "esis habitita"
args98(11).Name = "SearchItem.ReplaceString"
args98(11).Value = "habitesis"
args98(12).Name = "SearchItem.Locale"
args98(12).Value = 255
args98(13).Name = "SearchItem.ChangedChars"
args98(13).Value = 2
args98(14).Name = "SearchItem.DeletedChars"
args98(14).Value = 2
args98(15).Name = "SearchItem.InsertedChars"
args98(15).Value = 2
args98(16).Name = "SearchItem.TransliterateFlags"
args98(16).Value = 1280
args98(17).Name = "SearchItem.Command"
args98(17).Value = 3
args98(18).Name = "Quiet"
args98(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args98()

rem ----------------------------------------------------------------------
dim args99(18) as new com.sun.star.beans.PropertyValue
args99(0).Name = "SearchItem.StyleFamily"
args99(0).Value = 2
args99(1).Name = "SearchItem.CellType"
args99(1).Value = 0
args99(2).Name = "SearchItem.RowDirection"
args99(2).Value = true
args99(3).Name = "SearchItem.AllTables"
args99(3).Value = false
args99(4).Name = "SearchItem.Backward"
args99(4).Value = false
args99(5).Name = "SearchItem.Pattern"
args99(5).Value = false
args99(6).Name = "SearchItem.Content"
args99(6).Value = false
args99(7).Name = "SearchItem.AsianOptions"
args99(7).Value = false
args99(8).Name = "SearchItem.AlgorithmType"
args99(8).Value = 0
args99(9).Name = "SearchItem.SearchFlags"
args99(9).Value = 65536
args99(10).Name = "SearchItem.SearchString"
args99(10).Value = " guvernio "
args99(11).Name = "SearchItem.ReplaceString"
args99(11).Value = " guvernerio "
args99(12).Name = "SearchItem.Locale"
args99(12).Value = 255
args99(13).Name = "SearchItem.ChangedChars"
args99(13).Value = 2
args99(14).Name = "SearchItem.DeletedChars"
args99(14).Value = 2
args99(15).Name = "SearchItem.InsertedChars"
args99(15).Value = 2
args99(16).Name = "SearchItem.TransliterateFlags"
args99(16).Value = 1280
args99(17).Name = "SearchItem.Command"
args99(17).Value = 3
args99(18).Name = "Quiet"
args99(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args99()

rem ----------------------------------------------------------------------
dim args100(18) as new com.sun.star.beans.PropertyValue
args100(0).Name = "SearchItem.StyleFamily"
args100(0).Value = 2
args100(1).Name = "SearchItem.CellType"
args100(1).Value = 0
args100(2).Name = "SearchItem.RowDirection"
args100(2).Value = true
args100(3).Name = "SearchItem.AllTables"
args100(3).Value = false
args100(4).Name = "SearchItem.Backward"
args100(4).Value = false
args100(5).Name = "SearchItem.Pattern"
args100(5).Value = false
args100(6).Name = "SearchItem.Content"
args100(6).Value = false
args100(7).Name = "SearchItem.AsianOptions"
args100(7).Value = false
args100(8).Name = "SearchItem.AlgorithmType"
args100(8).Value = 0
args100(9).Name = "SearchItem.SearchFlags"
args100(9).Value = 65536
args100(10).Name = "SearchItem.SearchString"
args100(10).Value = "esis adoptita"
args100(11).Name = "SearchItem.ReplaceString"
args100(11).Value = "adoptesis"
args100(12).Name = "SearchItem.Locale"
args100(12).Value = 255
args100(13).Name = "SearchItem.ChangedChars"
args100(13).Value = 2
args100(14).Name = "SearchItem.DeletedChars"
args100(14).Value = 2
args100(15).Name = "SearchItem.InsertedChars"
args100(15).Value = 2
args100(16).Name = "SearchItem.TransliterateFlags"
args100(16).Value = 1280
args100(17).Name = "SearchItem.Command"
args100(17).Value = 3
args100(18).Name = "Quiet"
args100(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args100()

rem ----------------------------------------------------------------------
dim args101(18) as new com.sun.star.beans.PropertyValue
args101(0).Name = "SearchItem.StyleFamily"
args101(0).Value = 2
args101(1).Name = "SearchItem.CellType"
args101(1).Value = 0
args101(2).Name = "SearchItem.RowDirection"
args101(2).Value = true
args101(3).Name = "SearchItem.AllTables"
args101(3).Value = false
args101(4).Name = "SearchItem.Backward"
args101(4).Value = false
args101(5).Name = "SearchItem.Pattern"
args101(5).Value = false
args101(6).Name = "SearchItem.Content"
args101(6).Value = false
args101(7).Name = "SearchItem.AsianOptions"
args101(7).Value = false
args101(8).Name = "SearchItem.AlgorithmType"
args101(8).Value = 0
args101(9).Name = "SearchItem.SearchFlags"
args101(9).Value = 65536
args101(10).Name = "SearchItem.SearchString"
args101(10).Value = " ye [["
args101(11).Name = "SearchItem.ReplaceString"
args101(11).Value = " en [["
args101(12).Name = "SearchItem.Locale"
args101(12).Value = 255
args101(13).Name = "SearchItem.ChangedChars"
args101(13).Value = 2
args101(14).Name = "SearchItem.DeletedChars"
args101(14).Value = 2
args101(15).Name = "SearchItem.InsertedChars"
args101(15).Value = 2
args101(16).Name = "SearchItem.TransliterateFlags"
args101(16).Value = 1024
args101(17).Name = "SearchItem.Command"
args101(17).Value = 3
args101(18).Name = "Quiet"
args101(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args101())

rem ----------------------------------------------------------------------
dim args102(18) as new com.sun.star.beans.PropertyValue
args102(0).Name = "SearchItem.StyleFamily"
args102(0).Value = 2
args102(1).Name = "SearchItem.CellType"
args102(1).Value = 0
args102(2).Name = "SearchItem.RowDirection"
args102(2).Value = true
args102(3).Name = "SearchItem.AllTables"
args102(3).Value = false
args102(4).Name = "SearchItem.Backward"
args102(4).Value = false
args102(5).Name = "SearchItem.Pattern"
args102(5).Value = false
args102(6).Name = "SearchItem.Content"
args102(6).Value = false
args102(7).Name = "SearchItem.AsianOptions"
args102(7).Value = false
args102(8).Name = "SearchItem.AlgorithmType"
args102(8).Value = 0
args102(9).Name = "SearchItem.SearchFlags"
args102(9).Value = 65536
args102(10).Name = "SearchItem.SearchString"
args102(10).Value = " Ye "
args102(11).Name = "SearchItem.ReplaceString"
args102(11).Value = " En "
args102(12).Name = "SearchItem.Locale"
args102(12).Value = 255
args102(13).Name = "SearchItem.ChangedChars"
args102(13).Value = 2
args102(14).Name = "SearchItem.DeletedChars"
args102(14).Value = 2
args102(15).Name = "SearchItem.InsertedChars"
args102(15).Value = 2
args102(16).Name = "SearchItem.TransliterateFlags"
args102(16).Value = 1024
args102(17).Name = "SearchItem.Command"
args102(17).Value = 3
args102(18).Name = "Quiet"
args102(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args102())

rem ----------------------------------------------------------------------
dim args103(18) as new com.sun.star.beans.PropertyValue
args103(0).Name = "SearchItem.StyleFamily"
args103(0).Value = 2
args103(1).Name = "SearchItem.CellType"
args103(1).Value = 0
args103(2).Name = "SearchItem.RowDirection"
args103(2).Value = true
args103(3).Name = "SearchItem.AllTables"
args103(3).Value = false
args103(4).Name = "SearchItem.Backward"
args103(4).Value = false
args103(5).Name = "SearchItem.Pattern"
args103(5).Value = false
args103(6).Name = "SearchItem.Content"
args103(6).Value = false
args103(7).Name = "SearchItem.AsianOptions"
args103(7).Value = false
args103(8).Name = "SearchItem.AlgorithmType"
args103(8).Value = 0
args103(9).Name = "SearchItem.SearchFlags"
args103(9).Value = 65536
args103(10).Name = "SearchItem.SearchString"
args103(10).Value = "esis aprobita"
args103(11).Name = "SearchItem.ReplaceString"
args103(11).Value = "aprobesis"
args103(12).Name = "SearchItem.Locale"
args103(12).Value = 255
args103(13).Name = "SearchItem.ChangedChars"
args103(13).Value = 2
args103(14).Name = "SearchItem.DeletedChars"
args103(14).Value = 2
args103(15).Name = "SearchItem.InsertedChars"
args103(15).Value = 2
args103(16).Name = "SearchItem.TransliterateFlags"
args103(16).Value = 1024
args103(17).Name = "SearchItem.Command"
args103(17).Value = 3
args103(18).Name = "Quiet"
args103(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args103())

rem ----------------------------------------------------------------------
dim args104(18) as new com.sun.star.beans.PropertyValue
args104(0).Name = "SearchItem.StyleFamily"
args104(0).Value = 2
args104(1).Name = "SearchItem.CellType"
args104(1).Value = 0
args104(2).Name = "SearchItem.RowDirection"
args104(2).Value = true
args104(3).Name = "SearchItem.AllTables"
args104(3).Value = false
args104(4).Name = "SearchItem.Backward"
args104(4).Value = false
args104(5).Name = "SearchItem.Pattern"
args104(5).Value = false
args104(6).Name = "SearchItem.Content"
args104(6).Value = false
args104(7).Name = "SearchItem.AsianOptions"
args104(7).Value = false
args104(8).Name = "SearchItem.AlgorithmType"
args104(8).Value = 0
args104(9).Name = "SearchItem.SearchFlags"
args104(9).Value = 65536
args104(10).Name = "SearchItem.SearchString"
args104(10).Value = "nedependes"
args104(11).Name = "SearchItem.ReplaceString"
args104(11).Value = "nedepend"
args104(12).Name = "SearchItem.Locale"
args104(12).Value = 255
args104(13).Name = "SearchItem.ChangedChars"
args104(13).Value = 2
args104(14).Name = "SearchItem.DeletedChars"
args104(14).Value = 2
args104(15).Name = "SearchItem.InsertedChars"
args104(15).Value = 2
args104(16).Name = "SearchItem.TransliterateFlags"
args104(16).Value = 1024
args104(17).Name = "SearchItem.Command"
args104(17).Value = 3
args104(18).Name = "Quiet"
args104(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args104())

rem ----------------------------------------------------------------------
dim args105(18) as new com.sun.star.beans.PropertyValue
args105(0).Name = "SearchItem.StyleFamily"
args105(0).Value = 2
args105(1).Name = "SearchItem.CellType"
args105(1).Value = 0
args105(2).Name = "SearchItem.RowDirection"
args105(2).Value = true
args105(3).Name = "SearchItem.AllTables"
args105(3).Value = false
args105(4).Name = "SearchItem.Backward"
args105(4).Value = false
args105(5).Name = "SearchItem.Pattern"
args105(5).Value = false
args105(6).Name = "SearchItem.Content"
args105(6).Value = false
args105(7).Name = "SearchItem.AsianOptions"
args105(7).Value = false
args105(8).Name = "SearchItem.AlgorithmType"
args105(8).Value = 0
args105(9).Name = "SearchItem.SearchFlags"
args105(9).Value = 65536
args105(10).Name = "SearchItem.SearchString"
args105(10).Value = "militas kontra"
args105(11).Name = "SearchItem.ReplaceString"
args105(11).Value = "militas kontre"
args105(12).Name = "SearchItem.Locale"
args105(12).Value = 255
args105(13).Name = "SearchItem.ChangedChars"
args105(13).Value = 2
args105(14).Name = "SearchItem.DeletedChars"
args105(14).Value = 2
args105(15).Name = "SearchItem.InsertedChars"
args105(15).Value = 2
args105(16).Name = "SearchItem.TransliterateFlags"
args105(16).Value = 1024
args105(17).Name = "SearchItem.Command"
args105(17).Value = 3
args105(18).Name = "Quiet"
args105(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args105())

rem ----------------------------------------------------------------------
dim args106(18) as new com.sun.star.beans.PropertyValue
args106(0).Name = "SearchItem.StyleFamily"
args106(0).Value = 2
args106(1).Name = "SearchItem.CellType"
args106(1).Value = 0
args106(2).Name = "SearchItem.RowDirection"
args106(2).Value = true
args106(3).Name = "SearchItem.AllTables"
args106(3).Value = false
args106(4).Name = "SearchItem.Backward"
args106(4).Value = false
args106(5).Name = "SearchItem.Pattern"
args106(5).Value = false
args106(6).Name = "SearchItem.Content"
args106(6).Value = false
args106(7).Name = "SearchItem.AsianOptions"
args106(7).Value = false
args106(8).Name = "SearchItem.AlgorithmType"
args106(8).Value = 0
args106(9).Name = "SearchItem.SearchFlags"
args106(9).Value = 65536
args106(10).Name = "SearchItem.SearchString"
args106(10).Value = "fotografio"
args106(11).Name = "SearchItem.ReplaceString"
args106(11).Value = "fotografo"
args106(12).Name = "SearchItem.Locale"
args106(12).Value = 255
args106(13).Name = "SearchItem.ChangedChars"
args106(13).Value = 2
args106(14).Name = "SearchItem.DeletedChars"
args106(14).Value = 2
args106(15).Name = "SearchItem.InsertedChars"
args106(15).Value = 2
args106(16).Name = "SearchItem.TransliterateFlags"
args106(16).Value = 1024
args106(17).Name = "SearchItem.Command"
args106(17).Value = 3
args106(18).Name = "Quiet"
args106(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args106())

rem ----------------------------------------------------------------------
dim args108(18) as new com.sun.star.beans.PropertyValue
args108(0).Name = "SearchItem.StyleFamily"
args108(0).Value = 2
args108(1).Name = "SearchItem.CellType"
args108(1).Value = 0
args108(2).Name = "SearchItem.RowDirection"
args108(2).Value = true
args108(3).Name = "SearchItem.AllTables"
args108(3).Value = false
args108(4).Name = "SearchItem.Backward"
args108(4).Value = false
args108(5).Name = "SearchItem.Pattern"
args108(5).Value = false
args108(6).Name = "SearchItem.Content"
args108(6).Value = false
args108(7).Name = "SearchItem.AsianOptions"
args108(7).Value = false
args108(8).Name = "SearchItem.AlgorithmType"
args108(8).Value = 0
args108(9).Name = "SearchItem.SearchFlags"
args108(9).Value = 65536
args108(10).Name = "SearchItem.SearchString"
args108(10).Value = "militis kontra"
args108(11).Name = "SearchItem.ReplaceString"
args108(11).Value = "militis kontre"
args108(12).Name = "SearchItem.Locale"
args108(12).Value = 255
args108(13).Name = "SearchItem.ChangedChars"
args108(13).Value = 2
args108(14).Name = "SearchItem.DeletedChars"
args108(14).Value = 2
args108(15).Name = "SearchItem.InsertedChars"
args108(15).Value = 2
args108(16).Name = "SearchItem.TransliterateFlags"
args108(16).Value = 1024
args108(17).Name = "SearchItem.Command"
args108(17).Value = 3
args108(18).Name = "Quiet"
args108(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args108())

rem ----------------------------------------------------------------------
dim args109(18) as new com.sun.star.beans.PropertyValue
args109(0).Name = "SearchItem.StyleFamily"
args109(0).Value = 2
args109(1).Name = "SearchItem.CellType"
args109(1).Value = 0
args109(2).Name = "SearchItem.RowDirection"
args109(2).Value = true
args109(3).Name = "SearchItem.AllTables"
args109(3).Value = false
args109(4).Name = "SearchItem.Backward"
args109(4).Value = false
args109(5).Name = "SearchItem.Pattern"
args109(5).Value = false
args109(6).Name = "SearchItem.Content"
args109(6).Value = false
args109(7).Name = "SearchItem.AsianOptions"
args109(7).Value = false
args109(8).Name = "SearchItem.AlgorithmType"
args109(8).Value = 0
args109(9).Name = "SearchItem.SearchFlags"
args109(9).Value = 65536
args109(10).Name = "SearchItem.SearchString"
args109(10).Value = "esis fondita"
args109(11).Name = "SearchItem.ReplaceString"
args109(11).Value = "fondesis"
args109(12).Name = "SearchItem.Locale"
args109(12).Value = 255
args109(13).Name = "SearchItem.ChangedChars"
args109(13).Value = 2
args109(14).Name = "SearchItem.DeletedChars"
args109(14).Value = 2
args109(15).Name = "SearchItem.InsertedChars"
args109(15).Value = 2
args109(16).Name = "SearchItem.TransliterateFlags"
args109(16).Value = 1024
args109(17).Name = "SearchItem.Command"
args109(17).Value = 3
args109(18).Name = "Quiet"
args109(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args109())

rem ----------------------------------------------------------------------
dim args110(18) as new com.sun.star.beans.PropertyValue
args110(0).Name = "SearchItem.StyleFamily"
args110(0).Value = 2
args110(1).Name = "SearchItem.CellType"
args110(1).Value = 0
args110(2).Name = "SearchItem.RowDirection"
args110(2).Value = true
args110(3).Name = "SearchItem.AllTables"
args110(3).Value = false
args110(4).Name = "SearchItem.Backward"
args110(4).Value = false
args110(5).Name = "SearchItem.Pattern"
args110(5).Value = false
args110(6).Name = "SearchItem.Content"
args110(6).Value = false
args110(7).Name = "SearchItem.AsianOptions"
args110(7).Value = false
args110(8).Name = "SearchItem.AlgorithmType"
args110(8).Value = 0
args110(9).Name = "SearchItem.SearchFlags"
args110(9).Value = 65536
args110(10).Name = "SearchItem.SearchString"
args110(10).Value = "esis destruktita"
args110(11).Name = "SearchItem.ReplaceString"
args110(11).Value = "destruktesis"
args110(12).Name = "SearchItem.Locale"
args110(12).Value = 255
args110(13).Name = "SearchItem.ChangedChars"
args110(13).Value = 2
args110(14).Name = "SearchItem.DeletedChars"
args110(14).Value = 2
args110(15).Name = "SearchItem.InsertedChars"
args110(15).Value = 2
args110(16).Name = "SearchItem.TransliterateFlags"
args110(16).Value = 1024
args110(17).Name = "SearchItem.Command"
args110(17).Value = 3
args110(18).Name = "Quiet"
args110(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args110())

rem ----------------------------------------------------------------------
dim args111(18) as new com.sun.star.beans.PropertyValue
args111(0).Name = "SearchItem.StyleFamily"
args111(0).Value = 2
args111(1).Name = "SearchItem.CellType"
args111(1).Value = 0
args111(2).Name = "SearchItem.RowDirection"
args111(2).Value = true
args111(3).Name = "SearchItem.AllTables"
args111(3).Value = false
args111(4).Name = "SearchItem.Backward"
args111(4).Value = false
args111(5).Name = "SearchItem.Pattern"
args111(5).Value = false
args111(6).Name = "SearchItem.Content"
args111(6).Value = false
args111(7).Name = "SearchItem.AsianOptions"
args111(7).Value = false
args111(8).Name = "SearchItem.AlgorithmType"
args111(8).Value = 0
args111(9).Name = "SearchItem.SearchFlags"
args111(9).Value = 65536
args111(10).Name = "SearchItem.SearchString"
args111(10).Value = ":''Videz anke: [["
args111(11).Name = "SearchItem.ReplaceString"
args111(11).Value = "{{PA|"
args111(12).Name = "SearchItem.Locale"
args111(12).Value = 255
args111(13).Name = "SearchItem.ChangedChars"
args111(13).Value = 2
args111(14).Name = "SearchItem.DeletedChars"
args111(14).Value = 2
args111(15).Name = "SearchItem.InsertedChars"
args111(15).Value = 2
args111(16).Name = "SearchItem.TransliterateFlags"
args111(16).Value = 1024
args111(17).Name = "SearchItem.Command"
args111(17).Value = 3
args111(18).Name = "Quiet"
args111(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args111())

rem ----------------------------------------------------------------------
dim args112(18) as new com.sun.star.beans.PropertyValue
args112(0).Name = "SearchItem.StyleFamily"
args112(0).Value = 2
args112(1).Name = "SearchItem.CellType"
args112(1).Value = 0
args112(2).Name = "SearchItem.RowDirection"
args112(2).Value = true
args112(3).Name = "SearchItem.AllTables"
args112(3).Value = false
args112(4).Name = "SearchItem.Backward"
args112(4).Value = false
args112(5).Name = "SearchItem.Pattern"
args112(5).Value = false
args112(6).Name = "SearchItem.Content"
args112(6).Value = false
args112(7).Name = "SearchItem.AsianOptions"
args112(7).Value = false
args112(8).Name = "SearchItem.AlgorithmType"
args112(8).Value = 0
args112(9).Name = "SearchItem.SearchFlags"
args112(9).Value = 65536
args112(10).Name = "SearchItem.SearchString"
args112(10).Value = "vaporo-mashin"
args112(11).Name = "SearchItem.ReplaceString"
args112(11).Value = "vaporomashin"
args112(12).Name = "SearchItem.Locale"
args112(12).Value = 255
args112(13).Name = "SearchItem.ChangedChars"
args112(13).Value = 2
args112(14).Name = "SearchItem.DeletedChars"
args112(14).Value = 2
args112(15).Name = "SearchItem.InsertedChars"
args112(15).Value = 2
args112(16).Name = "SearchItem.TransliterateFlags"
args112(16).Value = 1024
args112(17).Name = "SearchItem.Command"
args112(17).Value = 3
args112(18).Name = "Quiet"
args112(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args112())

rem ----------------------------------------------------------------------
dim args113(18) as new com.sun.star.beans.PropertyValue
args113(0).Name = "SearchItem.StyleFamily"
args113(0).Value = 2
args113(1).Name = "SearchItem.CellType"
args113(1).Value = 0
args113(2).Name = "SearchItem.RowDirection"
args113(2).Value = true
args113(3).Name = "SearchItem.AllTables"
args113(3).Value = false
args113(4).Name = "SearchItem.Backward"
args113(4).Value = false
args113(5).Name = "SearchItem.Pattern"
args113(5).Value = false
args113(6).Name = "SearchItem.Content"
args113(6).Value = false
args113(7).Name = "SearchItem.AsianOptions"
args113(7).Value = false
args113(8).Name = "SearchItem.AlgorithmType"
args113(8).Value = 0
args113(9).Name = "SearchItem.SearchFlags"
args113(9).Value = 65536
args113(10).Name = "SearchItem.SearchString"
args113(10).Value = "maxim grava"
args113(11).Name = "SearchItem.ReplaceString"
args113(11).Value = "maxim importanta"
args113(12).Name = "SearchItem.Locale"
args113(12).Value = 255
args113(13).Name = "SearchItem.ChangedChars"
args113(13).Value = 2
args113(14).Name = "SearchItem.DeletedChars"
args113(14).Value = 2
args113(15).Name = "SearchItem.InsertedChars"
args113(15).Value = 2
args113(16).Name = "SearchItem.TransliterateFlags"
args113(16).Value = 1024
args113(17).Name = "SearchItem.Command"
args113(17).Value = 3
args113(18).Name = "Quiet"
args113(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args113())

rem ----------------------------------------------------------------------
dim args114(18) as new com.sun.star.beans.PropertyValue
args114(0).Name = "SearchItem.StyleFamily"
args114(0).Value = 2
args114(1).Name = "SearchItem.CellType"
args114(1).Value = 0
args114(2).Name = "SearchItem.RowDirection"
args114(2).Value = true
args114(3).Name = "SearchItem.AllTables"
args114(3).Value = false
args114(4).Name = "SearchItem.Backward"
args114(4).Value = false
args114(5).Name = "SearchItem.Pattern"
args114(5).Value = false
args114(6).Name = "SearchItem.Content"
args114(6).Value = false
args114(7).Name = "SearchItem.AsianOptions"
args114(7).Value = false
args114(8).Name = "SearchItem.AlgorithmType"
args114(8).Value = 0
args114(9).Name = "SearchItem.SearchFlags"
args114(9).Value = 65536
args114(10).Name = "SearchItem.SearchString"
args114(10).Value = "chefministri di Sri Lanka"
args114(11).Name = "SearchItem.ReplaceString"
args114(11).Value = "chefa ministri di Sri Lanka"
args114(12).Name = "SearchItem.Locale"
args114(12).Value = 255
args114(13).Name = "SearchItem.ChangedChars"
args114(13).Value = 2
args114(14).Name = "SearchItem.DeletedChars"
args114(14).Value = 2
args114(15).Name = "SearchItem.InsertedChars"
args114(15).Value = 2
args114(16).Name = "SearchItem.TransliterateFlags"
args114(16).Value = 1024
args114(17).Name = "SearchItem.Command"
args114(17).Value = 3
args114(18).Name = "Quiet"
args114(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args114())

rem ----------------------------------------------------------------------
dim args115(18) as new com.sun.star.beans.PropertyValue
args115(0).Name = "SearchItem.StyleFamily"
args115(0).Value = 2
args115(1).Name = "SearchItem.CellType"
args115(1).Value = 0
args115(2).Name = "SearchItem.RowDirection"
args115(2).Value = true
args115(3).Name = "SearchItem.AllTables"
args115(3).Value = false
args115(4).Name = "SearchItem.Backward"
args115(4).Value = false
args115(5).Name = "SearchItem.Pattern"
args115(5).Value = false
args115(6).Name = "SearchItem.Content"
args115(6).Value = false
args115(7).Name = "SearchItem.AsianOptions"
args115(7).Value = false
args115(8).Name = "SearchItem.AlgorithmType"
args115(8).Value = 0
args115(9).Name = "SearchItem.SearchFlags"
args115(9).Value = 65536
args115(10).Name = "SearchItem.SearchString"
args115(10).Value = "{{Ciencisto"
args115(11).Name = "SearchItem.ReplaceString"
args115(11).Value = "{{Biografio"
args115(12).Name = "SearchItem.Locale"
args115(12).Value = 255
args115(13).Name = "SearchItem.ChangedChars"
args115(13).Value = 2
args115(14).Name = "SearchItem.DeletedChars"
args115(14).Value = 2
args115(15).Name = "SearchItem.InsertedChars"
args115(15).Value = 2
args115(16).Name = "SearchItem.TransliterateFlags"
args115(16).Value = 1024
args115(17).Name = "SearchItem.Command"
args115(17).Value = 3
args115(18).Name = "Quiet"
args115(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args115())

rem ----------------------------------------------------------------------
dim args116(18) as new com.sun.star.beans.PropertyValue
args116(0).Name = "SearchItem.StyleFamily"
args116(0).Value = 2
args116(1).Name = "SearchItem.CellType"
args116(1).Value = 0
args116(2).Name = "SearchItem.RowDirection"
args116(2).Value = true
args116(3).Name = "SearchItem.AllTables"
args116(3).Value = false
args116(4).Name = "SearchItem.Backward"
args116(4).Value = false
args116(5).Name = "SearchItem.Pattern"
args116(5).Value = false
args116(6).Name = "SearchItem.Content"
args116(6).Value = false
args116(7).Name = "SearchItem.AsianOptions"
args116(7).Value = false
args116(8).Name = "SearchItem.AlgorithmType"
args116(8).Value = 0
args116(9).Name = "SearchItem.SearchFlags"
args116(9).Value = 65536
args116(10).Name = "SearchItem.SearchString"
args116(10).Value = "volleyball"
args116(11).Name = "SearchItem.ReplaceString"
args116(11).Value = "volebalo"
args116(12).Name = "SearchItem.Locale"
args116(12).Value = 255
args116(13).Name = "SearchItem.ChangedChars"
args116(13).Value = 2
args116(14).Name = "SearchItem.DeletedChars"
args116(14).Value = 2
args116(15).Name = "SearchItem.InsertedChars"
args116(15).Value = 2
args116(16).Name = "SearchItem.TransliterateFlags"
args116(16).Value = 1024
args116(17).Name = "SearchItem.Command"
args116(17).Value = 3
args116(18).Name = "Quiet"
args116(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args116())
rem ----------------------------------------------------------------------
dim args117(18) as new com.sun.star.beans.PropertyValue
args117(0).Name = "SearchItem.StyleFamily"
args117(0).Value = 2
args117(1).Name = "SearchItem.CellType"
args117(1).Value = 0
args117(2).Name = "SearchItem.RowDirection"
args117(2).Value = true
args117(3).Name = "SearchItem.AllTables"
args117(3).Value = false
args117(4).Name = "SearchItem.Backward"
args117(4).Value = false
args117(5).Name = "SearchItem.Pattern"
args117(5).Value = false
args117(6).Name = "SearchItem.Content"
args117(6).Value = false
args117(7).Name = "SearchItem.AsianOptions"
args117(7).Value = false
args117(8).Name = "SearchItem.AlgorithmType"
args117(8).Value = 0
args117(9).Name = "SearchItem.SearchFlags"
args117(9).Value = 65536
args117(10).Name = "SearchItem.SearchString"
args117(10).Value = "stadiono"
args117(11).Name = "SearchItem.ReplaceString"
args117(11).Value = "stadio"
args117(12).Name = "SearchItem.Locale"
args117(12).Value = 255
args117(13).Name = "SearchItem.ChangedChars"
args117(13).Value = 2
args117(14).Name = "SearchItem.DeletedChars"
args117(14).Value = 2
args117(15).Name = "SearchItem.InsertedChars"
args117(15).Value = 2
args117(16).Name = "SearchItem.TransliterateFlags"
args117(16).Value = 1024
args117(17).Name = "SearchItem.Command"
args117(17).Value = 3
args117(18).Name = "Quiet"
args117(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args117())

rem ----------------------------------------------------------------------
dim args118(18) as new com.sun.star.beans.PropertyValue
args118(0).Name = "SearchItem.StyleFamily"
args118(0).Value = 2
args118(1).Name = "SearchItem.CellType"
args118(1).Value = 0
args118(2).Name = "SearchItem.RowDirection"
args118(2).Value = true
args118(3).Name = "SearchItem.AllTables"
args118(3).Value = false
args118(4).Name = "SearchItem.Backward"
args118(4).Value = false
args118(5).Name = "SearchItem.Pattern"
args118(5).Value = false
args118(6).Name = "SearchItem.Content"
args118(6).Value = false
args118(7).Name = "SearchItem.AsianOptions"
args118(7).Value = false
args118(8).Name = "SearchItem.AlgorithmType"
args118(8).Value = 0
args118(9).Name = "SearchItem.SearchFlags"
args118(9).Value = 65536
args118(10).Name = "SearchItem.SearchString"
args118(10).Value = "substitucis"
args118(11).Name = "SearchItem.ReplaceString"
args118(11).Value = "remplasis"
args118(12).Name = "SearchItem.Locale"
args118(12).Value = 255
args118(13).Name = "SearchItem.ChangedChars"
args118(13).Value = 2
args118(14).Name = "SearchItem.DeletedChars"
args118(14).Value = 2
args118(15).Name = "SearchItem.InsertedChars"
args118(15).Value = 2
args118(16).Name = "SearchItem.TransliterateFlags"
args118(16).Value = 1024
args118(17).Name = "SearchItem.Command"
args118(17).Value = 3
args118(18).Name = "Quiet"
args118(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args118())

rem ----------------------------------------------------------------------
dim args119(18) as new com.sun.star.beans.PropertyValue
args119(0).Name = "SearchItem.StyleFamily"
args119(0).Value = 2
args119(1).Name = "SearchItem.CellType"
args119(1).Value = 0
args119(2).Name = "SearchItem.RowDirection"
args119(2).Value = true
args119(3).Name = "SearchItem.AllTables"
args119(3).Value = false
args119(4).Name = "SearchItem.Backward"
args119(4).Value = false
args119(5).Name = "SearchItem.Pattern"
args119(5).Value = false
args119(6).Name = "SearchItem.Content"
args119(6).Value = false
args119(7).Name = "SearchItem.AsianOptions"
args119(7).Value = false
args119(8).Name = "SearchItem.AlgorithmType"
args119(8).Value = 0
args119(9).Name = "SearchItem.SearchFlags"
args119(9).Value = 65536
args119(10).Name = "SearchItem.SearchString"
args119(10).Value = "substitucas"
args119(11).Name = "SearchItem.ReplaceString"
args119(11).Value = "remplasas"
args119(12).Name = "SearchItem.Locale"
args119(12).Value = 255
args119(13).Name = "SearchItem.ChangedChars"
args119(13).Value = 2
args119(14).Name = "SearchItem.DeletedChars"
args119(14).Value = 2
args119(15).Name = "SearchItem.InsertedChars"
args119(15).Value = 2
args119(16).Name = "SearchItem.TransliterateFlags"
args119(16).Value = 1024
args119(17).Name = "SearchItem.Command"
args119(17).Value = 3
args119(18).Name = "Quiet"
args119(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args119())

rem ----------------------------------------------------------------------
dim args120(18) as new com.sun.star.beans.PropertyValue
args120(0).Name = "SearchItem.StyleFamily"
args120(0).Value = 2
args120(1).Name = "SearchItem.CellType"
args120(1).Value = 0
args120(2).Name = "SearchItem.RowDirection"
args120(2).Value = true
args120(3).Name = "SearchItem.AllTables"
args120(3).Value = false
args120(4).Name = "SearchItem.Backward"
args120(4).Value = false
args120(5).Name = "SearchItem.Pattern"
args120(5).Value = false
args120(6).Name = "SearchItem.Content"
args120(6).Value = false
args120(7).Name = "SearchItem.AsianOptions"
args120(7).Value = false
args120(8).Name = "SearchItem.AlgorithmType"
args120(8).Value = 0
args120(9).Name = "SearchItem.SearchFlags"
args120(9).Value = 65536
args120(10).Name = "SearchItem.SearchString"
args120(10).Value = "ekonomiala aktivesi"
args120(11).Name = "SearchItem.ReplaceString"
args120(11).Value = "ekonomial agadi"
args120(12).Name = "SearchItem.Locale"
args120(12).Value = 255
args120(13).Name = "SearchItem.ChangedChars"
args120(13).Value = 2
args120(14).Name = "SearchItem.DeletedChars"
args120(14).Value = 2
args120(15).Name = "SearchItem.InsertedChars"
args120(15).Value = 2
args120(16).Name = "SearchItem.TransliterateFlags"
args120(16).Value = 1024
args120(17).Name = "SearchItem.Command"
args120(17).Value = 3
args120(18).Name = "Quiet"
args120(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args120())

rem ----------------------------------------------------------------------
dim args122(18) as new com.sun.star.beans.PropertyValue
args122(0).Name = "SearchItem.StyleFamily"
args122(0).Value = 2
args122(1).Name = "SearchItem.CellType"
args122(1).Value = 0
args122(2).Name = "SearchItem.RowDirection"
args122(2).Value = true
args122(3).Name = "SearchItem.AllTables"
args122(3).Value = false
args122(4).Name = "SearchItem.Backward"
args122(4).Value = false
args122(5).Name = "SearchItem.Pattern"
args122(5).Value = false
args122(6).Name = "SearchItem.Content"
args122(6).Value = false
args122(7).Name = "SearchItem.AsianOptions"
args122(7).Value = false
args122(8).Name = "SearchItem.AlgorithmType"
args122(8).Value = 0
args122(9).Name = "SearchItem.SearchFlags"
args122(9).Value = 65536
args122(10).Name = "SearchItem.SearchString"
args122(10).Value = "religiono"
args122(11).Name = "SearchItem.ReplaceString"
args122(11).Value = "religio"
args122(12).Name = "SearchItem.Locale"
args122(12).Value = 255
args122(13).Name = "SearchItem.ChangedChars"
args122(13).Value = 2
args122(14).Name = "SearchItem.DeletedChars"
args122(14).Value = 2
args122(15).Name = "SearchItem.InsertedChars"
args122(15).Value = 2
args122(16).Name = "SearchItem.TransliterateFlags"
args122(16).Value = 1024
args122(17).Name = "SearchItem.Command"
args122(17).Value = 3
args122(18).Name = "Quiet"
args122(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args122())

rem ----------------------------------------------------------------------
dim args123(18) as new com.sun.star.beans.PropertyValue
args123(0).Name = "SearchItem.StyleFamily"
args123(0).Value = 2
args123(1).Name = "SearchItem.CellType"
args123(1).Value = 0
args123(2).Name = "SearchItem.RowDirection"
args123(2).Value = true
args123(3).Name = "SearchItem.AllTables"
args123(3).Value = false
args123(4).Name = "SearchItem.Backward"
args123(4).Value = false
args123(5).Name = "SearchItem.Pattern"
args123(5).Value = false
args123(6).Name = "SearchItem.Content"
args123(6).Value = false
args123(7).Name = "SearchItem.AsianOptions"
args123(7).Value = false
args123(8).Name = "SearchItem.AlgorithmType"
args123(8).Value = 0
args123(9).Name = "SearchItem.SearchFlags"
args123(9).Value = 65536
args123(10).Name = "SearchItem.SearchString"
args123(10).Value = "formacesas da "
args123(11).Name = "SearchItem.ReplaceString"
args123(11).Value = "konsistas ek "
args123(12).Name = "SearchItem.Locale"
args123(12).Value = 255
args123(13).Name = "SearchItem.ChangedChars"
args123(13).Value = 2
args123(14).Name = "SearchItem.DeletedChars"
args123(14).Value = 2
args123(15).Name = "SearchItem.InsertedChars"
args123(15).Value = 2
args123(16).Name = "SearchItem.TransliterateFlags"
args123(16).Value = 1024
args123(17).Name = "SearchItem.Command"
args123(17).Value = 3
args123(18).Name = "Quiet"
args123(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args123())

rem ----------------------------------------------------------------------
dim args124(18) as new com.sun.star.beans.PropertyValue
args124(0).Name = "SearchItem.StyleFamily"
args124(0).Value = 2
args124(1).Name = "SearchItem.CellType"
args124(1).Value = 0
args124(2).Name = "SearchItem.RowDirection"
args124(2).Value = true
args124(3).Name = "SearchItem.AllTables"
args124(3).Value = false
args124(4).Name = "SearchItem.Backward"
args124(4).Value = false
args124(5).Name = "SearchItem.Pattern"
args124(5).Value = false
args124(6).Name = "SearchItem.Content"
args124(6).Value = false
args124(7).Name = "SearchItem.AsianOptions"
args124(7).Value = false
args124(8).Name = "SearchItem.AlgorithmType"
args124(8).Value = 0
args124(9).Name = "SearchItem.SearchFlags"
args124(9).Value = 65536
args124(10).Name = "SearchItem.SearchString"
args124(10).Value = "inflacio]"
args124(11).Name = "SearchItem.ReplaceString"
args124(11).Value = "inflaciono]"
args124(12).Name = "SearchItem.Locale"
args124(12).Value = 255
args124(13).Name = "SearchItem.ChangedChars"
args124(13).Value = 2
args124(14).Name = "SearchItem.DeletedChars"
args124(14).Value = 2
args124(15).Name = "SearchItem.InsertedChars"
args124(15).Value = 2
args124(16).Name = "SearchItem.TransliterateFlags"
args124(16).Value = 1024
args124(17).Name = "SearchItem.Command"
args124(17).Value = 3
args124(18).Name = "Quiet"
args124(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args124())

rem ----------------------------------------------------------------------
dim args125(18) as new com.sun.star.beans.PropertyValue
args125(0).Name = "SearchItem.StyleFamily"
args125(0).Value = 2
args125(1).Name = "SearchItem.CellType"
args125(1).Value = 0
args125(2).Name = "SearchItem.RowDirection"
args125(2).Value = true
args125(3).Name = "SearchItem.AllTables"
args125(3).Value = false
args125(4).Name = "SearchItem.Backward"
args125(4).Value = false
args125(5).Name = "SearchItem.Pattern"
args125(5).Value = false
args125(6).Name = "SearchItem.Content"
args125(6).Value = false
args125(7).Name = "SearchItem.AsianOptions"
args125(7).Value = false
args125(8).Name = "SearchItem.AlgorithmType"
args125(8).Value = 0
args125(9).Name = "SearchItem.SearchFlags"
args125(9).Value = 65536
args125(10).Name = "SearchItem.SearchString"
args125(10).Value = "Sovietana"
args125(11).Name = "SearchItem.ReplaceString"
args125(11).Value = "Sovietiana"
args125(12).Name = "SearchItem.Locale"
args125(12).Value = 255
args125(13).Name = "SearchItem.ChangedChars"
args125(13).Value = 2
args125(14).Name = "SearchItem.DeletedChars"
args125(14).Value = 2
args125(15).Name = "SearchItem.InsertedChars"
args125(15).Value = 2
args125(16).Name = "SearchItem.TransliterateFlags"
args125(16).Value = 1024
args125(17).Name = "SearchItem.Command"
args125(17).Value = 3
args125(18).Name = "Quiet"
args125(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args125())

rem ----------------------------------------------------------------------
dim args126(18) as new com.sun.star.beans.PropertyValue
args126(0).Name = "SearchItem.StyleFamily"
args126(0).Value = 2
args126(1).Name = "SearchItem.CellType"
args126(1).Value = 0
args126(2).Name = "SearchItem.RowDirection"
args126(2).Value = true
args126(3).Name = "SearchItem.AllTables"
args126(3).Value = false
args126(4).Name = "SearchItem.Backward"
args126(4).Value = false
args126(5).Name = "SearchItem.Pattern"
args126(5).Value = false
args126(6).Name = "SearchItem.Content"
args126(6).Value = false
args126(7).Name = "SearchItem.AsianOptions"
args126(7).Value = false
args126(8).Name = "SearchItem.AlgorithmType"
args126(8).Value = 0
args126(9).Name = "SearchItem.SearchFlags"
args126(9).Value = 65536
args126(10).Name = "SearchItem.SearchString"
args126(10).Value = "mulieri 18 yari"
args126(11).Name = "SearchItem.ReplaceString"
args126(11).Value = "mulieri evante 18 yari"
args126(12).Name = "SearchItem.Locale"
args126(12).Value = 255
args126(13).Name = "SearchItem.ChangedChars"
args126(13).Value = 2
args126(14).Name = "SearchItem.DeletedChars"
args126(14).Value = 2
args126(15).Name = "SearchItem.InsertedChars"
args126(15).Value = 2
args126(16).Name = "SearchItem.TransliterateFlags"
args126(16).Value = 1280
args126(17).Name = "SearchItem.Command"
args126(17).Value = 3
args126(18).Name = "Quiet"
args126(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args126())

rem ----------------------------------------------------------------------
dim args127(18) as new com.sun.star.beans.PropertyValue
args127(0).Name = "SearchItem.StyleFamily"
args127(0).Value = 2
args127(1).Name = "SearchItem.CellType"
args127(1).Value = 0
args127(2).Name = "SearchItem.RowDirection"
args127(2).Value = true
args127(3).Name = "SearchItem.AllTables"
args127(3).Value = false
args127(4).Name = "SearchItem.Backward"
args127(4).Value = false
args127(5).Name = "SearchItem.Pattern"
args127(5).Value = false
args127(6).Name = "SearchItem.Content"
args127(6).Value = false
args127(7).Name = "SearchItem.AsianOptions"
args127(7).Value = false
args127(8).Name = "SearchItem.AlgorithmType"
args127(8).Value = 0
args127(9).Name = "SearchItem.SearchFlags"
args127(9).Value = 65536
args127(10).Name = "SearchItem.SearchString"
args127(10).Value = "romantismo"
args127(11).Name = "SearchItem.ReplaceString"
args127(11).Value = "Romantikismo"
args127(12).Name = "SearchItem.Locale"
args127(12).Value = 255
args127(13).Name = "SearchItem.ChangedChars"
args127(13).Value = 2
args127(14).Name = "SearchItem.DeletedChars"
args127(14).Value = 2
args127(15).Name = "SearchItem.InsertedChars"
args127(15).Value = 2
args127(16).Name = "SearchItem.TransliterateFlags"
args127(16).Value = 1024
args127(17).Name = "SearchItem.Command"
args127(17).Value = 3
args127(18).Name = "Quiet"
args127(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args127())

rem ----------------------------------------------------------------------
dim args128(18) as new com.sun.star.beans.PropertyValue
args128(0).Name = "SearchItem.StyleFamily"
args128(0).Value = 2
args128(1).Name = "SearchItem.CellType"
args128(1).Value = 0
args128(2).Name = "SearchItem.RowDirection"
args128(2).Value = true
args128(3).Name = "SearchItem.AllTables"
args128(3).Value = false
args128(4).Name = "SearchItem.Backward"
args128(4).Value = false
args128(5).Name = "SearchItem.Pattern"
args128(5).Value = false
args128(6).Name = "SearchItem.Content"
args128(6).Value = false
args128(7).Name = "SearchItem.AsianOptions"
args128(7).Value = false
args128(8).Name = "SearchItem.AlgorithmType"
args128(8).Value = 0
args128(9).Name = "SearchItem.SearchFlags"
args128(9).Value = 65536
args128(10).Name = "SearchItem.SearchString"
args128(10).Value = "Nexta"
args128(11).Name = "SearchItem.ReplaceString"
args128(11).Value = "Sequanta"
args128(12).Name = "SearchItem.Locale"
args128(12).Value = 255
args128(13).Name = "SearchItem.ChangedChars"
args128(13).Value = 2
args128(14).Name = "SearchItem.DeletedChars"
args128(14).Value = 2
args128(15).Name = "SearchItem.InsertedChars"
args128(15).Value = 2
args128(16).Name = "SearchItem.TransliterateFlags"
args128(16).Value = 1024
args128(17).Name = "SearchItem.Command"
args128(17).Value = 3
args128(18).Name = "Quiet"
args128(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args128())

rem ----------------------------------------------------------------------
dim args129(18) as new com.sun.star.beans.PropertyValue
args129(0).Name = "SearchItem.StyleFamily"
args129(0).Value = 2
args129(1).Name = "SearchItem.CellType"
args129(1).Value = 0
args129(2).Name = "SearchItem.RowDirection"
args129(2).Value = true
args129(3).Name = "SearchItem.AllTables"
args129(3).Value = false
args129(4).Name = "SearchItem.Backward"
args129(4).Value = false
args129(5).Name = "SearchItem.Pattern"
args129(5).Value = false
args129(6).Name = "SearchItem.Content"
args129(6).Value = false
args129(7).Name = "SearchItem.AsianOptions"
args129(7).Value = false
args129(8).Name = "SearchItem.AlgorithmType"
args129(8).Value = 0
args129(9).Name = "SearchItem.SearchFlags"
args129(9).Value = 65536
args129(10).Name = "SearchItem.SearchString"
args129(10).Value = "Weimar Republiko"
args129(11).Name = "SearchItem.ReplaceString"
args129(11).Value = "Weimar-Republiko"
args129(12).Name = "SearchItem.Locale"
args129(12).Value = 255
args129(13).Name = "SearchItem.ChangedChars"
args129(13).Value = 2
args129(14).Name = "SearchItem.DeletedChars"
args129(14).Value = 2
args129(15).Name = "SearchItem.InsertedChars"
args129(15).Value = 2
args129(16).Name = "SearchItem.TransliterateFlags"
args129(16).Value = 1024
args129(17).Name = "SearchItem.Command"
args129(17).Value = 3
args129(18).Name = "Quiet"
args129(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args129())

rem ----------------------------------------------------------------------
dim args130(18) as new com.sun.star.beans.PropertyValue
args130(0).Name = "SearchItem.StyleFamily"
args130(0).Value = 2
args130(1).Name = "SearchItem.CellType"
args130(1).Value = 0
args130(2).Name = "SearchItem.RowDirection"
args130(2).Value = true
args130(3).Name = "SearchItem.AllTables"
args130(3).Value = false
args130(4).Name = "SearchItem.Backward"
args130(4).Value = false
args130(5).Name = "SearchItem.Pattern"
args130(5).Value = false
args130(6).Name = "SearchItem.Content"
args130(6).Value = false
args130(7).Name = "SearchItem.AsianOptions"
args130(7).Value = false
args130(8).Name = "SearchItem.AlgorithmType"
args130(8).Value = 0
args130(9).Name = "SearchItem.SearchFlags"
args130(9).Value = 65536
args130(10).Name = "SearchItem.SearchString"
args130(10).Value = "nedependenta"
args130(11).Name = "SearchItem.ReplaceString"
args130(11).Value = "nedependanta"
args130(12).Name = "SearchItem.Locale"
args130(12).Value = 255
args130(13).Name = "SearchItem.ChangedChars"
args130(13).Value = 2
args130(14).Name = "SearchItem.DeletedChars"
args130(14).Value = 2
args130(15).Name = "SearchItem.InsertedChars"
args130(15).Value = 2
args130(16).Name = "SearchItem.TransliterateFlags"
args130(16).Value = 1024
args130(17).Name = "SearchItem.Command"
args130(17).Value = 3
args130(18).Name = "Quiet"
args130(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args130())

rem ----------------------------------------------------------------------
dim args131(18) as new com.sun.star.beans.PropertyValue
args131(0).Name = "SearchItem.StyleFamily"
args131(0).Value = 2
args131(1).Name = "SearchItem.CellType"
args131(1).Value = 0
args131(2).Name = "SearchItem.RowDirection"
args131(2).Value = true
args131(3).Name = "SearchItem.AllTables"
args131(3).Value = false
args131(4).Name = "SearchItem.Backward"
args131(4).Value = false
args131(5).Name = "SearchItem.Pattern"
args131(5).Value = false
args131(6).Name = "SearchItem.Content"
args131(6).Value = false
args131(7).Name = "SearchItem.AsianOptions"
args131(7).Value = false
args131(8).Name = "SearchItem.AlgorithmType"
args131(8).Value = 0
args131(9).Name = "SearchItem.SearchFlags"
args131(9).Value = 65536
args131(10).Name = "SearchItem.SearchString"
args131(10).Value = "gradoze"
args131(11).Name = "SearchItem.ReplaceString"
args131(11).Value = "gradope"
args131(12).Name = "SearchItem.Locale"
args131(12).Value = 255
args131(13).Name = "SearchItem.ChangedChars"
args131(13).Value = 2
args131(14).Name = "SearchItem.DeletedChars"
args131(14).Value = 2
args131(15).Name = "SearchItem.InsertedChars"
args131(15).Value = 2
args131(16).Name = "SearchItem.TransliterateFlags"
args131(16).Value = 1024
args131(17).Name = "SearchItem.Command"
args131(17).Value = 3
args131(18).Name = "Quiet"
args131(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args131())

rem ----------------------------------------------------------------------
dim args132(18) as new com.sun.star.beans.PropertyValue
args132(0).Name = "SearchItem.StyleFamily"
args132(0).Value = 2
args132(1).Name = "SearchItem.CellType"
args132(1).Value = 0
args132(2).Name = "SearchItem.RowDirection"
args132(2).Value = true
args132(3).Name = "SearchItem.AllTables"
args132(3).Value = false
args132(4).Name = "SearchItem.Backward"
args132(4).Value = false
args132(5).Name = "SearchItem.Pattern"
args132(5).Value = false
args132(6).Name = "SearchItem.Content"
args132(6).Value = false
args132(7).Name = "SearchItem.AsianOptions"
args132(7).Value = false
args132(8).Name = "SearchItem.AlgorithmType"
args132(8).Value = 0
args132(9).Name = "SearchItem.SearchFlags"
args132(9).Value = 65536
args132(10).Name = "SearchItem.SearchString"
args132(10).Value = "lojanto-denseso"
args132(11).Name = "SearchItem.ReplaceString"
args132(11).Value = "denseso di habitantaro"
args132(12).Name = "SearchItem.Locale"
args132(12).Value = 255
args132(13).Name = "SearchItem.ChangedChars"
args132(13).Value = 2
args132(14).Name = "SearchItem.DeletedChars"
args132(14).Value = 2
args132(15).Name = "SearchItem.InsertedChars"
args132(15).Value = 2
args132(16).Name = "SearchItem.TransliterateFlags"
args132(16).Value = 1280
args132(17).Name = "SearchItem.Command"
args132(17).Value = 3
args132(18).Name = "Quiet"
args132(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args132())


rem ----------------------------------------------------------------------
dim args133(18) as new com.sun.star.beans.PropertyValue
args133(0).Name = "SearchItem.StyleFamily"
args133(0).Value = 2
args133(1).Name = "SearchItem.CellType"
args133(1).Value = 0
args133(2).Name = "SearchItem.RowDirection"
args133(2).Value = true
args133(3).Name = "SearchItem.AllTables"
args133(3).Value = false
args133(4).Name = "SearchItem.Backward"
args133(4).Value = false
args133(5).Name = "SearchItem.Pattern"
args133(5).Value = false
args133(6).Name = "SearchItem.Content"
args133(6).Value = false
args133(7).Name = "SearchItem.AsianOptions"
args133(7).Value = false
args133(8).Name = "SearchItem.AlgorithmType"
args133(8).Value = 0
args133(9).Name = "SearchItem.SearchFlags"
args133(9).Value = 65536
args133(10).Name = "SearchItem.SearchString"
args133(10).Value = "kondemn"
args133(11).Name = "SearchItem.ReplaceString"
args133(11).Value = "kondamn"
args133(12).Name = "SearchItem.Locale"
args133(12).Value = 255
args133(13).Name = "SearchItem.ChangedChars"
args133(13).Value = 2
args133(14).Name = "SearchItem.DeletedChars"
args133(14).Value = 2
args133(15).Name = "SearchItem.InsertedChars"
args133(15).Value = 2
args133(16).Name = "SearchItem.TransliterateFlags"
args133(16).Value = 1280
args133(17).Name = "SearchItem.Command"
args133(17).Value = 3
args133(18).Name = "Quiet"
args133(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args133())

rem ----------------------------------------------------------------------
dim args134(18) as new com.sun.star.beans.PropertyValue

dim OKTOB1, OKTOB2 as string
For I = 1 to 31
OKTOB1 = Ltrim(Str$(I))+" di oktob"
OKTOB2 = Ltrim(Str$(I))+"ma di oktob"

args134(0).Name = "SearchItem.StyleFamily"
args134(0).Value = 2
args134(1).Name = "SearchItem.CellType"
args134(1).Value = 0
args134(2).Name = "SearchItem.RowDirection"
args134(2).Value = true
args134(3).Name = "SearchItem.AllTables"
args134(3).Value = false
args134(4).Name = "SearchItem.Backward"
args134(4).Value = false
args134(5).Name = "SearchItem.Pattern"
args134(5).Value = false
args134(6).Name = "SearchItem.Content"
args134(6).Value = false
args134(7).Name = "SearchItem.AsianOptions"
args134(7).Value = false
args134(8).Name = "SearchItem.AlgorithmType"
args134(8).Value = 0
args134(9).Name = "SearchItem.SearchFlags"
args134(9).Value = 65536
args134(10).Name = "SearchItem.SearchString"
args134(10).Value = OKTOB1
args134(11).Name = "SearchItem.ReplaceString"
args134(11).Value = OKTOB2
args134(12).Name = "SearchItem.Locale"
args134(12).Value = 255
args134(13).Name = "SearchItem.ChangedChars"
args134(13).Value = 2
args134(14).Name = "SearchItem.DeletedChars"
args134(14).Value = 2
args134(15).Name = "SearchItem.InsertedChars"
args134(15).Value = 2
args134(16).Name = "SearchItem.TransliterateFlags"
args134(16).Value = 1280
args134(17).Name = "SearchItem.Command"
args134(17).Value = 3
args134(18).Name = "Quiet"
args134(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args134()

Next I

rem ----------------------------------------------------------------------
dim args135(18) as new com.sun.star.beans.PropertyValue
args135(0).Name = "SearchItem.StyleFamily"
args135(0).Value = 2
args135(1).Name = "SearchItem.CellType"
args135(1).Value = 0
args135(2).Name = "SearchItem.RowDirection"
args135(2).Value = true
args135(3).Name = "SearchItem.AllTables"
args135(3).Value = false
args135(4).Name = "SearchItem.Backward"
args135(4).Value = false
args135(5).Name = "SearchItem.Pattern"
args135(5).Value = false
args135(6).Name = "SearchItem.Content"
args135(6).Value = false
args135(7).Name = "SearchItem.AsianOptions"
args135(7).Value = false
args135(8).Name = "SearchItem.AlgorithmType"
args135(8).Value = 0
args135(9).Name = "SearchItem.SearchFlags"
args135(9).Value = 65536
args135(10).Name = "SearchItem.SearchString"
args135(10).Value = "civiten"
args135(11).Name = "SearchItem.ReplaceString"
args135(11).Value = "civitan"
args135(12).Name = "SearchItem.Locale"
args135(12).Value = 255
args135(13).Name = "SearchItem.ChangedChars"
args135(13).Value = 2
args135(14).Name = "SearchItem.DeletedChars"
args135(14).Value = 2
args135(15).Name = "SearchItem.InsertedChars"
args135(15).Value = 2
args135(16).Name = "SearchItem.TransliterateFlags"
args135(16).Value = 1280
args135(17).Name = "SearchItem.Command"
args135(17).Value = 3
args135(18).Name = "Quiet"
args135(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args135())

rem ----------------------------------------------------------------------
dim args136(18) as new com.sun.star.beans.PropertyValue
args136(0).Name = "SearchItem.StyleFamily"
args136(0).Value = 2
args136(1).Name = "SearchItem.CellType"
args136(1).Value = 0
args136(2).Name = "SearchItem.RowDirection"
args136(2).Value = true
args136(3).Name = "SearchItem.AllTables"
args136(3).Value = false
args136(4).Name = "SearchItem.Backward"
args136(4).Value = false
args136(5).Name = "SearchItem.Pattern"
args136(5).Value = false
args136(6).Name = "SearchItem.Content"
args136(6).Value = false
args136(7).Name = "SearchItem.AsianOptions"
args136(7).Value = false
args136(8).Name = "SearchItem.AlgorithmType"
args136(8).Value = 0
args136(9).Name = "SearchItem.SearchFlags"
args136(9).Value = 65536
args136(10).Name = "SearchItem.SearchString"
args136(10).Value = "piloto"
args136(11).Name = "SearchItem.ReplaceString"
args136(11).Value = "pilotisto"
args136(12).Name = "SearchItem.Locale"
args136(12).Value = 255
args136(13).Name = "SearchItem.ChangedChars"
args136(13).Value = 2
args136(14).Name = "SearchItem.DeletedChars"
args136(14).Value = 2
args136(15).Name = "SearchItem.InsertedChars"
args136(15).Value = 2
args136(16).Name = "SearchItem.TransliterateFlags"
args136(16).Value = 1280
args136(17).Name = "SearchItem.Command"
args136(17).Value = 3
args136(18).Name = "Quiet"
args136(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args136())

rem ----------------------------------------------------------------------
dim args137(18) as new com.sun.star.beans.PropertyValue
args137(0).Name = "SearchItem.StyleFamily"
args137(0).Value = 2
args137(1).Name = "SearchItem.CellType"
args137(1).Value = 0
args137(2).Name = "SearchItem.RowDirection"
args137(2).Value = true
args137(3).Name = "SearchItem.AllTables"
args137(3).Value = false
args137(4).Name = "SearchItem.Backward"
args137(4).Value = false
args137(5).Name = "SearchItem.Pattern"
args137(5).Value = false
args137(6).Name = "SearchItem.Content"
args137(6).Value = false
args137(7).Name = "SearchItem.AsianOptions"
args137(7).Value = false
args137(8).Name = "SearchItem.AlgorithmType"
args137(8).Value = 0
args137(9).Name = "SearchItem.SearchFlags"
args137(9).Value = 65536
args137(10).Name = "SearchItem.SearchString"
args137(10).Value = "renoncis"
args137(11).Name = "SearchItem.ReplaceString"
args137(11).Value = "renuncis"
args137(12).Name = "SearchItem.Locale"
args137(12).Value = 255
args137(13).Name = "SearchItem.ChangedChars"
args137(13).Value = 2
args137(14).Name = "SearchItem.DeletedChars"
args137(14).Value = 2
args137(15).Name = "SearchItem.InsertedChars"
args137(15).Value = 2
args137(16).Name = "SearchItem.TransliterateFlags"
args137(16).Value = 1280
args137(17).Name = "SearchItem.Command"
args137(17).Value = 3
args137(18).Name = "Quiet"
args137(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args137())

rem ----------------------------------------------------------------------
dim args138(18) as new com.sun.star.beans.PropertyValue
args138(0).Name = "SearchItem.StyleFamily"
args138(0).Value = 2
args138(1).Name = "SearchItem.CellType"
args138(1).Value = 0
args138(2).Name = "SearchItem.RowDirection"
args138(2).Value = true
args138(3).Name = "SearchItem.AllTables"
args138(3).Value = false
args138(4).Name = "SearchItem.Backward"
args138(4).Value = false
args138(5).Name = "SearchItem.Pattern"
args138(5).Value = false
args138(6).Name = "SearchItem.Content"
args138(6).Value = false
args138(7).Name = "SearchItem.AsianOptions"
args138(7).Value = false
args138(8).Name = "SearchItem.AlgorithmType"
args138(8).Value = 0
args138(9).Name = "SearchItem.SearchFlags"
args138(9).Value = 65536
args138(10).Name = "SearchItem.SearchString"
args138(10).Value = "presenco"
args138(11).Name = "SearchItem.ReplaceString"
args138(11).Value = "prezenteso"
args138(12).Name = "SearchItem.Locale"
args138(12).Value = 255
args138(13).Name = "SearchItem.ChangedChars"
args138(13).Value = 2
args138(14).Name = "SearchItem.DeletedChars"
args138(14).Value = 2
args138(15).Name = "SearchItem.InsertedChars"
args138(15).Value = 2
args138(16).Name = "SearchItem.TransliterateFlags"
args138(16).Value = 1280
args138(17).Name = "SearchItem.Command"
args138(17).Value = 3
args138(18).Name = "Quiet"
args138(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args138())

rem ----------------------------------------------------------------------
dim args139(18) as new com.sun.star.beans.PropertyValue
args139(0).Name = "SearchItem.StyleFamily"
args139(0).Value = 2
args139(1).Name = "SearchItem.CellType"
args139(1).Value = 0
args139(2).Name = "SearchItem.RowDirection"
args139(2).Value = true
args139(3).Name = "SearchItem.AllTables"
args139(3).Value = false
args139(4).Name = "SearchItem.Backward"
args139(4).Value = false
args139(5).Name = "SearchItem.Pattern"
args139(5).Value = false
args139(6).Name = "SearchItem.Content"
args139(6).Value = false
args139(7).Name = "SearchItem.AsianOptions"
args139(7).Value = false
args139(8).Name = "SearchItem.AlgorithmType"
args139(8).Value = 0
args139(9).Name = "SearchItem.SearchFlags"
args139(9).Value = 65536
args139(10).Name = "SearchItem.SearchString"
args139(10).Value = "hemisfero"
args139(11).Name = "SearchItem.ReplaceString"
args139(11).Value = "misfero"
args139(12).Name = "SearchItem.Locale"
args139(12).Value = 255
args139(13).Name = "SearchItem.ChangedChars"
args139(13).Value = 2
args139(14).Name = "SearchItem.DeletedChars"
args139(14).Value = 2
args139(15).Name = "SearchItem.InsertedChars"
args139(15).Value = 2
args139(16).Name = "SearchItem.TransliterateFlags"
args139(16).Value = 1280
args139(17).Name = "SearchItem.Command"
args139(17).Value = 3
args139(18).Name = "Quiet"
args139(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args139())

rem ----------------------------------------------------------------------
dim args140(18) as new com.sun.star.beans.PropertyValue
args140(0).Name = "SearchItem.StyleFamily"
args140(0).Value = 2
args140(1).Name = "SearchItem.CellType"
args140(1).Value = 0
args140(2).Name = "SearchItem.RowDirection"
args140(2).Value = true
args140(3).Name = "SearchItem.AllTables"
args140(3).Value = false
args140(4).Name = "SearchItem.Backward"
args140(4).Value = false
args140(5).Name = "SearchItem.Pattern"
args140(5).Value = false
args140(6).Name = "SearchItem.Content"
args140(6).Value = false
args140(7).Name = "SearchItem.AsianOptions"
args140(7).Value = false
args140(8).Name = "SearchItem.AlgorithmType"
args140(8).Value = 0
args140(9).Name = "SearchItem.SearchFlags"
args140(9).Value = 65536
args140(10).Name = "SearchItem.SearchString"
args140(10).Value = "ke okazis"
args140(11).Name = "SearchItem.ReplaceString"
args140(11).Value = "qua eventis"
args140(12).Name = "SearchItem.Locale"
args140(12).Value = 255
args140(13).Name = "SearchItem.ChangedChars"
args140(13).Value = 2
args140(14).Name = "SearchItem.DeletedChars"
args140(14).Value = 2
args140(15).Name = "SearchItem.InsertedChars"
args140(15).Value = 2
args140(16).Name = "SearchItem.TransliterateFlags"
args140(16).Value = 1280
args140(17).Name = "SearchItem.Command"
args140(17).Value = 3
args140(18).Name = "Quiet"
args140(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args140())

rem ----------------------------------------------------------------------
dim args141(18) as new com.sun.star.beans.PropertyValue
args141(0).Name = "SearchItem.StyleFamily"
args141(0).Value = 2
args141(1).Name = "SearchItem.CellType"
args141(1).Value = 0
args141(2).Name = "SearchItem.RowDirection"
args141(2).Value = true
args141(3).Name = "SearchItem.AllTables"
args141(3).Value = false
args141(4).Name = "SearchItem.Backward"
args141(4).Value = false
args141(5).Name = "SearchItem.Pattern"
args141(5).Value = false
args141(6).Name = "SearchItem.Content"
args141(6).Value = false
args141(7).Name = "SearchItem.AsianOptions"
args141(7).Value = false
args141(8).Name = "SearchItem.AlgorithmType"
args141(8).Value = 0
args141(9).Name = "SearchItem.SearchFlags"
args141(9).Value = 65536
args141(10).Name = "SearchItem.SearchString"
args141(10).Value = "importas"
args141(11).Name = "SearchItem.ReplaceString"
args141(11).Value = "importacas"
args141(12).Name = "SearchItem.Locale"
args141(12).Value = 255
args141(13).Name = "SearchItem.ChangedChars"
args141(13).Value = 2
args141(14).Name = "SearchItem.DeletedChars"
args141(14).Value = 2
args141(15).Name = "SearchItem.InsertedChars"
args141(15).Value = 2
args141(16).Name = "SearchItem.TransliterateFlags"
args141(16).Value = 1280
args141(17).Name = "SearchItem.Command"
args141(17).Value = 3
args141(18).Name = "Quiet"
args141(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args141())

rem ----------------------------------------------------------------------
dim args142(18) as new com.sun.star.beans.PropertyValue
args142(0).Name = "SearchItem.StyleFamily"
args142(0).Value = 2
args142(1).Name = "SearchItem.CellType"
args142(1).Value = 0
args142(2).Name = "SearchItem.RowDirection"
args142(2).Value = true
args142(3).Name = "SearchItem.AllTables"
args142(3).Value = false
args142(4).Name = "SearchItem.Backward"
args142(4).Value = false
args142(5).Name = "SearchItem.Pattern"
args142(5).Value = false
args142(6).Name = "SearchItem.Content"
args142(6).Value = false
args142(7).Name = "SearchItem.AsianOptions"
args142(7).Value = false
args142(8).Name = "SearchItem.AlgorithmType"
args142(8).Value = 0
args142(9).Name = "SearchItem.SearchFlags"
args142(9).Value = 65536
args142(10).Name = "SearchItem.SearchString"
args142(10).Value = "por unesma foyo"
args142(11).Name = "SearchItem.ReplaceString"
args142(11).Value = "unesma-foye"
args142(12).Name = "SearchItem.Locale"
args142(12).Value = 255
args142(13).Name = "SearchItem.ChangedChars"
args142(13).Value = 2
args142(14).Name = "SearchItem.DeletedChars"
args142(14).Value = 2
args142(15).Name = "SearchItem.InsertedChars"
args142(15).Value = 2
args142(16).Name = "SearchItem.TransliterateFlags"
args142(16).Value = 1280
args142(17).Name = "SearchItem.Command"
args142(17).Value = 3
args142(18).Name = "Quiet"
args142(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args142())

rem ----------------------------------------------------------------------
dim args143(18) as new com.sun.star.beans.PropertyValue
args143(0).Name = "SearchItem.StyleFamily"
args143(0).Value = 2
args143(1).Name = "SearchItem.CellType"
args143(1).Value = 0
args143(2).Name = "SearchItem.RowDirection"
args143(2).Value = true
args143(3).Name = "SearchItem.AllTables"
args143(3).Value = false
args143(4).Name = "SearchItem.Backward"
args143(4).Value = false
args143(5).Name = "SearchItem.Pattern"
args143(5).Value = false
args143(6).Name = "SearchItem.Content"
args143(6).Value = false
args143(7).Name = "SearchItem.AsianOptions"
args143(7).Value = false
args143(8).Name = "SearchItem.AlgorithmType"
args143(8).Value = 0
args143(9).Name = "SearchItem.SearchFlags"
args143(9).Value = 65536
args143(10).Name = "SearchItem.SearchString"
args143(10).Value = "rinominesas"
args143(11).Name = "SearchItem.ReplaceString"
args143(11).Value = "rinomizesas"
args143(12).Name = "SearchItem.Locale"
args143(12).Value = 255
args143(13).Name = "SearchItem.ChangedChars"
args143(13).Value = 2
args143(14).Name = "SearchItem.DeletedChars"
args143(14).Value = 2
args143(15).Name = "SearchItem.InsertedChars"
args143(15).Value = 2
args143(16).Name = "SearchItem.TransliterateFlags"
args143(16).Value = 1280
args143(17).Name = "SearchItem.Command"
args143(17).Value = 3
args143(18).Name = "Quiet"
args143(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args143())

rem ----------------------------------------------------------------------
dim args144(18) as new com.sun.star.beans.PropertyValue
args144(0).Name = "SearchItem.StyleFamily"
args144(0).Value = 2
args144(1).Name = "SearchItem.CellType"
args144(1).Value = 0
args144(2).Name = "SearchItem.RowDirection"
args144(2).Value = true
args144(3).Name = "SearchItem.AllTables"
args144(3).Value = false
args144(4).Name = "SearchItem.Backward"
args144(4).Value = false
args144(5).Name = "SearchItem.Pattern"
args144(5).Value = false
args144(6).Name = "SearchItem.Content"
args144(6).Value = false
args144(7).Name = "SearchItem.AsianOptions"
args144(7).Value = false
args144(8).Name = "SearchItem.AlgorithmType"
args144(8).Value = 0
args144(9).Name = "SearchItem.SearchFlags"
args144(9).Value = 65536
args144(10).Name = "SearchItem.SearchString"
args144(10).Value = "konferenc"
args144(11).Name = "SearchItem.ReplaceString"
args144(11).Value = "konfer"
args144(12).Name = "SearchItem.Locale"
args144(12).Value = 255
args144(13).Name = "SearchItem.ChangedChars"
args144(13).Value = 2
args144(14).Name = "SearchItem.DeletedChars"
args144(14).Value = 2
args144(15).Name = "SearchItem.InsertedChars"
args144(15).Value = 2
args144(16).Name = "SearchItem.TransliterateFlags"
args144(16).Value = 1280
args144(17).Name = "SearchItem.Command"
args144(17).Value = 3
args144(18).Name = "Quiet"
args144(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args144())

rem ----------------------------------------------------------------------
dim args145(18) as new com.sun.star.beans.PropertyValue
args145(0).Name = "SearchItem.StyleFamily"
args145(0).Value = 2
args145(1).Name = "SearchItem.CellType"
args145(1).Value = 0
args145(2).Name = "SearchItem.RowDirection"
args145(2).Value = true
args145(3).Name = "SearchItem.AllTables"
args145(3).Value = false
args145(4).Name = "SearchItem.Backward"
args145(4).Value = false
args145(5).Name = "SearchItem.Pattern"
args145(5).Value = false
args145(6).Name = "SearchItem.Content"
args145(6).Value = false
args145(7).Name = "SearchItem.AsianOptions"
args145(7).Value = false
args145(8).Name = "SearchItem.AlgorithmType"
args145(8).Value = 0
args145(9).Name = "SearchItem.SearchFlags"
args145(9).Value = 65536
args145(10).Name = "SearchItem.SearchString"
args145(10).Value = "konstruita"
args145(11).Name = "SearchItem.ReplaceString"
args145(11).Value = "konstruktita"
args145(12).Name = "SearchItem.Locale"
args145(12).Value = 255
args145(13).Name = "SearchItem.ChangedChars"
args145(13).Value = 2
args145(14).Name = "SearchItem.DeletedChars"
args145(14).Value = 2
args145(15).Name = "SearchItem.InsertedChars"
args145(15).Value = 2
args145(16).Name = "SearchItem.TransliterateFlags"
args145(16).Value = 1280
args145(17).Name = "SearchItem.Command"
args145(17).Value = 3
args145(18).Name = "Quiet"
args145(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args145())

rem ----------------------------------------------------------------------
dim args146(18) as new com.sun.star.beans.PropertyValue
args146(0).Name = "SearchItem.StyleFamily"
args146(0).Value = 2
args146(1).Name = "SearchItem.CellType"
args146(1).Value = 0
args146(2).Name = "SearchItem.RowDirection"
args146(2).Value = true
args146(3).Name = "SearchItem.AllTables"
args146(3).Value = false
args146(4).Name = "SearchItem.Backward"
args146(4).Value = false
args146(5).Name = "SearchItem.Pattern"
args146(5).Value = false
args146(6).Name = "SearchItem.Content"
args146(6).Value = false
args146(7).Name = "SearchItem.AsianOptions"
args146(7).Value = false
args146(8).Name = "SearchItem.AlgorithmType"
args146(8).Value = 0
args146(9).Name = "SearchItem.SearchFlags"
args146(9).Value = 65536
args146(10).Name = "SearchItem.SearchString"
args146(10).Value = "autoritatema"
args146(11).Name = "SearchItem.ReplaceString"
args146(11).Value = "autoritatoza"
args146(12).Name = "SearchItem.Locale"
args146(12).Value = 255
args146(13).Name = "SearchItem.ChangedChars"
args146(13).Value = 2
args146(14).Name = "SearchItem.DeletedChars"
args146(14).Value = 2
args146(15).Name = "SearchItem.InsertedChars"
args146(15).Value = 2
args146(16).Name = "SearchItem.TransliterateFlags"
args146(16).Value = 1280
args146(17).Name = "SearchItem.Command"
args146(17).Value = 3
args146(18).Name = "Quiet"
args146(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args146())

rem ----------------------------------------------------------------------
dim args147(18) as new com.sun.star.beans.PropertyValue
args147(0).Name = "SearchItem.StyleFamily"
args147(0).Value = 2
args147(1).Name = "SearchItem.CellType"
args147(1).Value = 0
args147(2).Name = "SearchItem.RowDirection"
args147(2).Value = true
args147(3).Name = "SearchItem.AllTables"
args147(3).Value = false
args147(4).Name = "SearchItem.Backward"
args147(4).Value = false
args147(5).Name = "SearchItem.Pattern"
args147(5).Value = false
args147(6).Name = "SearchItem.Content"
args147(6).Value = false
args147(7).Name = "SearchItem.AsianOptions"
args147(7).Value = false
args147(8).Name = "SearchItem.AlgorithmType"
args147(8).Value = 0
args147(9).Name = "SearchItem.SearchFlags"
args147(9).Value = 65536
args147(10).Name = "SearchItem.SearchString"
args147(10).Value = "partoprenis a "
args147(11).Name = "SearchItem.ReplaceString"
args147(11).Value = "apartenis a "
args147(12).Name = "SearchItem.Locale"
args147(12).Value = 255
args147(13).Name = "SearchItem.ChangedChars"
args147(13).Value = 2
args147(14).Name = "SearchItem.DeletedChars"
args147(14).Value = 2
args147(15).Name = "SearchItem.InsertedChars"
args147(15).Value = 2
args147(16).Name = "SearchItem.TransliterateFlags"
args147(16).Value = 1280
args147(17).Name = "SearchItem.Command"
args147(17).Value = 3
args147(18).Name = "Quiet"
args147(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args147())

rem ----------------------------------------------------------------------
dim args148(18) as new com.sun.star.beans.PropertyValue
args148(0).Name = "SearchItem.StyleFamily"
args148(0).Value = 2
args148(1).Name = "SearchItem.CellType"
args148(1).Value = 0
args148(2).Name = "SearchItem.RowDirection"
args148(2).Value = true
args148(3).Name = "SearchItem.AllTables"
args148(3).Value = false
args148(4).Name = "SearchItem.Backward"
args148(4).Value = false
args148(5).Name = "SearchItem.Pattern"
args148(5).Value = false
args148(6).Name = "SearchItem.Content"
args148(6).Value = false
args148(7).Name = "SearchItem.AsianOptions"
args148(7).Value = false
args148(8).Name = "SearchItem.AlgorithmType"
args148(8).Value = 0
args148(9).Name = "SearchItem.SearchFlags"
args148(9).Value = 65536
args148(10).Name = "SearchItem.SearchString"
args148(10).Value = "defendar"
args148(11).Name = "SearchItem.ReplaceString"
args148(11).Value = "defensar"
args148(12).Name = "SearchItem.Locale"
args148(12).Value = 255
args148(13).Name = "SearchItem.ChangedChars"
args148(13).Value = 2
args148(14).Name = "SearchItem.DeletedChars"
args148(14).Value = 2
args148(15).Name = "SearchItem.InsertedChars"
args148(15).Value = 2
args148(16).Name = "SearchItem.TransliterateFlags"
args148(16).Value = 1280
args148(17).Name = "SearchItem.Command"
args148(17).Value = 3
args148(18).Name = "Quiet"
args148(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args148())

rem ----------------------------------------------------------------------
dim args149(18) as new com.sun.star.beans.PropertyValue
args149(0).Name = "SearchItem.StyleFamily"
args149(0).Value = 2
args149(1).Name = "SearchItem.CellType"
args149(1).Value = 0
args149(2).Name = "SearchItem.RowDirection"
args149(2).Value = true
args149(3).Name = "SearchItem.AllTables"
args149(3).Value = false
args149(4).Name = "SearchItem.Backward"
args149(4).Value = false
args149(5).Name = "SearchItem.Pattern"
args149(5).Value = false
args149(6).Name = "SearchItem.Content"
args149(6).Value = false
args149(7).Name = "SearchItem.AsianOptions"
args149(7).Value = false
args149(8).Name = "SearchItem.AlgorithmType"
args149(8).Value = 0
args149(9).Name = "SearchItem.SearchFlags"
args149(9).Value = 65536
args149(10).Name = "SearchItem.SearchString"
args149(10).Value = "defendas"
args149(11).Name = "SearchItem.ReplaceString"
args149(11).Value = "defensas"
args149(12).Name = "SearchItem.Locale"
args149(12).Value = 255
args149(13).Name = "SearchItem.ChangedChars"
args149(13).Value = 2
args149(14).Name = "SearchItem.DeletedChars"
args149(14).Value = 2
args149(15).Name = "SearchItem.InsertedChars"
args149(15).Value = 2
args149(16).Name = "SearchItem.TransliterateFlags"
args149(16).Value = 1280
args149(17).Name = "SearchItem.Command"
args149(17).Value = 3
args149(18).Name = "Quiet"
args149(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args149())

rem ----------------------------------------------------------------------
dim args150(18) as new com.sun.star.beans.PropertyValue
args150(0).Name = "SearchItem.StyleFamily"
args150(0).Value = 2
args150(1).Name = "SearchItem.CellType"
args150(1).Value = 0
args150(2).Name = "SearchItem.RowDirection"
args150(2).Value = true
args150(3).Name = "SearchItem.AllTables"
args150(3).Value = false
args150(4).Name = "SearchItem.Backward"
args150(4).Value = false
args150(5).Name = "SearchItem.Pattern"
args150(5).Value = false
args150(6).Name = "SearchItem.Content"
args150(6).Value = false
args150(7).Name = "SearchItem.AsianOptions"
args150(7).Value = false
args150(8).Name = "SearchItem.AlgorithmType"
args150(8).Value = 0
args150(9).Name = "SearchItem.SearchFlags"
args150(9).Value = 65536
args150(10).Name = "SearchItem.SearchString"
args150(10).Value = "speciale"
args150(11).Name = "SearchItem.ReplaceString"
args150(11).Value = "specale"
args150(12).Name = "SearchItem.Locale"
args150(12).Value = 255
args150(13).Name = "SearchItem.ChangedChars"
args150(13).Value = 2
args150(14).Name = "SearchItem.DeletedChars"
args150(14).Value = 2
args150(15).Name = "SearchItem.InsertedChars"
args150(15).Value = 2
args150(16).Name = "SearchItem.TransliterateFlags"
args150(16).Value = 1280
args150(17).Name = "SearchItem.Command"
args150(17).Value = 3
args150(18).Name = "Quiet"
args150(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args150())

rem ----------------------------------------------------------------------
dim args151(18) as new com.sun.star.beans.PropertyValue
args151(0).Name = "SearchItem.StyleFamily"
args151(0).Value = 2
args151(1).Name = "SearchItem.CellType"
args151(1).Value = 0
args151(2).Name = "SearchItem.RowDirection"
args151(2).Value = true
args151(3).Name = "SearchItem.AllTables"
args151(3).Value = false
args151(4).Name = "SearchItem.Backward"
args151(4).Value = false
args151(5).Name = "SearchItem.Pattern"
args151(5).Value = false
args151(6).Name = "SearchItem.Content"
args151(6).Value = false
args151(7).Name = "SearchItem.AsianOptions"
args151(7).Value = false
args151(8).Name = "SearchItem.AlgorithmType"
args151(8).Value = 0
args151(9).Name = "SearchItem.SearchFlags"
args151(9).Value = 65536
args151(10).Name = "SearchItem.SearchString"
args151(10).Value = "okuras"
args151(11).Name = "SearchItem.ReplaceString"
args151(11).Value = "eventas"
args151(12).Name = "SearchItem.Locale"
args151(12).Value = 255
args151(13).Name = "SearchItem.ChangedChars"
args151(13).Value = 2
args151(14).Name = "SearchItem.DeletedChars"
args151(14).Value = 2
args151(15).Name = "SearchItem.InsertedChars"
args151(15).Value = 2
args151(16).Name = "SearchItem.TransliterateFlags"
args151(16).Value = 1280
args151(17).Name = "SearchItem.Command"
args151(17).Value = 3
args151(18).Name = "Quiet"
args151(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args151())

rem ----------------------------------------------------------------------
dim args152(18) as new com.sun.star.beans.PropertyValue
args152(0).Name = "SearchItem.StyleFamily"
args152(0).Value = 2
args152(1).Name = "SearchItem.CellType"
args152(1).Value = 0
args152(2).Name = "SearchItem.RowDirection"
args152(2).Value = true
args152(3).Name = "SearchItem.AllTables"
args152(3).Value = false
args152(4).Name = "SearchItem.Backward"
args152(4).Value = false
args152(5).Name = "SearchItem.Pattern"
args152(5).Value = false
args152(6).Name = "SearchItem.Content"
args152(6).Value = false
args152(7).Name = "SearchItem.AsianOptions"
args152(7).Value = false
args152(8).Name = "SearchItem.AlgorithmType"
args152(8).Value = 0
args152(9).Name = "SearchItem.SearchFlags"
args152(9).Value = 65536
args152(10).Name = "SearchItem.SearchString"
args152(10).Value = "esis_dvidita"
args152(11).Name = "SearchItem.ReplaceString"
args152(11).Value = "dividesis"
args152(12).Name = "SearchItem.Locale"
args152(12).Value = 255
args152(13).Name = "SearchItem.ChangedChars"
args152(13).Value = 2
args152(14).Name = "SearchItem.DeletedChars"
args152(14).Value = 2
args152(15).Name = "SearchItem.InsertedChars"
args152(15).Value = 2
args152(16).Name = "SearchItem.TransliterateFlags"
args152(16).Value = 1280
args152(17).Name = "SearchItem.Command"
args152(17).Value = 3
args152(18).Name = "Quiet"
args152(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args152())

rem ----------------------------------------------------------------------
dim args153(18) as new com.sun.star.beans.PropertyValue
args153(0).Name = "SearchItem.StyleFamily"
args153(0).Value = 2
args153(1).Name = "SearchItem.CellType"
args153(1).Value = 0
args153(2).Name = "SearchItem.RowDirection"
args153(2).Value = true
args153(3).Name = "SearchItem.AllTables"
args153(3).Value = false
args153(4).Name = "SearchItem.Backward"
args153(4).Value = false
args153(5).Name = "SearchItem.Pattern"
args153(5).Value = false
args153(6).Name = "SearchItem.Content"
args153(6).Value = false
args153(7).Name = "SearchItem.AsianOptions"
args153(7).Value = false
args153(8).Name = "SearchItem.AlgorithmType"
args153(8).Value = 0
args153(9).Name = "SearchItem.SearchFlags"
args153(9).Value = 65536
args153(10).Name = "SearchItem.SearchString"
args153(10).Value = "esis facata"
args153(11).Name = "SearchItem.ReplaceString"
args153(11).Value = "facesis"
args153(12).Name = "SearchItem.Locale"
args153(12).Value = 255
args153(13).Name = "SearchItem.ChangedChars"
args153(13).Value = 2
args153(14).Name = "SearchItem.DeletedChars"
args153(14).Value = 2
args153(15).Name = "SearchItem.InsertedChars"
args153(15).Value = 2
args153(16).Name = "SearchItem.TransliterateFlags"
args153(16).Value = 1280
args153(17).Name = "SearchItem.Command"
args153(17).Value = 3
args153(18).Name = "Quiet"
args153(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args153())

rem ----------------------------------------------------------------------
dim args154(18) as new com.sun.star.beans.PropertyValue
args154(0).Name = "SearchItem.StyleFamily"
args154(0).Value = 2
args154(1).Name = "SearchItem.CellType"
args154(1).Value = 0
args154(2).Name = "SearchItem.RowDirection"
args154(2).Value = true
args154(3).Name = "SearchItem.AllTables"
args154(3).Value = false
args154(4).Name = "SearchItem.Backward"
args154(4).Value = false
args154(5).Name = "SearchItem.Pattern"
args154(5).Value = false
args154(6).Name = "SearchItem.Content"
args154(6).Value = false
args154(7).Name = "SearchItem.AsianOptions"
args154(7).Value = false
args154(8).Name = "SearchItem.AlgorithmType"
args154(8).Value = 0
args154(9).Name = "SearchItem.SearchFlags"
args154(9).Value = 65536
args154(10).Name = "SearchItem.SearchString"
args154(10).Value = "direkt"
args154(11).Name = "SearchItem.ReplaceString"
args154(11).Value = "diret"
args154(12).Name = "SearchItem.Locale"
args154(12).Value = 255
args154(13).Name = "SearchItem.ChangedChars"
args154(13).Value = 2
args154(14).Name = "SearchItem.DeletedChars"
args154(14).Value = 2
args154(15).Name = "SearchItem.InsertedChars"
args154(15).Value = 2
args154(16).Name = "SearchItem.TransliterateFlags"
args154(16).Value = 1280
args154(17).Name = "SearchItem.Command"
args154(17).Value = 3
args154(18).Name = "Quiet"
args154(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args154())

rem ----------------------------------------------------------------------
dim args155(18) as new com.sun.star.beans.PropertyValue
args155(0).Name = "SearchItem.StyleFamily"
args155(0).Value = 2
args155(1).Name = "SearchItem.CellType"
args155(1).Value = 0
args155(2).Name = "SearchItem.RowDirection"
args155(2).Value = true
args155(3).Name = "SearchItem.AllTables"
args155(3).Value = false
args155(4).Name = "SearchItem.Backward"
args155(4).Value = false
args155(5).Name = "SearchItem.Pattern"
args155(5).Value = false
args155(6).Name = "SearchItem.Content"
args155(6).Value = false
args155(7).Name = "SearchItem.AsianOptions"
args155(7).Value = false
args155(8).Name = "SearchItem.AlgorithmType"
args155(8).Value = 0
args155(9).Name = "SearchItem.SearchFlags"
args155(9).Value = 65536
args155(10).Name = "SearchItem.SearchString"
args155(10).Value = "okuris"
args155(11).Name = "SearchItem.ReplaceString"
args155(11).Value = "eventis"
args155(12).Name = "SearchItem.Locale"
args155(12).Value = 255
args155(13).Name = "SearchItem.ChangedChars"
args155(13).Value = 2
args155(14).Name = "SearchItem.DeletedChars"
args155(14).Value = 2
args155(15).Name = "SearchItem.InsertedChars"
args155(15).Value = 2
args155(16).Name = "SearchItem.TransliterateFlags"
args155(16).Value = 1280
args155(17).Name = "SearchItem.Command"
args155(17).Value = 3
args155(18).Name = "Quiet"
args155(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args155())

rem ----------------------------------------------------------------------
dim args156(18) as new com.sun.star.beans.PropertyValue
args156(0).Name = "SearchItem.StyleFamily"
args156(0).Value = 2
args156(1).Name = "SearchItem.CellType"
args156(1).Value = 0
args156(2).Name = "SearchItem.RowDirection"
args156(2).Value = true
args156(3).Name = "SearchItem.AllTables"
args156(3).Value = false
args156(4).Name = "SearchItem.Backward"
args156(4).Value = false
args156(5).Name = "SearchItem.Pattern"
args156(5).Value = false
args156(6).Name = "SearchItem.Content"
args156(6).Value = false
args156(7).Name = "SearchItem.AsianOptions"
args156(7).Value = false
args156(8).Name = "SearchItem.AlgorithmType"
args156(8).Value = 0
args156(9).Name = "SearchItem.SearchFlags"
args156(9).Value = 65536
args156(10).Name = "SearchItem.SearchString"
args156(10).Value = "imediate"
args156(11).Name = "SearchItem.ReplaceString"
args156(11).Value = "nemediate"
args156(12).Name = "SearchItem.Locale"
args156(12).Value = 255
args156(13).Name = "SearchItem.ChangedChars"
args156(13).Value = 2
args156(14).Name = "SearchItem.DeletedChars"
args156(14).Value = 2
args156(15).Name = "SearchItem.InsertedChars"
args156(15).Value = 2
args156(16).Name = "SearchItem.TransliterateFlags"
args156(16).Value = 1280
args156(17).Name = "SearchItem.Command"
args156(17).Value = 3
args156(18).Name = "Quiet"
args156(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args156())

rem ----------------------------------------------------------------------
dim args157(18) as new com.sun.star.beans.PropertyValue
args157(0).Name = "SearchItem.StyleFamily"
args157(0).Value = 2
args157(1).Name = "SearchItem.CellType"
args157(1).Value = 0
args157(2).Name = "SearchItem.RowDirection"
args157(2).Value = true
args157(3).Name = "SearchItem.AllTables"
args157(3).Value = false
args157(4).Name = "SearchItem.Backward"
args157(4).Value = false
args157(5).Name = "SearchItem.Pattern"
args157(5).Value = false
args157(6).Name = "SearchItem.Content"
args157(6).Value = false
args157(7).Name = "SearchItem.AsianOptions"
args157(7).Value = false
args157(8).Name = "SearchItem.AlgorithmType"
args157(8).Value = 0
args157(9).Name = "SearchItem.SearchFlags"
args157(9).Value = 65536
args157(10).Name = "SearchItem.SearchString"
args157(10).Value = "Britanii"
args157(11).Name = "SearchItem.ReplaceString"
args157(11).Value = "Britaniani"
args157(12).Name = "SearchItem.Locale"
args157(12).Value = 255
args157(13).Name = "SearchItem.ChangedChars"
args157(13).Value = 2
args157(14).Name = "SearchItem.DeletedChars"
args157(14).Value = 2
args157(15).Name = "SearchItem.InsertedChars"
args157(15).Value = 2
args157(16).Name = "SearchItem.TransliterateFlags"
args157(16).Value = 1280
args157(17).Name = "SearchItem.Command"
args157(17).Value = 3
args157(18).Name = "Quiet"
args157(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args157())

rem ----------------------------------------------------------------------
dim args158(18) as new com.sun.star.beans.PropertyValue
args158(0).Name = "SearchItem.StyleFamily"
args158(0).Value = 2
args158(1).Name = "SearchItem.CellType"
args158(1).Value = 0
args158(2).Name = "SearchItem.RowDirection"
args158(2).Value = true
args158(3).Name = "SearchItem.AllTables"
args158(3).Value = false
args158(4).Name = "SearchItem.Backward"
args158(4).Value = false
args158(5).Name = "SearchItem.Pattern"
args158(5).Value = false
args158(6).Name = "SearchItem.Content"
args158(6).Value = false
args158(7).Name = "SearchItem.AsianOptions"
args158(7).Value = false
args158(8).Name = "SearchItem.AlgorithmType"
args158(8).Value = 0
args158(9).Name = "SearchItem.SearchFlags"
args158(9).Value = 65536
args158(10).Name = "SearchItem.SearchString"
args158(10).Value = "esis konquestita da "
args158(11).Name = "SearchItem.ReplaceString"
args158(11).Value = "konquestesis dal "
args158(12).Name = "SearchItem.Locale"
args158(12).Value = 255
args158(13).Name = "SearchItem.ChangedChars"
args158(13).Value = 2
args158(14).Name = "SearchItem.DeletedChars"
args158(14).Value = 2
args158(15).Name = "SearchItem.InsertedChars"
args158(15).Value = 2
args158(16).Name = "SearchItem.TransliterateFlags"
args158(16).Value = 1280
args158(17).Name = "SearchItem.Command"
args158(17).Value = 3
args158(18).Name = "Quiet"
args158(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args158())

rem ----------------------------------------------------------------------
dim args159(18) as new com.sun.star.beans.PropertyValue
args159(0).Name = "SearchItem.StyleFamily"
args159(0).Value = 2
args159(1).Name = "SearchItem.CellType"
args159(1).Value = 0
args159(2).Name = "SearchItem.RowDirection"
args159(2).Value = true
args159(3).Name = "SearchItem.AllTables"
args159(3).Value = false
args159(4).Name = "SearchItem.Backward"
args159(4).Value = false
args159(5).Name = "SearchItem.Pattern"
args159(5).Value = false
args159(6).Name = "SearchItem.Content"
args159(6).Value = false
args159(7).Name = "SearchItem.AsianOptions"
args159(7).Value = false
args159(8).Name = "SearchItem.AlgorithmType"
args159(8).Value = 0
args159(9).Name = "SearchItem.SearchFlags"
args159(9).Value = 65536
args159(10).Name = "SearchItem.SearchString"
args159(10).Value = "mezala denseso "
args159(11).Name = "SearchItem.ReplaceString"
args159(11).Value = "mezvalora denseso "
args159(12).Name = "SearchItem.Locale"
args159(12).Value = 255
args159(13).Name = "SearchItem.ChangedChars"
args159(13).Value = 2
args159(14).Name = "SearchItem.DeletedChars"
args159(14).Value = 2
args159(15).Name = "SearchItem.InsertedChars"
args159(15).Value = 2
args159(16).Name = "SearchItem.TransliterateFlags"
args159(16).Value = 1280
args159(17).Name = "SearchItem.Command"
args159(17).Value = 3
args159(18).Name = "Quiet"
args159(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args159())

rem ------------------------------------------------------------------
dim args160(18) as new com.sun.star.beans.PropertyValue
args160(0).Name = "SearchItem.StyleFamily"
args160(0).Value = 2
args160(1).Name = "SearchItem.CellType"
args160(1).Value = 0
args160(2).Name = "SearchItem.RowDirection"
args160(2).Value = true
args160(3).Name = "SearchItem.AllTables"
args160(3).Value = false
args160(4).Name = "SearchItem.Backward"
args160(4).Value = false
args160(5).Name = "SearchItem.Pattern"
args160(5).Value = false
args160(6).Name = "SearchItem.Content"
args160(6).Value = false
args160(7).Name = "SearchItem.AsianOptions"
args160(7).Value = false
args160(8).Name = "SearchItem.AlgorithmType"
args160(8).Value = 0
args160(9).Name = "SearchItem.SearchFlags"
args160(9).Value = 65536
args160(10).Name = "SearchItem.SearchString"
args160(10).Value = "&sup2;"
args160(11).Name = "SearchItem.ReplaceString"
args160(11).Value = "²"
args160(12).Name = "SearchItem.Locale"
args160(12).Value = 255
args160(13).Name = "SearchItem.ChangedChars"
args160(13).Value = 2
args160(14).Name = "SearchItem.DeletedChars"
args160(14).Value = 2
args160(15).Name = "SearchItem.InsertedChars"
args160(15).Value = 2
args160(16).Name = "SearchItem.TransliterateFlags"
args160(16).Value = 1280
args160(17).Name = "SearchItem.Command"
args160(17).Value = 3
args160(18).Name = "Quiet"
args160(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args160())

rem ----------------------------------------------------------------------
dim args161(18) as new com.sun.star.beans.PropertyValue
args161(0).Name = "SearchItem.StyleFamily"
args161(0).Value = 2
args161(1).Name = "SearchItem.CellType"
args161(1).Value = 0
args161(2).Name = "SearchItem.RowDirection"
args161(2).Value = true
args161(3).Name = "SearchItem.AllTables"
args161(3).Value = false
args161(4).Name = "SearchItem.Backward"
args161(4).Value = false
args161(5).Name = "SearchItem.Pattern"
args161(5).Value = false
args161(6).Name = "SearchItem.Content"
args161(6).Value = false
args161(7).Name = "SearchItem.AsianOptions"
args161(7).Value = false
args161(8).Name = "SearchItem.AlgorithmType"
args161(8).Value = 0
args161(9).Name = "SearchItem.SearchFlags"
args161(9).Value = 65536
args161(10).Name = "SearchItem.SearchString"
args161(10).Value = " prencipua"
args161(11).Name = "SearchItem.ReplaceString"
args161(11).Value = " precipua"
args161(12).Name = "SearchItem.Locale"
args161(12).Value = 255
args161(13).Name = "SearchItem.ChangedChars"
args161(13).Value = 2
args161(14).Name = "SearchItem.DeletedChars"
args161(14).Value = 2
args161(15).Name = "SearchItem.InsertedChars"
args161(15).Value = 2
args161(16).Name = "SearchItem.TransliterateFlags"
args161(16).Value = 1280
args161(17).Name = "SearchItem.Command"
args161(17).Value = 3
args161(18).Name = "Quiet"
args161(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args161())

rem ----------------------------------------------------------------------
dim args162(18) as new com.sun.star.beans.PropertyValue

dim FEBRUARO1, FEBRUARO2 as string
For I = 1 to 29
FEBRUARO1 = Ltrim(Str$(I))+" di februaro"
FEBRUARO2 = Ltrim(Str$(I))+"ma di februaro"

args162(0).Name = "SearchItem.StyleFamily"
args162(0).Value = 2
args162(1).Name = "SearchItem.CellType"
args162(1).Value = 0
args162(2).Name = "SearchItem.RowDirection"
args162(2).Value = true
args162(3).Name = "SearchItem.AllTables"
args162(3).Value = false
args162(4).Name = "SearchItem.Backward"
args162(4).Value = false
args162(5).Name = "SearchItem.Pattern"
args162(5).Value = false
args162(6).Name = "SearchItem.Content"
args162(6).Value = false
args162(7).Name = "SearchItem.AsianOptions"
args162(7).Value = false
args162(8).Name = "SearchItem.AlgorithmType"
args162(8).Value = 0
args162(9).Name = "SearchItem.SearchFlags"
args162(9).Value = 65536
args162(10).Name = "SearchItem.SearchString"
args162(10).Value = FEBRUARO1
args162(11).Name = "SearchItem.ReplaceString"
args162(11).Value = FEBRUARO2
args162(12).Name = "SearchItem.Locale"
args162(12).Value = 255
args162(13).Name = "SearchItem.ChangedChars"
args162(13).Value = 2
args162(14).Name = "SearchItem.DeletedChars"
args162(14).Value = 2
args162(15).Name = "SearchItem.InsertedChars"
args162(15).Value = 2
args162(16).Name = "SearchItem.TransliterateFlags"
args162(16).Value = 1280
args162(17).Name = "SearchItem.Command"
args162(17).Value = 3
args162(18).Name = "Quiet"
args162(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args162())

Next I

rem ----------------------------------------------------------------------
dim args163(18) as new com.sun.star.beans.PropertyValue
args163(0).Name = "SearchItem.StyleFamily"
args163(0).Value = 2
args163(1).Name = "SearchItem.CellType"
args163(1).Value = 0
args163(2).Name = "SearchItem.RowDirection"
args163(2).Value = true
args163(3).Name = "SearchItem.AllTables"
args163(3).Value = false
args163(4).Name = "SearchItem.Backward"
args163(4).Value = false
args163(5).Name = "SearchItem.Pattern"
args163(5).Value = false
args163(6).Name = "SearchItem.Content"
args163(6).Value = false
args163(7).Name = "SearchItem.AsianOptions"
args163(7).Value = false
args163(8).Name = "SearchItem.AlgorithmType"
args163(8).Value = 0
args163(9).Name = "SearchItem.SearchFlags"
args163(9).Value = 65536
args163(10).Name = "SearchItem.SearchString"
args163(10).Value = "por homini"
args163(11).Name = "SearchItem.ReplaceString"
args163(11).Value = "por la mulieri"
args163(12).Name = "SearchItem.Locale"
args163(12).Value = 255
args163(13).Name = "SearchItem.ChangedChars"
args163(13).Value = 2
args163(14).Name = "SearchItem.DeletedChars"
args163(14).Value = 2
args163(15).Name = "SearchItem.InsertedChars"
args163(15).Value = 2
args163(16).Name = "SearchItem.TransliterateFlags"
args163(16).Value = 1280
args163(17).Name = "SearchItem.Command"
args163(17).Value = 3
args163(18).Name = "Quiet"
args163(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args163())

rem ----------------------------------------------------------------------
dim args164(18) as new com.sun.star.beans.PropertyValue
args164(0).Name = "SearchItem.StyleFamily"
args164(0).Value = 2
args164(1).Name = "SearchItem.CellType"
args164(1).Value = 0
args164(2).Name = "SearchItem.RowDirection"
args164(2).Value = true
args164(3).Name = "SearchItem.AllTables"
args164(3).Value = false
args164(4).Name = "SearchItem.Backward"
args164(4).Value = false
args164(5).Name = "SearchItem.Pattern"
args164(5).Value = false
args164(6).Name = "SearchItem.Content"
args164(6).Value = false
args164(7).Name = "SearchItem.AsianOptions"
args164(7).Value = false
args164(8).Name = "SearchItem.AlgorithmType"
args164(8).Value = 0
args164(9).Name = "SearchItem.SearchFlags"
args164(9).Value = 65536
args164(10).Name = "SearchItem.SearchString"
args164(10).Value = " interakto"
args164(11).Name = "SearchItem.ReplaceString"
args164(11).Value = " interago"
args164(12).Name = "SearchItem.Locale"
args164(12).Value = 255
args164(13).Name = "SearchItem.ChangedChars"
args164(13).Value = 2
args164(14).Name = "SearchItem.DeletedChars"
args164(14).Value = 2
args164(15).Name = "SearchItem.InsertedChars"
args164(15).Value = 2
args164(16).Name = "SearchItem.TransliterateFlags"
args164(16).Value = 1280
args164(17).Name = "SearchItem.Command"
args164(17).Value = 3
args164(18).Name = "Quiet"
args164(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args164())

rem ----------------------------------------------------------------------
dim args165(18) as new com.sun.star.beans.PropertyValue
args165(0).Name = "SearchItem.StyleFamily"
args165(0).Value = 2
args165(1).Name = "SearchItem.CellType"
args165(1).Value = 0
args165(2).Name = "SearchItem.RowDirection"
args165(2).Value = true
args165(3).Name = "SearchItem.AllTables"
args165(3).Value = false
args165(4).Name = "SearchItem.Backward"
args165(4).Value = false
args165(5).Name = "SearchItem.Pattern"
args165(5).Value = false
args165(6).Name = "SearchItem.Content"
args165(6).Value = false
args165(7).Name = "SearchItem.AsianOptions"
args165(7).Value = false
args165(8).Name = "SearchItem.AlgorithmType"
args165(8).Value = 0
args165(9).Name = "SearchItem.SearchFlags"
args165(9).Value = 65536
args165(10).Name = "SearchItem.SearchString"
args165(10).Value = " e nove "
args165(11).Name = "SearchItem.ReplaceString"
args165(11).Value = " ed itere "
args165(12).Name = "SearchItem.Locale"
args165(12).Value = 255
args165(13).Name = "SearchItem.ChangedChars"
args165(13).Value = 2
args165(14).Name = "SearchItem.DeletedChars"
args165(14).Value = 2
args165(15).Name = "SearchItem.InsertedChars"
args165(15).Value = 2
args165(16).Name = "SearchItem.TransliterateFlags"
args165(16).Value = 1280
args165(17).Name = "SearchItem.Command"
args165(17).Value = 3
args165(18).Name = "Quiet"
args165(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args165())

rem ----------------------------------------------------------------------
dim args166(18) as new com.sun.star.beans.PropertyValue
args166(0).Name = "SearchItem.StyleFamily"
args166(0).Value = 2
args166(1).Name = "SearchItem.CellType"
args166(1).Value = 0
args166(2).Name = "SearchItem.RowDirection"
args166(2).Value = true
args166(3).Name = "SearchItem.AllTables"
args166(3).Value = false
args166(4).Name = "SearchItem.Backward"
args166(4).Value = false
args166(5).Name = "SearchItem.Pattern"
args166(5).Value = false
args166(6).Name = "SearchItem.Content"
args166(6).Value = false
args166(7).Name = "SearchItem.AsianOptions"
args166(7).Value = false
args166(8).Name = "SearchItem.AlgorithmType"
args166(8).Value = 0
args166(9).Name = "SearchItem.SearchFlags"
args166(9).Value = 65536
args166(10).Name = "SearchItem.SearchString"
args166(10).Value = "kompletis"
args166(11).Name = "SearchItem.ReplaceString"
args166(11).Value = "kompletigis"
args166(12).Name = "SearchItem.Locale"
args166(12).Value = 255
args166(13).Name = "SearchItem.ChangedChars"
args166(13).Value = 2
args166(14).Name = "SearchItem.DeletedChars"
args166(14).Value = 2
args166(15).Name = "SearchItem.InsertedChars"
args166(15).Value = 2
args166(16).Name = "SearchItem.TransliterateFlags"
args166(16).Value = 1280
args166(17).Name = "SearchItem.Command"
args166(17).Value = 3
args166(18).Name = "Quiet"
args166(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args166())

rem ----------------------------------------------------------------------
dim args167(18) as new com.sun.star.beans.PropertyValue
args167(0).Name = "SearchItem.StyleFamily"
args167(0).Value = 2
args167(1).Name = "SearchItem.CellType"
args167(1).Value = 0
args167(2).Name = "SearchItem.RowDirection"
args167(2).Value = true
args167(3).Name = "SearchItem.AllTables"
args167(3).Value = false
args167(4).Name = "SearchItem.Backward"
args167(4).Value = false
args167(5).Name = "SearchItem.Pattern"
args167(5).Value = false
args167(6).Name = "SearchItem.Content"
args167(6).Value = false
args167(7).Name = "SearchItem.AsianOptions"
args167(7).Value = false
args167(8).Name = "SearchItem.AlgorithmType"
args167(8).Value = 0
args167(9).Name = "SearchItem.SearchFlags"
args167(9).Value = 65536
args167(10).Name = "SearchItem.SearchString"
args167(10).Value = "maxima extenso"
args167(11).Name = "SearchItem.ReplaceString"
args167(11).Value = "maxima ampleso"
args167(12).Name = "SearchItem.Locale"
args167(12).Value = 255
args167(13).Name = "SearchItem.ChangedChars"
args167(13).Value = 2
args167(14).Name = "SearchItem.DeletedChars"
args167(14).Value = 2
args167(15).Name = "SearchItem.InsertedChars"
args167(15).Value = 2
args167(16).Name = "SearchItem.TransliterateFlags"
args167(16).Value = 1280
args167(17).Name = "SearchItem.Command"
args167(17).Value = 3
args167(18).Name = "Quiet"
args167(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args167())

rem ----------------------------------------------------------------------
dim args168(18) as new com.sun.star.beans.PropertyValue
args168(0).Name = "SearchItem.StyleFamily"
args168(0).Value = 2
args168(1).Name = "SearchItem.CellType"
args168(1).Value = 0
args168(2).Name = "SearchItem.RowDirection"
args168(2).Value = true
args168(3).Name = "SearchItem.AllTables"
args168(3).Value = false
args168(4).Name = "SearchItem.Backward"
args168(4).Value = false
args168(5).Name = "SearchItem.Pattern"
args168(5).Value = false
args168(6).Name = "SearchItem.Content"
args168(6).Value = false
args168(7).Name = "SearchItem.AsianOptions"
args168(7).Value = false
args168(8).Name = "SearchItem.AlgorithmType"
args168(8).Value = 0
args168(9).Name = "SearchItem.SearchFlags"
args168(9).Value = 65536
args168(10).Name = "SearchItem.SearchString"
args168(10).Value = "konquist"
args168(11).Name = "SearchItem.ReplaceString"
args168(11).Value = "konquest"
args168(12).Name = "SearchItem.Locale"
args168(12).Value = 255
args168(13).Name = "SearchItem.ChangedChars"
args168(13).Value = 2
args168(14).Name = "SearchItem.DeletedChars"
args168(14).Value = 2
args168(15).Name = "SearchItem.InsertedChars"
args168(15).Value = 2
args168(16).Name = "SearchItem.TransliterateFlags"
args168(16).Value = 1280
args168(17).Name = "SearchItem.Command"
args168(17).Value = 3
args168(18).Name = "Quiet"
args168(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args168())

rem ----------------------------------------------------------------------
dim args169(18) as new com.sun.star.beans.PropertyValue
args169(0).Name = "SearchItem.StyleFamily"
args169(0).Value = 2
args169(1).Name = "SearchItem.CellType"
args169(1).Value = 0
args169(2).Name = "SearchItem.RowDirection"
args169(2).Value = true
args169(3).Name = "SearchItem.AllTables"
args169(3).Value = false
args169(4).Name = "SearchItem.Backward"
args169(4).Value = false
args169(5).Name = "SearchItem.Pattern"
args169(5).Value = false
args169(6).Name = "SearchItem.Content"
args169(6).Value = false
args169(7).Name = "SearchItem.AsianOptions"
args169(7).Value = false
args169(8).Name = "SearchItem.AlgorithmType"
args169(8).Value = 0
args169(9).Name = "SearchItem.SearchFlags"
args169(9).Value = 65536
args169(10).Name = "SearchItem.SearchString"
args169(10).Value = "deskribas"
args169(11).Name = "SearchItem.ReplaceString"
args169(11).Value = "deskriptas"
args169(12).Name = "SearchItem.Locale"
args169(12).Value = 255
args169(13).Name = "SearchItem.ChangedChars"
args169(13).Value = 2
args169(14).Name = "SearchItem.DeletedChars"
args169(14).Value = 2
args169(15).Name = "SearchItem.InsertedChars"
args169(15).Value = 2
args169(16).Name = "SearchItem.TransliterateFlags"
args169(16).Value = 1280
args169(17).Name = "SearchItem.Command"
args169(17).Value = 3
args169(18).Name = "Quiet"
args169(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args169())

rem ----------------------------------------------------------------------
dim args170(18) as new com.sun.star.beans.PropertyValue
args170(0).Name = "SearchItem.StyleFamily"
args170(0).Value = 2
args170(1).Name = "SearchItem.CellType"
args170(1).Value = 0
args170(2).Name = "SearchItem.RowDirection"
args170(2).Value = true
args170(3).Name = "SearchItem.AllTables"
args170(3).Value = false
args170(4).Name = "SearchItem.Backward"
args170(4).Value = false
args170(5).Name = "SearchItem.Pattern"
args170(5).Value = false
args170(6).Name = "SearchItem.Content"
args170(6).Value = false
args170(7).Name = "SearchItem.AsianOptions"
args170(7).Value = false
args170(8).Name = "SearchItem.AlgorithmType"
args170(8).Value = 0
args170(9).Name = "SearchItem.SearchFlags"
args170(9).Value = 65536
args170(10).Name = "SearchItem.SearchString"
args170(10).Value = "informala"
args170(11).Name = "SearchItem.ReplaceString"
args170(11).Value = "neformala"
args170(12).Name = "SearchItem.Locale"
args170(12).Value = 255
args170(13).Name = "SearchItem.ChangedChars"
args170(13).Value = 2
args170(14).Name = "SearchItem.DeletedChars"
args170(14).Value = 2
args170(15).Name = "SearchItem.InsertedChars"
args170(15).Value = 2
args170(16).Name = "SearchItem.TransliterateFlags"
args170(16).Value = 1280
args170(17).Name = "SearchItem.Command"
args170(17).Value = 3
args170(18).Name = "Quiet"
args170(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args170())

rem ----------------------------------------------------------------------
dim args171(18) as new com.sun.star.beans.PropertyValue
args171(0).Name = "SearchItem.StyleFamily"
args171(0).Value = 2
args171(1).Name = "SearchItem.CellType"
args171(1).Value = 0
args171(2).Name = "SearchItem.RowDirection"
args171(2).Value = true
args171(3).Name = "SearchItem.AllTables"
args171(3).Value = false
args171(4).Name = "SearchItem.Backward"
args171(4).Value = false
args171(5).Name = "SearchItem.Pattern"
args171(5).Value = false
args171(6).Name = "SearchItem.Content"
args171(6).Value = false
args171(7).Name = "SearchItem.AsianOptions"
args171(7).Value = false
args171(8).Name = "SearchItem.AlgorithmType"
args171(8).Value = 0
args171(9).Name = "SearchItem.SearchFlags"
args171(9).Value = 65536
args171(10).Name = "SearchItem.SearchString"
args171(10).Value = "deskribita"
args171(11).Name = "SearchItem.ReplaceString"
args171(11).Value = "deskriptita"
args171(12).Name = "SearchItem.Locale"
args171(12).Value = 255
args171(13).Name = "SearchItem.ChangedChars"
args171(13).Value = 2
args171(14).Name = "SearchItem.DeletedChars"
args171(14).Value = 2
args171(15).Name = "SearchItem.InsertedChars"
args171(15).Value = 2
args171(16).Name = "SearchItem.TransliterateFlags"
args171(16).Value = 1280
args171(17).Name = "SearchItem.Command"
args171(17).Value = 3
args171(18).Name = "Quiet"
args171(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args171())

rem ----------------------------------------------------------------------
dim args172(18) as new com.sun.star.beans.PropertyValue
args172(0).Name = "SearchItem.StyleFamily"
args172(0).Value = 2
args172(1).Name = "SearchItem.CellType"
args172(1).Value = 0
args172(2).Name = "SearchItem.RowDirection"
args172(2).Value = true
args172(3).Name = "SearchItem.AllTables"
args172(3).Value = false
args172(4).Name = "SearchItem.Backward"
args172(4).Value = false
args172(5).Name = "SearchItem.Pattern"
args172(5).Value = false
args172(6).Name = "SearchItem.Content"
args172(6).Value = false
args172(7).Name = "SearchItem.AsianOptions"
args172(7).Value = false
args172(8).Name = "SearchItem.AlgorithmType"
args172(8).Value = 0
args172(9).Name = "SearchItem.SearchFlags"
args172(9).Value = 65536
args172(10).Name = "SearchItem.SearchString"
args172(10).Value = "Lasta-Diala Santi"
args172(11).Name = "SearchItem.ReplaceString"
args172(11).Value = "Lasta-Dia Santi"
args172(12).Name = "SearchItem.Locale"
args172(12).Value = 255
args172(13).Name = "SearchItem.ChangedChars"
args172(13).Value = 2
args172(14).Name = "SearchItem.DeletedChars"
args172(14).Value = 2
args172(15).Name = "SearchItem.InsertedChars"
args172(15).Value = 2
args172(16).Name = "SearchItem.TransliterateFlags"
args172(16).Value = 1280
args172(17).Name = "SearchItem.Command"
args172(17).Value = 3
args172(18).Name = "Quiet"
args172(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args172())

rem ----------------------------------------------------------------------
dim args173(18) as new com.sun.star.beans.PropertyValue
args173(0).Name = "SearchItem.StyleFamily"
args173(0).Value = 2
args173(1).Name = "SearchItem.CellType"
args173(1).Value = 0
args173(2).Name = "SearchItem.RowDirection"
args173(2).Value = true
args173(3).Name = "SearchItem.AllTables"
args173(3).Value = false
args173(4).Name = "SearchItem.Backward"
args173(4).Value = false
args173(5).Name = "SearchItem.Pattern"
args173(5).Value = false
args173(6).Name = "SearchItem.Content"
args173(6).Value = false
args173(7).Name = "SearchItem.AsianOptions"
args173(7).Value = false
args173(8).Name = "SearchItem.AlgorithmType"
args173(8).Value = 0
args173(9).Name = "SearchItem.SearchFlags"
args173(9).Value = 65536
args173(10).Name = "SearchItem.SearchString"
args173(10).Value = "permezas"
args173(11).Name = "SearchItem.ReplaceString"
args173(11).Value = "permisas"
args173(12).Name = "SearchItem.Locale"
args173(12).Value = 255
args173(13).Name = "SearchItem.ChangedChars"
args173(13).Value = 2
args173(14).Name = "SearchItem.DeletedChars"
args173(14).Value = 2
args173(15).Name = "SearchItem.InsertedChars"
args173(15).Value = 2
args173(16).Name = "SearchItem.TransliterateFlags"
args173(16).Value = 1280
args173(17).Name = "SearchItem.Command"
args173(17).Value = 3
args173(18).Name = "Quiet"
args173(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args173())

rem ----------------------------------------------------------------------
dim args174(18) as new com.sun.star.beans.PropertyValue
args174(0).Name = "SearchItem.StyleFamily"
args174(0).Value = 2
args174(1).Name = "SearchItem.CellType"
args174(1).Value = 0
args174(2).Name = "SearchItem.RowDirection"
args174(2).Value = true
args174(3).Name = "SearchItem.AllTables"
args174(3).Value = false
args174(4).Name = "SearchItem.Backward"
args174(4).Value = false
args174(5).Name = "SearchItem.Pattern"
args174(5).Value = false
args174(6).Name = "SearchItem.Content"
args174(6).Value = false
args174(7).Name = "SearchItem.AsianOptions"
args174(7).Value = false
args174(8).Name = "SearchItem.AlgorithmType"
args174(8).Value = 0
args174(9).Name = "SearchItem.SearchFlags"
args174(9).Value = 65536
args174(10).Name = "SearchItem.SearchString"
args174(10).Value = "deskribita"
args174(11).Name = "SearchItem.ReplaceString"
args174(11).Value = "deskriptita"
args174(12).Name = "SearchItem.Locale"
args174(12).Value = 255
args174(13).Name = "SearchItem.ChangedChars"
args174(13).Value = 2
args174(14).Name = "SearchItem.DeletedChars"
args174(14).Value = 2
args174(15).Name = "SearchItem.InsertedChars"
args174(15).Value = 2
args174(16).Name = "SearchItem.TransliterateFlags"
args174(16).Value = 1280
args174(17).Name = "SearchItem.Command"
args174(17).Value = 3
args174(18).Name = "Quiet"
args174(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args174())

rem ----------------------------------------------------------------------
dim args175(18) as new com.sun.star.beans.PropertyValue
args175(0).Name = "SearchItem.StyleFamily"
args175(0).Value = 2
args175(1).Name = "SearchItem.CellType"
args175(1).Value = 0
args175(2).Name = "SearchItem.RowDirection"
args175(2).Value = true
args175(3).Name = "SearchItem.AllTables"
args175(3).Value = false
args175(4).Name = "SearchItem.Backward"
args175(4).Value = false
args175(5).Name = "SearchItem.Pattern"
args175(5).Value = false
args175(6).Name = "SearchItem.Content"
args175(6).Value = false
args175(7).Name = "SearchItem.AsianOptions"
args175(7).Value = false
args175(8).Name = "SearchItem.AlgorithmType"
args175(8).Value = 0
args175(9).Name = "SearchItem.SearchFlags"
args175(9).Value = 65536
args175(10).Name = "SearchItem.SearchString"
args175(10).Value = "konocita kam"
args175(11).Name = "SearchItem.ReplaceString"
args175(11).Value = "konocita kom"
args175(12).Name = "SearchItem.Locale"
args175(12).Value = 255
args175(13).Name = "SearchItem.ChangedChars"
args175(13).Value = 2
args175(14).Name = "SearchItem.DeletedChars"
args175(14).Value = 2
args175(15).Name = "SearchItem.InsertedChars"
args175(15).Value = 2
args175(16).Name = "SearchItem.TransliterateFlags"
args175(16).Value = 1280
args175(17).Name = "SearchItem.Command"
args175(17).Value = 3
args175(18).Name = "Quiet"
args175(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args175())

rem ----------------------------------------------------------------------
dim args176(18) as new com.sun.star.beans.PropertyValue
args176(0).Name = "SearchItem.StyleFamily"
args176(0).Value = 2
args176(1).Name = "SearchItem.CellType"
args176(1).Value = 0
args176(2).Name = "SearchItem.RowDirection"
args176(2).Value = true
args176(3).Name = "SearchItem.AllTables"
args176(3).Value = false
args176(4).Name = "SearchItem.Backward"
args176(4).Value = false
args176(5).Name = "SearchItem.Pattern"
args176(5).Value = false
args176(6).Name = "SearchItem.Content"
args176(6).Value = false
args176(7).Name = "SearchItem.AsianOptions"
args176(7).Value = false
args176(8).Name = "SearchItem.AlgorithmType"
args176(8).Value = 0
args176(9).Name = "SearchItem.SearchFlags"
args176(9).Value = 65536
args176(10).Name = "SearchItem.SearchString"
args176(10).Value = "manjaji"
args176(11).Name = "SearchItem.ReplaceString"
args176(11).Value = "nutrivi"
args176(12).Name = "SearchItem.Locale"
args176(12).Value = 255
args176(13).Name = "SearchItem.ChangedChars"
args176(13).Value = 2
args176(14).Name = "SearchItem.DeletedChars"
args176(14).Value = 2
args176(15).Name = "SearchItem.InsertedChars"
args176(15).Value = 2
args176(16).Name = "SearchItem.TransliterateFlags"
args176(16).Value = 1280
args176(17).Name = "SearchItem.Command"
args176(17).Value = 3
args176(18).Name = "Quiet"
args176(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args176())

rem ----------------------------------------------------------------------
dim args177(18) as new com.sun.star.beans.PropertyValue
args177(0).Name = "SearchItem.StyleFamily"
args177(0).Value = 2
args177(1).Name = "SearchItem.CellType"
args177(1).Value = 0
args177(2).Name = "SearchItem.RowDirection"
args177(2).Value = true
args177(3).Name = "SearchItem.AllTables"
args177(3).Value = false
args177(4).Name = "SearchItem.Backward"
args177(4).Value = false
args177(5).Name = "SearchItem.Pattern"
args177(5).Value = false
args177(6).Name = "SearchItem.Content"
args177(6).Value = false
args177(7).Name = "SearchItem.AsianOptions"
args177(7).Value = false
args177(8).Name = "SearchItem.AlgorithmType"
args177(8).Value = 0
args177(9).Name = "SearchItem.SearchFlags"
args177(9).Value = 65536
args177(10).Name = "SearchItem.SearchString"
args177(10).Value = " nove "
args177(11).Name = "SearchItem.ReplaceString"
args177(11).Value = " itere "
args177(12).Name = "SearchItem.Locale"
args177(12).Value = 255
args177(13).Name = "SearchItem.ChangedChars"
args177(13).Value = 2
args177(14).Name = "SearchItem.DeletedChars"
args177(14).Value = 2
args177(15).Name = "SearchItem.InsertedChars"
args177(15).Value = 2
args177(16).Name = "SearchItem.TransliterateFlags"
args177(16).Value = 1280
args177(17).Name = "SearchItem.Command"
args177(17).Value = 3
args177(18).Name = "Quiet"
args177(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args177())

rem ----------------------------------------------------------------------
dim args178(18) as new com.sun.star.beans.PropertyValue
args178(0).Name = "SearchItem.StyleFamily"
args178(0).Value = 2
args178(1).Name = "SearchItem.CellType"
args178(1).Value = 0
args178(2).Name = "SearchItem.RowDirection"
args178(2).Value = true
args178(3).Name = "SearchItem.AllTables"
args178(3).Value = false
args178(4).Name = "SearchItem.Backward"
args178(4).Value = false
args178(5).Name = "SearchItem.Pattern"
args178(5).Value = false
args178(6).Name = "SearchItem.Content"
args178(6).Value = false
args178(7).Name = "SearchItem.AsianOptions"
args178(7).Value = false
args178(8).Name = "SearchItem.AlgorithmType"
args178(8).Value = 0
args178(9).Name = "SearchItem.SearchFlags"
args178(9).Value = 65536
args178(10).Name = "SearchItem.SearchString"
args178(10).Value = "formesis"
args178(11).Name = "SearchItem.ReplaceString"
args178(11).Value = "formacesis"
args178(12).Name = "SearchItem.Locale"
args178(12).Value = 255
args178(13).Name = "SearchItem.ChangedChars"
args178(13).Value = 2
args178(14).Name = "SearchItem.DeletedChars"
args178(14).Value = 2
args178(15).Name = "SearchItem.InsertedChars"
args178(15).Value = 2
args178(16).Name = "SearchItem.TransliterateFlags"
args178(16).Value = 1280
args178(17).Name = "SearchItem.Command"
args178(17).Value = 3
args178(18).Name = "Quiet"
args178(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args178())

rem ----------------------------------------------------------------------
dim args179(18) as new com.sun.star.beans.PropertyValue
args179(0).Name = "SearchItem.StyleFamily"
args179(0).Value = 2
args179(1).Name = "SearchItem.CellType"
args179(1).Value = 0
args179(2).Name = "SearchItem.RowDirection"
args179(2).Value = true
args179(3).Name = "SearchItem.AllTables"
args179(3).Value = false
args179(4).Name = "SearchItem.Backward"
args179(4).Value = false
args179(5).Name = "SearchItem.Pattern"
args179(5).Value = false
args179(6).Name = "SearchItem.Content"
args179(6).Value = false
args179(7).Name = "SearchItem.AsianOptions"
args179(7).Value = false
args179(8).Name = "SearchItem.AlgorithmType"
args179(8).Value = 0
args179(9).Name = "SearchItem.SearchFlags"
args179(9).Value = 65536
args179(10).Name = "SearchItem.SearchString"
args179(10).Value = "publikis"
args179(11).Name = "SearchItem.ReplaceString"
args179(11).Value = "publikigis"
args179(12).Name = "SearchItem.Locale"
args179(12).Value = 255
args179(13).Name = "SearchItem.ChangedChars"
args179(13).Value = 2
args179(14).Name = "SearchItem.DeletedChars"
args179(14).Value = 2
args179(15).Name = "SearchItem.InsertedChars"
args179(15).Value = 2
args179(16).Name = "SearchItem.TransliterateFlags"
args179(16).Value = 1280
args179(17).Name = "SearchItem.Command"
args179(17).Value = 3
args179(18).Name = "Quiet"
args179(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args179())

rem ----------------------------------------------------------------------
dim args180(18) as new com.sun.star.beans.PropertyValue
args180(0).Name = "SearchItem.StyleFamily"
args180(0).Value = 2
args180(1).Name = "SearchItem.CellType"
args180(1).Value = 0
args180(2).Name = "SearchItem.RowDirection"
args180(2).Value = true
args180(3).Name = "SearchItem.AllTables"
args180(3).Value = false
args180(4).Name = "SearchItem.Backward"
args180(4).Value = false
args180(5).Name = "SearchItem.Pattern"
args180(5).Value = false
args180(6).Name = "SearchItem.Content"
args180(6).Value = false
args180(7).Name = "SearchItem.AsianOptions"
args180(7).Value = false
args180(8).Name = "SearchItem.AlgorithmType"
args180(8).Value = 0
args180(9).Name = "SearchItem.SearchFlags"
args180(9).Value = 65536
args180(10).Name = "SearchItem.SearchString"
args180(10).Value = "mikrabiologiisto"
args180(11).Name = "SearchItem.ReplaceString"
args180(11).Value = "mikrobiologiisto"
args180(12).Name = "SearchItem.Locale"
args180(12).Value = 255
args180(13).Name = "SearchItem.ChangedChars"
args180(13).Value = 2
args180(14).Name = "SearchItem.DeletedChars"
args180(14).Value = 2
args180(15).Name = "SearchItem.InsertedChars"
args180(15).Value = 2
args180(16).Name = "SearchItem.TransliterateFlags"
args180(16).Value = 1280
args180(17).Name = "SearchItem.Command"
args180(17).Value = 3
args180(18).Name = "Quiet"
args180(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args180())

rem ----------------------------------------------------------------------
dim args181(18) as new com.sun.star.beans.PropertyValue
args181(0).Name = "SearchItem.StyleFamily"
args181(0).Value = 2
args181(1).Name = "SearchItem.CellType"
args181(1).Value = 0
args181(2).Name = "SearchItem.RowDirection"
args181(2).Value = true
args181(3).Name = "SearchItem.AllTables"
args181(3).Value = false
args181(4).Name = "SearchItem.Backward"
args181(4).Value = false
args181(5).Name = "SearchItem.Pattern"
args181(5).Value = false
args181(6).Name = "SearchItem.Content"
args181(6).Value = false
args181(7).Name = "SearchItem.AsianOptions"
args181(7).Value = false
args181(8).Name = "SearchItem.AlgorithmType"
args181(8).Value = 0
args181(9).Name = "SearchItem.SearchFlags"
args181(9).Value = 65536
args181(10).Name = "SearchItem.SearchString"
args181(10).Value = "stadiono"
args181(11).Name = "SearchItem.ReplaceString"
args181(11).Value = "stadio"
args181(12).Name = "SearchItem.Locale"
args181(12).Value = 255
args181(13).Name = "SearchItem.ChangedChars"
args181(13).Value = 2
args181(14).Name = "SearchItem.DeletedChars"
args181(14).Value = 2
args181(15).Name = "SearchItem.InsertedChars"
args181(15).Value = 2
args181(16).Name = "SearchItem.TransliterateFlags"
args181(16).Value = 1280
args181(17).Name = "SearchItem.Command"
args181(17).Value = 3
args181(18).Name = "Quiet"
args181(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args181())

rem ----------------------------------------------------------------------
dim args182(18) as new com.sun.star.beans.PropertyValue
args182(0).Name = "SearchItem.StyleFamily"
args182(0).Value = 2
args182(1).Name = "SearchItem.CellType"
args182(1).Value = 0
args182(2).Name = "SearchItem.RowDirection"
args182(2).Value = true
args182(3).Name = "SearchItem.AllTables"
args182(3).Value = false
args182(4).Name = "SearchItem.Backward"
args182(4).Value = false
args182(5).Name = "SearchItem.Pattern"
args182(5).Value = false
args182(6).Name = "SearchItem.Content"
args182(6).Value = false
args182(7).Name = "SearchItem.AsianOptions"
args182(7).Value = false
args182(8).Name = "SearchItem.AlgorithmType"
args182(8).Value = 0
args182(9).Name = "SearchItem.SearchFlags"
args182(9).Value = 65536
args182(10).Name = "SearchItem.SearchString"
args182(10).Value = "boviari"
args182(11).Name = "SearchItem.ReplaceString"
args182(11).Value = "bovaro"
args182(12).Name = "SearchItem.Locale"
args182(12).Value = 255
args182(13).Name = "SearchItem.ChangedChars"
args182(13).Value = 2
args182(14).Name = "SearchItem.DeletedChars"
args182(14).Value = 2
args182(15).Name = "SearchItem.InsertedChars"
args182(15).Value = 2
args182(16).Name = "SearchItem.TransliterateFlags"
args182(16).Value = 1280
args182(17).Name = "SearchItem.Command"
args182(17).Value = 3
args182(18).Name = "Quiet"
args182(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args182())

rem ----------------------------------------------------------------------
dim args183(18) as new com.sun.star.beans.PropertyValue
args183(0).Name = "SearchItem.StyleFamily"
args183(0).Value = 2
args183(1).Name = "SearchItem.CellType"
args183(1).Value = 0
args183(2).Name = "SearchItem.RowDirection"
args183(2).Value = true
args183(3).Name = "SearchItem.AllTables"
args183(3).Value = false
args183(4).Name = "SearchItem.Backward"
args183(4).Value = false
args183(5).Name = "SearchItem.Pattern"
args183(5).Value = false
args183(6).Name = "SearchItem.Content"
args183(6).Value = false
args183(7).Name = "SearchItem.AsianOptions"
args183(7).Value = false
args183(8).Name = "SearchItem.AlgorithmType"
args183(8).Value = 0
args183(9).Name = "SearchItem.SearchFlags"
args183(9).Value = 65536
args183(10).Name = "SearchItem.SearchString"
args183(10).Value = "bovari"
args183(11).Name = "SearchItem.ReplaceString"
args183(11).Value = "bovaro"
args183(12).Name = "SearchItem.Locale"
args183(12).Value = 255
args183(13).Name = "SearchItem.ChangedChars"
args183(13).Value = 2
args183(14).Name = "SearchItem.DeletedChars"
args183(14).Value = 2
args183(15).Name = "SearchItem.InsertedChars"
args183(15).Value = 2
args183(16).Name = "SearchItem.TransliterateFlags"
args183(16).Value = 1280
args183(17).Name = "SearchItem.Command"
args183(17).Value = 3
args183(18).Name = "Quiet"
args183(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args183())

rem ----------------------------------------------------------------------
dim args184(18) as new com.sun.star.beans.PropertyValue
args184(0).Name = "SearchItem.StyleFamily"
args184(0).Value = 2
args184(1).Name = "SearchItem.CellType"
args184(1).Value = 0
args184(2).Name = "SearchItem.RowDirection"
args184(2).Value = true
args184(3).Name = "SearchItem.AllTables"
args184(3).Value = false
args184(4).Name = "SearchItem.Backward"
args184(4).Value = false
args184(5).Name = "SearchItem.Pattern"
args184(5).Value = false
args184(6).Name = "SearchItem.Content"
args184(6).Value = false
args184(7).Name = "SearchItem.AsianOptions"
args184(7).Value = false
args184(8).Name = "SearchItem.AlgorithmType"
args184(8).Value = 0
args184(9).Name = "SearchItem.SearchFlags"
args184(9).Value = 65536
args184(10).Name = "SearchItem.SearchString"
args184(10).Value = " bazita en"
args184(11).Name = "SearchItem.ReplaceString"
args184(11).Value = " apogata sur"
args184(12).Name = "SearchItem.Locale"
args184(12).Value = 255
args184(13).Name = "SearchItem.ChangedChars"
args184(13).Value = 2
args184(14).Name = "SearchItem.DeletedChars"
args184(14).Value = 2
args184(15).Name = "SearchItem.InsertedChars"
args184(15).Value = 2
args184(16).Name = "SearchItem.TransliterateFlags"
args184(16).Value = 1280
args184(17).Name = "SearchItem.Command"
args184(17).Value = 3
args184(18).Name = "Quiet"
args184(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args184())

rem ----------------------------------------------------------------------
dim args185(18) as new com.sun.star.beans.PropertyValue
args185(0).Name = "SearchItem.StyleFamily"
args185(0).Value = 2
args185(1).Name = "SearchItem.CellType"
args185(1).Value = 0
args185(2).Name = "SearchItem.RowDirection"
args185(2).Value = true
args185(3).Name = "SearchItem.AllTables"
args185(3).Value = false
args185(4).Name = "SearchItem.Backward"
args185(4).Value = false
args185(5).Name = "SearchItem.Pattern"
args185(5).Value = false
args185(6).Name = "SearchItem.Content"
args185(6).Value = false
args185(7).Name = "SearchItem.AsianOptions"
args185(7).Value = false
args185(8).Name = "SearchItem.AlgorithmType"
args185(8).Value = 0
args185(9).Name = "SearchItem.SearchFlags"
args185(9).Value = 65536
args185(10).Name = "SearchItem.SearchString"
args185(10).Value = "transformaces"
args185(11).Name = "SearchItem.ReplaceString"
args185(11).Value = "transformes"
args185(12).Name = "SearchItem.Locale"
args185(12).Value = 255
args185(13).Name = "SearchItem.ChangedChars"
args185(13).Value = 2
args185(14).Name = "SearchItem.DeletedChars"
args185(14).Value = 2
args185(15).Name = "SearchItem.InsertedChars"
args185(15).Value = 2
args185(16).Name = "SearchItem.TransliterateFlags"
args185(16).Value = 1280
args185(17).Name = "SearchItem.Command"
args185(17).Value = 3
args185(18).Name = "Quiet"
args185(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args185())

rem ----------------------------------------------------------------------
dim args186(18) as new com.sun.star.beans.PropertyValue
args186(0).Name = "SearchItem.StyleFamily"
args186(0).Value = 2
args186(1).Name = "SearchItem.CellType"
args186(1).Value = 0
args186(2).Name = "SearchItem.RowDirection"
args186(2).Value = true
args186(3).Name = "SearchItem.AllTables"
args186(3).Value = false
args186(4).Name = "SearchItem.Backward"
args186(4).Value = false
args186(5).Name = "SearchItem.Pattern"
args186(5).Value = false
args186(6).Name = "SearchItem.Content"
args186(6).Value = false
args186(7).Name = "SearchItem.AsianOptions"
args186(7).Value = false
args186(8).Name = "SearchItem.AlgorithmType"
args186(8).Value = 0
args186(9).Name = "SearchItem.SearchFlags"
args186(9).Value = 65536
args186(10).Name = "SearchItem.SearchString"
args186(10).Value = "veloceso"
args186(11).Name = "SearchItem.ReplaceString"
args186(11).Value = "rapideso"
args186(12).Name = "SearchItem.Locale"
args186(12).Value = 255
args186(13).Name = "SearchItem.ChangedChars"
args186(13).Value = 2
args186(14).Name = "SearchItem.DeletedChars"
args186(14).Value = 2
args186(15).Name = "SearchItem.InsertedChars"
args186(15).Value = 2
args186(16).Name = "SearchItem.TransliterateFlags"
args186(16).Value = 1280
args186(17).Name = "SearchItem.Command"
args186(17).Value = 3
args186(18).Name = "Quiet"
args186(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args186())

rem ----------------------------------------------------------------------
dim args187(18) as new com.sun.star.beans.PropertyValue
args187(0).Name = "SearchItem.StyleFamily"
args187(0).Value = 2
args187(1).Name = "SearchItem.CellType"
args187(1).Value = 0
args187(2).Name = "SearchItem.RowDirection"
args187(2).Value = true
args187(3).Name = "SearchItem.AllTables"
args187(3).Value = false
args187(4).Name = "SearchItem.Backward"
args187(4).Value = false
args187(5).Name = "SearchItem.Pattern"
args187(5).Value = false
args187(6).Name = "SearchItem.Content"
args187(6).Value = false
args187(7).Name = "SearchItem.AsianOptions"
args187(7).Value = false
args187(8).Name = "SearchItem.AlgorithmType"
args187(8).Value = 0
args187(9).Name = "SearchItem.SearchFlags"
args187(9).Value = 65536
args187(10).Name = "SearchItem.SearchString"
args187(10).Value = "Ukrainia"
args187(11).Name = "SearchItem.ReplaceString"
args187(11).Value = "Ukraina"
args187(12).Name = "SearchItem.Locale"
args187(12).Value = 255
args187(13).Name = "SearchItem.ChangedChars"
args187(13).Value = 2
args187(14).Name = "SearchItem.DeletedChars"
args187(14).Value = 2
args187(15).Name = "SearchItem.InsertedChars"
args187(15).Value = 2
args187(16).Name = "SearchItem.TransliterateFlags"
args187(16).Value = 1280
args187(17).Name = "SearchItem.Command"
args187(17).Value = 3
args187(18).Name = "Quiet"
args187(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args187())

rem ----------------------------------------------------------------------
dim args188(18) as new com.sun.star.beans.PropertyValue
args188(0).Name = "SearchItem.StyleFamily"
args188(0).Value = 2
args188(1).Name = "SearchItem.CellType"
args188(1).Value = 0
args188(2).Name = "SearchItem.RowDirection"
args188(2).Value = true
args188(3).Name = "SearchItem.AllTables"
args188(3).Value = false
args188(4).Name = "SearchItem.Backward"
args188(4).Value = false
args188(5).Name = "SearchItem.Pattern"
args188(5).Value = false
args188(6).Name = "SearchItem.Content"
args188(6).Value = false
args188(7).Name = "SearchItem.AsianOptions"
args188(7).Value = false
args188(8).Name = "SearchItem.AlgorithmType"
args188(8).Value = 0
args188(9).Name = "SearchItem.SearchFlags"
args188(9).Value = 65536
args188(10).Name = "SearchItem.SearchString"
args188(10).Value = "Saudia Arabia"
args188(11).Name = "SearchItem.ReplaceString"
args188(11).Value = "Saudi-Arabia"
args188(12).Name = "SearchItem.Locale"
args188(12).Value = 255
args188(13).Name = "SearchItem.ChangedChars"
args188(13).Value = 2
args188(14).Name = "SearchItem.DeletedChars"
args188(14).Value = 2
args188(15).Name = "SearchItem.InsertedChars"
args188(15).Value = 2
args188(16).Name = "SearchItem.TransliterateFlags"
args188(16).Value = 1280
args188(17).Name = "SearchItem.Command"
args188(17).Value = 3
args188(18).Name = "Quiet"
args188(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args188())

rem ----------------------------------------------------------------------
dim args189(18) as new com.sun.star.beans.PropertyValue
args189(0).Name = "SearchItem.StyleFamily"
args189(0).Value = 2
args189(1).Name = "SearchItem.CellType"
args189(1).Value = 0
args189(2).Name = "SearchItem.RowDirection"
args189(2).Value = true
args189(3).Name = "SearchItem.AllTables"
args189(3).Value = false
args189(4).Name = "SearchItem.Backward"
args189(4).Value = false
args189(5).Name = "SearchItem.Pattern"
args189(5).Value = false
args189(6).Name = "SearchItem.Content"
args189(6).Value = false
args189(7).Name = "SearchItem.AsianOptions"
args189(7).Value = false
args189(8).Name = "SearchItem.AlgorithmType"
args189(8).Value = 0
args189(9).Name = "SearchItem.SearchFlags"
args189(9).Value = 65536
args189(10).Name = "SearchItem.SearchString"
args189(10).Value = "frequenta"
args189(11).Name = "SearchItem.ReplaceString"
args189(11).Value = "ofta"
args189(12).Name = "SearchItem.Locale"
args189(12).Value = 255
args189(13).Name = "SearchItem.ChangedChars"
args189(13).Value = 2
args189(14).Name = "SearchItem.DeletedChars"
args189(14).Value = 2
args189(15).Name = "SearchItem.InsertedChars"
args189(15).Value = 2
args189(16).Name = "SearchItem.TransliterateFlags"
args189(16).Value = 1280
args189(17).Name = "SearchItem.Command"
args189(17).Value = 3
args189(18).Name = "Quiet"
args189(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args189())

rem ----------------------------------------------------------------------
dim args190(18) as new com.sun.star.beans.PropertyValue
args190(0).Name = "SearchItem.StyleFamily"
args190(0).Value = 2
args190(1).Name = "SearchItem.CellType"
args190(1).Value = 0
args190(2).Name = "SearchItem.RowDirection"
args190(2).Value = true
args190(3).Name = "SearchItem.AllTables"
args190(3).Value = false
args190(4).Name = "SearchItem.Backward"
args190(4).Value = false
args190(5).Name = "SearchItem.Pattern"
args190(5).Value = false
args190(6).Name = "SearchItem.Content"
args190(6).Value = false
args190(7).Name = "SearchItem.AsianOptions"
args190(7).Value = false
args190(8).Name = "SearchItem.AlgorithmType"
args190(8).Value = 0
args190(9).Name = "SearchItem.SearchFlags"
args190(9).Value = 65536
args190(10).Name = "SearchItem.SearchString"
args190(10).Value = "terorigant"
args190(11).Name = "SearchItem.ReplaceString"
args190(11).Value = "terorist"
args190(12).Name = "SearchItem.Locale"
args190(12).Value = 255
args190(13).Name = "SearchItem.ChangedChars"
args190(13).Value = 2
args190(14).Name = "SearchItem.DeletedChars"
args190(14).Value = 2
args190(15).Name = "SearchItem.InsertedChars"
args190(15).Value = 2
args190(16).Name = "SearchItem.TransliterateFlags"
args190(16).Value = 1280
args190(17).Name = "SearchItem.Command"
args190(17).Value = 3
args190(18).Name = "Quiet"
args190(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args190())

rem ----------------------------------------------------------------------
dim args191(18) as new com.sun.star.beans.PropertyValue

dim MARTO1, MARTO2 as string
For I = 1 to 31
MARTO1 = Ltrim(Str$(I))+" di marto"
MARTO2 = Ltrim(Str$(I))+"ma di marto"

args191(0).Name = "SearchItem.StyleFamily"
args191(0).Value = 2
args191(1).Name = "SearchItem.CellType"
args191(1).Value = 0
args191(2).Name = "SearchItem.RowDirection"
args191(2).Value = true
args191(3).Name = "SearchItem.AllTables"
args191(3).Value = false
args191(4).Name = "SearchItem.Backward"
args191(4).Value = false
args191(5).Name = "SearchItem.Pattern"
args191(5).Value = false
args191(6).Name = "SearchItem.Content"
args191(6).Value = false
args191(7).Name = "SearchItem.AsianOptions"
args191(7).Value = false
args191(8).Name = "SearchItem.AlgorithmType"
args191(8).Value = 0
args191(9).Name = "SearchItem.SearchFlags"
args191(9).Value = 65536
args191(10).Name = "SearchItem.SearchString"
args191(10).Value = MARTO1
args191(11).Name = "SearchItem.ReplaceString"
args191(11).Value = MARTO2
args191(12).Name = "SearchItem.Locale"
args191(12).Value = 255
args191(13).Name = "SearchItem.ChangedChars"
args191(13).Value = 2
args191(14).Name = "SearchItem.DeletedChars"
args191(14).Value = 2
args191(15).Name = "SearchItem.InsertedChars"
args191(15).Value = 2
args191(16).Name = "SearchItem.TransliterateFlags"
args191(16).Value = 1280
args191(17).Name = "SearchItem.Command"
args191(17).Value = 3
args191(18).Name = "Quiet"
args191(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args191())

Next I

rem ----------------------------------------------------------------------
dim args192(18) as new com.sun.star.beans.PropertyValue
args192(0).Name = "SearchItem.StyleFamily"
args192(0).Value = 2
args192(1).Name = "SearchItem.CellType"
args192(1).Value = 0
args192(2).Name = "SearchItem.RowDirection"
args192(2).Value = true
args192(3).Name = "SearchItem.AllTables"
args192(3).Value = false
args192(4).Name = "SearchItem.Backward"
args192(4).Value = false
args192(5).Name = "SearchItem.Pattern"
args192(5).Value = false
args192(6).Name = "SearchItem.Content"
args192(6).Value = false
args192(7).Name = "SearchItem.AsianOptions"
args192(7).Value = false
args192(8).Name = "SearchItem.AlgorithmType"
args192(8).Value = 0
args192(9).Name = "SearchItem.SearchFlags"
args192(9).Value = 65536
args192(10).Name = "SearchItem.SearchString"
args192(10).Value = "reporto"
args192(11).Name = "SearchItem.ReplaceString"
args192(11).Value = "raporto"
args192(12).Name = "SearchItem.Locale"
args192(12).Value = 255
args192(13).Name = "SearchItem.ChangedChars"
args192(13).Value = 2
args192(14).Name = "SearchItem.DeletedChars"
args192(14).Value = 2
args192(15).Name = "SearchItem.InsertedChars"
args192(15).Value = 2
args192(16).Name = "SearchItem.TransliterateFlags"
args192(16).Value = 1280
args192(17).Name = "SearchItem.Command"
args192(17).Value = 3
args192(18).Name = "Quiet"
args192(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args192())

rem ----------------------------------------------------------------------
dim args193(18) as new com.sun.star.beans.PropertyValue
args193(0).Name = "SearchItem.StyleFamily"
args193(0).Value = 2
args193(1).Name = "SearchItem.CellType"
args193(1).Value = 0
args193(2).Name = "SearchItem.RowDirection"
args193(2).Value = true
args193(3).Name = "SearchItem.AllTables"
args193(3).Value = false
args193(4).Name = "SearchItem.Backward"
args193(4).Value = false
args193(5).Name = "SearchItem.Pattern"
args193(5).Value = false
args193(6).Name = "SearchItem.Content"
args193(6).Value = false
args193(7).Name = "SearchItem.AsianOptions"
args193(7).Value = false
args193(8).Name = "SearchItem.AlgorithmType"
args193(8).Value = 0
args193(9).Name = "SearchItem.SearchFlags"
args193(9).Value = 65536
args193(10).Name = "SearchItem.SearchString"
args193(10).Value = "anoncis"
args193(11).Name = "SearchItem.ReplaceString"
args193(11).Value = "anuncis"
args193(12).Name = "SearchItem.Locale"
args193(12).Value = 255
args193(13).Name = "SearchItem.ChangedChars"
args193(13).Value = 2
args193(14).Name = "SearchItem.DeletedChars"
args193(14).Value = 2
args193(15).Name = "SearchItem.InsertedChars"
args193(15).Value = 2
args193(16).Name = "SearchItem.TransliterateFlags"
args193(16).Value = 1280
args193(17).Name = "SearchItem.Command"
args193(17).Value = 3
args193(18).Name = "Quiet"
args193(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args193())

rem ----------------------------------------------------------------------
dim args194(18) as new com.sun.star.beans.PropertyValue
args194(0).Name = "SearchItem.StyleFamily"
args194(0).Value = 2
args194(1).Name = "SearchItem.CellType"
args194(1).Value = 0
args194(2).Name = "SearchItem.RowDirection"
args194(2).Value = true
args194(3).Name = "SearchItem.AllTables"
args194(3).Value = false
args194(4).Name = "SearchItem.Backward"
args194(4).Value = false
args194(5).Name = "SearchItem.Pattern"
args194(5).Value = false
args194(6).Name = "SearchItem.Content"
args194(6).Value = false
args194(7).Name = "SearchItem.AsianOptions"
args194(7).Value = false
args194(8).Name = "SearchItem.AlgorithmType"
args194(8).Value = 0
args194(9).Name = "SearchItem.SearchFlags"
args194(9).Value = 65536
args194(10).Name = "SearchItem.SearchString"
args194(10).Value = "kombusteblajo"
args194(11).Name = "SearchItem.ReplaceString"
args194(11).Value = "fuelo"
args194(12).Name = "SearchItem.Locale"
args194(12).Value = 255
args194(13).Name = "SearchItem.ChangedChars"
args194(13).Value = 2
args194(14).Name = "SearchItem.DeletedChars"
args194(14).Value = 2
args194(15).Name = "SearchItem.InsertedChars"
args194(15).Value = 2
args194(16).Name = "SearchItem.TransliterateFlags"
args194(16).Value = 1280
args194(17).Name = "SearchItem.Command"
args194(17).Value = 3
args194(18).Name = "Quiet"
args194(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args194())

rem ----------------------------------------------------------------------
dim args195(18) as new com.sun.star.beans.PropertyValue
args195(0).Name = "SearchItem.StyleFamily"
args195(0).Value = 2
args195(1).Name = "SearchItem.CellType"
args195(1).Value = 0
args195(2).Name = "SearchItem.RowDirection"
args195(2).Value = true
args195(3).Name = "SearchItem.AllTables"
args195(3).Value = false
args195(4).Name = "SearchItem.Backward"
args195(4).Value = false
args195(5).Name = "SearchItem.Pattern"
args195(5).Value = false
args195(6).Name = "SearchItem.Content"
args195(6).Value = false
args195(7).Name = "SearchItem.AsianOptions"
args195(7).Value = false
args195(8).Name = "SearchItem.AlgorithmType"
args195(8).Value = 0
args195(9).Name = "SearchItem.SearchFlags"
args195(9).Value = 65536
args195(10).Name = "SearchItem.SearchString"
args195(10).Value = "Timor-Leste"
args195(11).Name = "SearchItem.ReplaceString"
args195(11).Value = "Estal Timor"
args195(12).Name = "SearchItem.Locale"
args195(12).Value = 255
args195(13).Name = "SearchItem.ChangedChars"
args195(13).Value = 2
args195(14).Name = "SearchItem.DeletedChars"
args195(14).Value = 2
args195(15).Name = "SearchItem.InsertedChars"
args195(15).Value = 2
args195(16).Name = "SearchItem.TransliterateFlags"
args195(16).Value = 1280
args195(17).Name = "SearchItem.Command"
args195(17).Value = 3
args195(18).Name = "Quiet"
args195(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args195())

rem ----------------------------------------------------------------------
dim args196(18) as new com.sun.star.beans.PropertyValue
args196(0).Name = "SearchItem.StyleFamily"
args196(0).Value = 2
args196(1).Name = "SearchItem.CellType"
args196(1).Value = 0
args196(2).Name = "SearchItem.RowDirection"
args196(2).Value = true
args196(3).Name = "SearchItem.AllTables"
args196(3).Value = false
args196(4).Name = "SearchItem.Backward"
args196(4).Value = false
args196(5).Name = "SearchItem.Pattern"
args196(5).Value = false
args196(6).Name = "SearchItem.Content"
args196(6).Value = false
args196(7).Name = "SearchItem.AsianOptions"
args196(7).Value = false
args196(8).Name = "SearchItem.AlgorithmType"
args196(8).Value = 0
args196(9).Name = "SearchItem.SearchFlags"
args196(9).Value = 65536
args196(10).Name = "SearchItem.SearchString"
args196(10).Value = "en homajo a"
args196(11).Name = "SearchItem.ReplaceString"
args196(11).Value = "memoriganta"
args196(12).Name = "SearchItem.Locale"
args196(12).Value = 255
args196(13).Name = "SearchItem.ChangedChars"
args196(13).Value = 2
args196(14).Name = "SearchItem.DeletedChars"
args196(14).Value = 2
args196(15).Name = "SearchItem.InsertedChars"
args196(15).Value = 2
args196(16).Name = "SearchItem.TransliterateFlags"
args196(16).Value = 1280
args196(17).Name = "SearchItem.Command"
args196(17).Value = 3
args196(18).Name = "Quiet"
args196(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args196())

rem ----------------------------------------------------------------------
dim args197(18) as new com.sun.star.beans.PropertyValue
args197(0).Name = "SearchItem.StyleFamily"
args197(0).Value = 2
args197(1).Name = "SearchItem.CellType"
args197(1).Value = 0
args197(2).Name = "SearchItem.RowDirection"
args197(2).Value = true
args197(3).Name = "SearchItem.AllTables"
args197(3).Value = false
args197(4).Name = "SearchItem.Backward"
args197(4).Value = false
args197(5).Name = "SearchItem.Pattern"
args197(5).Value = false
args197(6).Name = "SearchItem.Content"
args197(6).Value = false
args197(7).Name = "SearchItem.AsianOptions"
args197(7).Value = false
args197(8).Name = "SearchItem.AlgorithmType"
args197(8).Value = 0
args197(9).Name = "SearchItem.SearchFlags"
args197(9).Value = 65536
args197(10).Name = "SearchItem.SearchString"
args197(10).Value = "Richter-skalo"
args197(11).Name = "SearchItem.ReplaceString"
args197(11).Value = "skalo di Richter"
args197(12).Name = "SearchItem.Locale"
args197(12).Value = 255
args197(13).Name = "SearchItem.ChangedChars"
args197(13).Value = 2
args197(14).Name = "SearchItem.DeletedChars"
args197(14).Value = 2
args197(15).Name = "SearchItem.InsertedChars"
args197(15).Value = 2
args197(16).Name = "SearchItem.TransliterateFlags"
args197(16).Value = 1280
args197(17).Name = "SearchItem.Command"
args197(17).Value = 3
args197(18).Name = "Quiet"
args197(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args197())

rem ----------------------------------------------------------------------
dim args198(18) as new com.sun.star.beans.PropertyValue
args198(0).Name = "SearchItem.StyleFamily"
args198(0).Value = 2
args198(1).Name = "SearchItem.CellType"
args198(1).Value = 0
args198(2).Name = "SearchItem.RowDirection"
args198(2).Value = true
args198(3).Name = "SearchItem.AllTables"
args198(3).Value = false
args198(4).Name = "SearchItem.Backward"
args198(4).Value = false
args198(5).Name = "SearchItem.Pattern"
args198(5).Value = false
args198(6).Name = "SearchItem.Content"
args198(6).Value = false
args198(7).Name = "SearchItem.AsianOptions"
args198(7).Value = false
args198(8).Name = "SearchItem.AlgorithmType"
args198(8).Value = 0
args198(9).Name = "SearchItem.SearchFlags"
args198(9).Value = 65536
args198(10).Name = "SearchItem.SearchString"
args198(10).Value = "disputo"
args198(11).Name = "SearchItem.ReplaceString"
args198(11).Value = "konkurso"
args198(12).Name = "SearchItem.Locale"
args198(12).Value = 255
args198(13).Name = "SearchItem.ChangedChars"
args198(13).Value = 2
args198(14).Name = "SearchItem.DeletedChars"
args198(14).Value = 2
args198(15).Name = "SearchItem.InsertedChars"
args198(15).Value = 2
args198(16).Name = "SearchItem.TransliterateFlags"
args198(16).Value = 1280
args198(17).Name = "SearchItem.Command"
args198(17).Value = 3
args198(18).Name = "Quiet"
args198(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args198())

rem ----------------------------------------------------------------------
dim args199(18) as new com.sun.star.beans.PropertyValue
args199(0).Name = "SearchItem.StyleFamily"
args199(0).Value = 2
args199(1).Name = "SearchItem.CellType"
args199(1).Value = 0
args199(2).Name = "SearchItem.RowDirection"
args199(2).Value = true
args199(3).Name = "SearchItem.AllTables"
args199(3).Value = false
args199(4).Name = "SearchItem.Backward"
args199(4).Value = false
args199(5).Name = "SearchItem.Pattern"
args199(5).Value = false
args199(6).Name = "SearchItem.Content"
args199(6).Value = false
args199(7).Name = "SearchItem.AsianOptions"
args199(7).Value = false
args199(8).Name = "SearchItem.AlgorithmType"
args199(8).Value = 0
args199(9).Name = "SearchItem.SearchFlags"
args199(9).Value = 65536
args199(10).Name = "SearchItem.SearchString"
args199(10).Value = "ordinatr"
args199(11).Name = "SearchItem.ReplaceString"
args199(11).Value = "komputer"
args199(12).Name = "SearchItem.Locale"
args199(12).Value = 255
args199(13).Name = "SearchItem.ChangedChars"
args199(13).Value = 2
args199(14).Name = "SearchItem.DeletedChars"
args199(14).Value = 2
args199(15).Name = "SearchItem.InsertedChars"
args199(15).Value = 2
args199(16).Name = "SearchItem.TransliterateFlags"
args199(16).Value = 1280
args199(17).Name = "SearchItem.Command"
args199(17).Value = 3
args199(18).Name = "Quiet"
args199(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args199())

rem ----------------------------------------------------------------------
dim args200(18) as new com.sun.star.beans.PropertyValue
args200(0).Name = "SearchItem.StyleFamily"
args200(0).Value = 2
args200(1).Name = "SearchItem.CellType"
args200(1).Value = 0
args200(2).Name = "SearchItem.RowDirection"
args200(2).Value = true
args200(3).Name = "SearchItem.AllTables"
args200(3).Value = false
args200(4).Name = "SearchItem.Backward"
args200(4).Value = false
args200(5).Name = "SearchItem.Pattern"
args200(5).Value = false
args200(6).Name = "SearchItem.Content"
args200(6).Value = false
args200(7).Name = "SearchItem.AsianOptions"
args200(7).Value = false
args200(8).Name = "SearchItem.AlgorithmType"
args200(8).Value = 0
args200(9).Name = "SearchItem.SearchFlags"
args200(9).Value = 65536
args200(10).Name = "SearchItem.SearchString"
args200(10).Value = "sufris intensa"
args200(11).Name = "SearchItem.ReplaceString"
args200(11).Value = "subisis intensa"
args200(12).Name = "SearchItem.Locale"
args200(12).Value = 255
args200(13).Name = "SearchItem.ChangedChars"
args200(13).Value = 2
args200(14).Name = "SearchItem.DeletedChars"
args200(14).Value = 2
args200(15).Name = "SearchItem.InsertedChars"
args200(15).Value = 2
args200(16).Name = "SearchItem.TransliterateFlags"
args200(16).Value = 1280
args200(17).Name = "SearchItem.Command"
args200(17).Value = 3
args200(18).Name = "Quiet"
args200(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args200())


rem ----------------------------------------------------------------------
dim args201(18) as new com.sun.star.beans.PropertyValue
args201(0).Name = "SearchItem.StyleFamily"
args201(0).Value = 2
args201(1).Name = "SearchItem.CellType"
args201(1).Value = 0
args201(2).Name = "SearchItem.RowDirection"
args201(2).Value = true
args201(3).Name = "SearchItem.AllTables"
args201(3).Value = false
args201(4).Name = "SearchItem.Backward"
args201(4).Value = false
args201(5).Name = "SearchItem.Pattern"
args201(5).Value = false
args201(6).Name = "SearchItem.Content"
args201(6).Value = false
args201(7).Name = "SearchItem.AsianOptions"
args201(7).Value = false
args201(8).Name = "SearchItem.AlgorithmType"
args201(8).Value = 0
args201(9).Name = "SearchItem.SearchFlags"
args201(9).Value = 65536
args201(10).Name = "SearchItem.SearchString"
args201(10).Value = " fluvio]]"
args201(11).Name = "SearchItem.ReplaceString"
args201(11).Value = ""
args201(12).Name = "SearchItem.Locale"
args201(12).Value = 255
args201(13).Name = "SearchItem.ChangedChars"
args201(13).Value = 2
args201(14).Name = "SearchItem.DeletedChars"
args201(14).Value = 2
args201(15).Name = "SearchItem.InsertedChars"
args201(15).Value = 2
args201(16).Name = "SearchItem.TransliterateFlags"
args201(16).Value = 1280
args201(17).Name = "SearchItem.Command"
args201(17).Value = 1
args201(18).Name = "Quiet"
args201(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args201())

rem ----------------------------------------------------------------------
dim args202(18) as new com.sun.star.beans.PropertyValue
args202(0).Name = "SearchItem.StyleFamily"
args202(0).Value = 2
args202(1).Name = "SearchItem.CellType"
args202(1).Value = 0
args202(2).Name = "SearchItem.RowDirection"
args202(2).Value = true
args202(3).Name = "SearchItem.AllTables"
args202(3).Value = false
args202(4).Name = "SearchItem.Backward"
args202(4).Value = false
args202(5).Name = "SearchItem.Pattern"
args202(5).Value = false
args202(6).Name = "SearchItem.Content"
args202(6).Value = false
args202(7).Name = "SearchItem.AsianOptions"
args202(7).Value = false
args202(8).Name = "SearchItem.AlgorithmType"
args202(8).Value = 0
args202(9).Name = "SearchItem.SearchFlags"
args202(9).Value = 65536
args202(10).Name = "SearchItem.SearchString"
args202(10).Value = "UsanaStati|"
args202(11).Name = "SearchItem.ReplaceString"
args202(11).Value = "Provinci stati kantoni"
args202(12).Name = "SearchItem.Locale"
args202(12).Value = 255
args202(13).Name = "SearchItem.ChangedChars"
args202(13).Value = 2
args202(14).Name = "SearchItem.DeletedChars"
args202(14).Value = 2
args202(15).Name = "SearchItem.InsertedChars"
args202(15).Value = 2
args202(16).Name = "SearchItem.TransliterateFlags"
args202(16).Value = 1280
args202(17).Name = "SearchItem.Command"
args202(17).Value = 3
args202(18).Name = "Quiet"
args202(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args202())

rem ----------------------------------------------------------------------
dim args203(18) as new com.sun.star.beans.PropertyValue
args203(0).Name = "SearchItem.StyleFamily"
args203(0).Value = 2
args203(1).Name = "SearchItem.CellType"
args203(1).Value = 0
args203(2).Name = "SearchItem.RowDirection"
args203(2).Value = true
args203(3).Name = "SearchItem.AllTables"
args203(3).Value = false
args203(4).Name = "SearchItem.Backward"
args203(4).Value = false
args203(5).Name = "SearchItem.Pattern"
args203(5).Value = false
args203(6).Name = "SearchItem.Content"
args203(6).Value = false
args203(7).Name = "SearchItem.AsianOptions"
args203(7).Value = false
args203(8).Name = "SearchItem.AlgorithmType"
args203(8).Value = 0
args203(9).Name = "SearchItem.SearchFlags"
args203(9).Value = 65536
args203(10).Name = "SearchItem.SearchString"
args203(10).Value = "ekonomikala aktivesi"
args203(11).Name = "SearchItem.ReplaceString"
args203(11).Value = "ekonomial agadi"
args203(12).Name = "SearchItem.Locale"
args203(12).Value = 255
args203(13).Name = "SearchItem.ChangedChars"
args203(13).Value = 2
args203(14).Name = "SearchItem.DeletedChars"
args203(14).Value = 2
args203(15).Name = "SearchItem.InsertedChars"
args203(15).Value = 2
args203(16).Name = "SearchItem.TransliterateFlags"
args203(16).Value = 1280
args203(17).Name = "SearchItem.Command"
args203(17).Value = 3
args203(18).Name = "Quiet"
args203(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args203())

rem ----------------------------------------------------------------------
dim args204(18) as new com.sun.star.beans.PropertyValue
args204(0).Name = "SearchItem.StyleFamily"
args204(0).Value = 2
args204(1).Name = "SearchItem.CellType"
args204(1).Value = 0
args204(2).Name = "SearchItem.RowDirection"
args204(2).Value = true
args204(3).Name = "SearchItem.AllTables"
args204(3).Value = false
args204(4).Name = "SearchItem.Backward"
args204(4).Value = false
args204(5).Name = "SearchItem.Pattern"
args204(5).Value = false
args204(6).Name = "SearchItem.Content"
args204(6).Value = false
args204(7).Name = "SearchItem.AsianOptions"
args204(7).Value = false
args204(8).Name = "SearchItem.AlgorithmType"
args204(8).Value = 0
args204(9).Name = "SearchItem.SearchFlags"
args204(9).Value = 65536
args204(10).Name = "SearchItem.SearchString"
args204(10).Value = "esas konstitucita"
args204(11).Name = "SearchItem.ReplaceString"
args204(11).Value = "formacesas"
args204(12).Name = "SearchItem.Locale"
args204(12).Value = 255
args204(13).Name = "SearchItem.ChangedChars"
args204(13).Value = 2
args204(14).Name = "SearchItem.DeletedChars"
args204(14).Value = 2
args204(15).Name = "SearchItem.InsertedChars"
args204(15).Value = 2
args204(16).Name = "SearchItem.TransliterateFlags"
args204(16).Value = 1280
args204(17).Name = "SearchItem.Command"
args204(17).Value = 3
args204(18).Name = "Quiet"
args204(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args204())

rem ----------------------------------------------------------------------
dim args205(18) as new com.sun.star.beans.PropertyValue
args205(0).Name = "SearchItem.StyleFamily"
args205(0).Value = 2
args205(1).Name = "SearchItem.CellType"
args205(1).Value = 0
args205(2).Name = "SearchItem.RowDirection"
args205(2).Value = true
args205(3).Name = "SearchItem.AllTables"
args205(3).Value = false
args205(4).Name = "SearchItem.Backward"
args205(4).Value = false
args205(5).Name = "SearchItem.Pattern"
args205(5).Value = false
args205(6).Name = "SearchItem.Content"
args205(6).Value = false
args205(7).Name = "SearchItem.AsianOptions"
args205(7).Value = false
args205(8).Name = "SearchItem.AlgorithmType"
args205(8).Value = 0
args205(9).Name = "SearchItem.SearchFlags"
args205(9).Value = 65536
args205(10).Name = "SearchItem.SearchString"
args205(10).Value = "mesur"
args205(11).Name = "SearchItem.ReplaceString"
args205(11).Value = "mezur"
args205(12).Name = "SearchItem.Locale"
args205(12).Value = 255
args205(13).Name = "SearchItem.ChangedChars"
args205(13).Value = 2
args205(14).Name = "SearchItem.DeletedChars"
args205(14).Value = 2
args205(15).Name = "SearchItem.InsertedChars"
args205(15).Value = 2
args205(16).Name = "SearchItem.TransliterateFlags"
args205(16).Value = 1280
args205(17).Name = "SearchItem.Command"
args205(17).Value = 3
args205(18).Name = "Quiet"
args205(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args205())

rem ----------------------------------------------------------------------
dim args206(18) as new com.sun.star.beans.PropertyValue
args206(0).Name = "SearchItem.StyleFamily"
args206(0).Value = 2
args206(1).Name = "SearchItem.CellType"
args206(1).Value = 0
args206(2).Name = "SearchItem.RowDirection"
args206(2).Value = true
args206(3).Name = "SearchItem.AllTables"
args206(3).Value = false
args206(4).Name = "SearchItem.Backward"
args206(4).Value = false
args206(5).Name = "SearchItem.Pattern"
args206(5).Value = false
args206(6).Name = "SearchItem.Content"
args206(6).Value = false
args206(7).Name = "SearchItem.AsianOptions"
args206(7).Value = false
args206(8).Name = "SearchItem.AlgorithmType"
args206(8).Value = 0
args206(9).Name = "SearchItem.SearchFlags"
args206(9).Value = 65536
args206(10).Name = "SearchItem.SearchString"
args206(10).Value = "revenuo por familio esas"
args206(11).Name = "SearchItem.ReplaceString"
args206(11).Value = "revenuo per familio esis"
args206(12).Name = "SearchItem.Locale"
args206(12).Value = 255
args206(13).Name = "SearchItem.ChangedChars"
args206(13).Value = 2
args206(14).Name = "SearchItem.DeletedChars"
args206(14).Value = 2
args206(15).Name = "SearchItem.InsertedChars"
args206(15).Value = 2
args206(16).Name = "SearchItem.TransliterateFlags"
args206(16).Value = 1280
args206(17).Name = "SearchItem.Command"
args206(17).Value = 3
args206(18).Name = "Quiet"
args206(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args206())

rem ----------------------------------------------------------------------
dim args207(18) as new com.sun.star.beans.PropertyValue
args207(0).Name = "SearchItem.StyleFamily"
args207(0).Value = 2
args207(1).Name = "SearchItem.CellType"
args207(1).Value = 0
args207(2).Name = "SearchItem.RowDirection"
args207(2).Value = true
args207(3).Name = "SearchItem.AllTables"
args207(3).Value = false
args207(4).Name = "SearchItem.Backward"
args207(4).Value = false
args207(5).Name = "SearchItem.Pattern"
args207(5).Value = false
args207(6).Name = "SearchItem.Content"
args207(6).Value = false
args207(7).Name = "SearchItem.AsianOptions"
args207(7).Value = false
args207(8).Name = "SearchItem.AlgorithmType"
args207(8).Value = 0
args207(9).Name = "SearchItem.SearchFlags"
args207(9).Value = 65536
args207(10).Name = "SearchItem.SearchString"
args207(10).Value = " tota la "
args207(11).Name = "SearchItem.ReplaceString"
args207(11).Value = " omna "
args207(12).Name = "SearchItem.Locale"
args207(12).Value = 255
args207(13).Name = "SearchItem.ChangedChars"
args207(13).Value = 2
args207(14).Name = "SearchItem.DeletedChars"
args207(14).Value = 2
args207(15).Name = "SearchItem.InsertedChars"
args207(15).Value = 2
args207(16).Name = "SearchItem.TransliterateFlags"
args207(16).Value = 1280
args207(17).Name = "SearchItem.Command"
args207(17).Value = 3
args207(18).Name = "Quiet"
args207(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args207())

rem ----------------------------------------------------------------------
dim args208(18) as new com.sun.star.beans.PropertyValue
args208(0).Name = "SearchItem.StyleFamily"
args208(0).Value = 2
args208(1).Name = "SearchItem.CellType"
args208(1).Value = 0
args208(2).Name = "SearchItem.RowDirection"
args208(2).Value = true
args208(3).Name = "SearchItem.AllTables"
args208(3).Value = false
args208(4).Name = "SearchItem.Backward"
args208(4).Value = false
args208(5).Name = "SearchItem.Pattern"
args208(5).Value = false
args208(6).Name = "SearchItem.Content"
args208(6).Value = false
args208(7).Name = "SearchItem.AsianOptions"
args208(7).Value = false
args208(8).Name = "SearchItem.AlgorithmType"
args208(8).Value = 0
args208(9).Name = "SearchItem.SearchFlags"
args208(9).Value = 65536
args208(10).Name = "SearchItem.SearchString"
args208(10).Value = "Indijena amerikana"
args208(11).Name = "SearchItem.ReplaceString"
args208(11).Value = "indijeni Amerikana"
args208(12).Name = "SearchItem.Locale"
args208(12).Value = 255
args208(13).Name = "SearchItem.ChangedChars"
args208(13).Value = 2
args208(14).Name = "SearchItem.DeletedChars"
args208(14).Value = 2
args208(15).Name = "SearchItem.InsertedChars"
args208(15).Value = 2
args208(16).Name = "SearchItem.TransliterateFlags"
args208(16).Value = 1280
args208(17).Name = "SearchItem.Command"
args208(17).Value = 3
args208(18).Name = "Quiet"
args208(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args208())

rem ----------------------------------------------------------------------
dim args209(18) as new com.sun.star.beans.PropertyValue
args209(0).Name = "SearchItem.StyleFamily"
args209(0).Value = 2
args209(1).Name = "SearchItem.CellType"
args209(1).Value = 0
args209(2).Name = "SearchItem.RowDirection"
args209(2).Value = true
args209(3).Name = "SearchItem.AllTables"
args209(3).Value = false
args209(4).Name = "SearchItem.Backward"
args209(4).Value = false
args209(5).Name = "SearchItem.Pattern"
args209(5).Value = false
args209(6).Name = "SearchItem.Content"
args209(6).Value = false
args209(7).Name = "SearchItem.AsianOptions"
args209(7).Value = false
args209(8).Name = "SearchItem.AlgorithmType"
args209(8).Value = 0
args209(9).Name = "SearchItem.SearchFlags"
args209(9).Value = 65536
args209(10).Name = "SearchItem.SearchString"
args209(10).Value = " plu di "
args209(11).Name = "SearchItem.ReplaceString"
args209(11).Value = " plu kam "
args209(12).Name = "SearchItem.Locale"
args209(12).Value = 255
args209(13).Name = "SearchItem.ChangedChars"
args209(13).Value = 2
args209(14).Name = "SearchItem.DeletedChars"
args209(14).Value = 2
args209(15).Name = "SearchItem.InsertedChars"
args209(15).Value = 2
args209(16).Name = "SearchItem.TransliterateFlags"
args209(16).Value = 1280
args209(17).Name = "SearchItem.Command"
args209(17).Value = 3
args209(18).Name = "Quiet"
args209(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args209())

rem ----------------------------------------------------------------------
dim args210(18) as new com.sun.star.beans.PropertyValue
args210(0).Name = "SearchItem.StyleFamily"
args210(0).Value = 2
args210(1).Name = "SearchItem.CellType"
args210(1).Value = 0
args210(2).Name = "SearchItem.RowDirection"
args210(2).Value = true
args210(3).Name = "SearchItem.AllTables"
args210(3).Value = false
args210(4).Name = "SearchItem.Backward"
args210(4).Value = false
args210(5).Name = "SearchItem.Pattern"
args210(5).Value = false
args210(6).Name = "SearchItem.Content"
args210(6).Value = false
args210(7).Name = "SearchItem.AsianOptions"
args210(7).Value = false
args210(8).Name = "SearchItem.AlgorithmType"
args210(8).Value = 0
args210(9).Name = "SearchItem.SearchFlags"
args210(9).Value = 65536
args210(10).Name = "SearchItem.SearchString"
args210(10).Value = "Britan "
args210(11).Name = "SearchItem.ReplaceString"
args210(11).Value = "Britanian "
args210(12).Name = "SearchItem.Locale"
args210(12).Value = 255
args210(13).Name = "SearchItem.ChangedChars"
args210(13).Value = 2
args210(14).Name = "SearchItem.DeletedChars"
args210(14).Value = 2
args210(15).Name = "SearchItem.InsertedChars"
args210(15).Value = 2
args210(16).Name = "SearchItem.TransliterateFlags"
args210(16).Value = 1280
args210(17).Name = "SearchItem.Command"
args210(17).Value = 3
args210(18).Name = "Quiet"
args210(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args210())

rem ----------------------------------------------------------------------
dim args211(18) as new com.sun.star.beans.PropertyValue
args211(0).Name = "SearchItem.StyleFamily"
args211(0).Value = 2
args211(1).Name = "SearchItem.CellType"
args211(1).Value = 0
args211(2).Name = "SearchItem.RowDirection"
args211(2).Value = true
args211(3).Name = "SearchItem.AllTables"
args211(3).Value = false
args211(4).Name = "SearchItem.Backward"
args211(4).Value = false
args211(5).Name = "SearchItem.Pattern"
args211(5).Value = false
args211(6).Name = "SearchItem.Content"
args211(6).Value = false
args211(7).Name = "SearchItem.AsianOptions"
args211(7).Value = false
args211(8).Name = "SearchItem.AlgorithmType"
args211(8).Value = 0
args211(9).Name = "SearchItem.SearchFlags"
args211(9).Value = 65536
args211(10).Name = "SearchItem.SearchString"
args211(10).Value = "ed ek tota la habitantaro"
args211(11).Name = "SearchItem.ReplaceString"
args211(11).Value = "ed ek omna habitantaro"
args211(12).Name = "SearchItem.Locale"
args211(12).Value = 255
args211(13).Name = "SearchItem.ChangedChars"
args211(13).Value = 2
args211(14).Name = "SearchItem.DeletedChars"
args211(14).Value = 2
args211(15).Name = "SearchItem.InsertedChars"
args211(15).Value = 2
args211(16).Name = "SearchItem.TransliterateFlags"
args211(16).Value = 1280
args211(17).Name = "SearchItem.Command"
args211(17).Value = 3
args211(18).Name = "Quiet"
args211(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args211())

rem ----------------------------------------------------------------------
dim args212(18) as new com.sun.star.beans.PropertyValue
args212(0).Name = "SearchItem.StyleFamily"
args212(0).Value = 2
args212(1).Name = "SearchItem.CellType"
args212(1).Value = 0
args212(2).Name = "SearchItem.RowDirection"
args212(2).Value = true
args212(3).Name = "SearchItem.AllTables"
args212(3).Value = false
args212(4).Name = "SearchItem.Backward"
args212(4).Value = false
args212(5).Name = "SearchItem.Pattern"
args212(5).Value = false
args212(6).Name = "SearchItem.Content"
args212(6).Value = false
args212(7).Name = "SearchItem.AsianOptions"
args212(7).Value = false
args212(8).Name = "SearchItem.AlgorithmType"
args212(8).Value = 0
args212(9).Name = "SearchItem.SearchFlags"
args212(9).Value = 65536
args212(10).Name = "SearchItem.SearchString"
args212(10).Value = "Ek la tota habitantaro "
args212(11).Name = "SearchItem.ReplaceString"
args212(11).Value = "Ek omna habitantaro "
args212(12).Name = "SearchItem.Locale"
args212(12).Value = 255
args212(13).Name = "SearchItem.ChangedChars"
args212(13).Value = 2
args212(14).Name = "SearchItem.DeletedChars"
args212(14).Value = 2
args212(15).Name = "SearchItem.InsertedChars"
args212(15).Value = 2
args212(16).Name = "SearchItem.TransliterateFlags"
args212(16).Value = 1024
args212(17).Name = "SearchItem.Command"
args212(17).Value = 3
args212(18).Name = "Quiet"
args212(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args212())

rem ----------------------------------------------------------------------
dim args213(18) as new com.sun.star.beans.PropertyValue
args213(0).Name = "SearchItem.StyleFamily"
args213(0).Value = 2
args213(1).Name = "SearchItem.CellType"
args213(1).Value = 0
args213(2).Name = "SearchItem.RowDirection"
args213(2).Value = true
args213(3).Name = "SearchItem.AllTables"
args213(3).Value = false
args213(4).Name = "SearchItem.Backward"
args213(4).Value = false
args213(5).Name = "SearchItem.Pattern"
args213(5).Value = false
args213(6).Name = "SearchItem.Content"
args213(6).Value = false
args213(7).Name = "SearchItem.AsianOptions"
args213(7).Value = false
args213(8).Name = "SearchItem.AlgorithmType"
args213(8).Value = 0
args213(9).Name = "SearchItem.SearchFlags"
args213(9).Value = 65536
args213(10).Name = "SearchItem.SearchString"
args213(10).Value = "Image:"
args213(11).Name = "SearchItem.ReplaceString"
args213(11).Value = "Arkivo:"
args213(12).Name = "SearchItem.Locale"
args213(12).Value = 255
args213(13).Name = "SearchItem.ChangedChars"
args213(13).Value = 2
args213(14).Name = "SearchItem.DeletedChars"
args213(14).Value = 2
args213(15).Name = "SearchItem.InsertedChars"
args213(15).Value = 2
args213(16).Name = "SearchItem.TransliterateFlags"
args213(16).Value = 1280
args213(17).Name = "SearchItem.Command"
args213(17).Value = 3
args213(18).Name = "Quiet"
args213(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args213())

rem ----------------------------------------------------------------------
dim args214(18) as new com.sun.star.beans.PropertyValue
args214(0).Name = "SearchItem.StyleFamily"
args214(0).Value = 2
args214(1).Name = "SearchItem.CellType"
args214(1).Value = 0
args214(2).Name = "SearchItem.RowDirection"
args214(2).Value = true
args214(3).Name = "SearchItem.AllTables"
args214(3).Value = false
args214(4).Name = "SearchItem.Backward"
args214(4).Value = false
args214(5).Name = "SearchItem.Pattern"
args214(5).Value = false
args214(6).Name = "SearchItem.Content"
args214(6).Value = false
args214(7).Name = "SearchItem.AsianOptions"
args214(7).Value = false
args214(8).Name = "SearchItem.AlgorithmType"
args214(8).Value = 0
args214(9).Name = "SearchItem.SearchFlags"
args214(9).Value = 65536
args214(10).Name = "SearchItem.SearchString"
args214(10).Value = " kun bomb"
args214(11).Name = "SearchItem.ReplaceString"
args214(11).Value = " per bomb"
args214(12).Name = "SearchItem.Locale"
args214(12).Value = 255
args214(13).Name = "SearchItem.ChangedChars"
args214(13).Value = 2
args214(14).Name = "SearchItem.DeletedChars"
args214(14).Value = 2
args214(15).Name = "SearchItem.InsertedChars"
args214(15).Value = 2
args214(16).Name = "SearchItem.TransliterateFlags"
args214(16).Value = 1280
args214(17).Name = "SearchItem.Command"
args214(17).Value = 3
args214(18).Name = "Quiet"
args214(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args214())

rem ----------------------------------------------------------------------
dim args215(18) as new com.sun.star.beans.PropertyValue
args215(0).Name = "SearchItem.StyleFamily"
args215(0).Value = 2
args215(1).Name = "SearchItem.CellType"
args215(1).Value = 0
args215(2).Name = "SearchItem.RowDirection"
args215(2).Value = true
args215(3).Name = "SearchItem.AllTables"
args215(3).Value = false
args215(4).Name = "SearchItem.Backward"
args215(4).Value = false
args215(5).Name = "SearchItem.Pattern"
args215(5).Value = false
args215(6).Name = "SearchItem.Content"
args215(6).Value = false
args215(7).Name = "SearchItem.AsianOptions"
args215(7).Value = false
args215(8).Name = "SearchItem.AlgorithmType"
args215(8).Value = 0
args215(9).Name = "SearchItem.SearchFlags"
args215(9).Value = 65536
args215(10).Name = "SearchItem.SearchString"
args215(10).Value = " mezala "
args215(11).Name = "SearchItem.ReplaceString"
args215(11).Value = " mezvalora "
args215(12).Name = "SearchItem.Locale"
args215(12).Value = 255
args215(13).Name = "SearchItem.ChangedChars"
args215(13).Value = 2
args215(14).Name = "SearchItem.DeletedChars"
args215(14).Value = 2
args215(15).Name = "SearchItem.InsertedChars"
args215(15).Value = 2
args215(16).Name = "SearchItem.TransliterateFlags"
args215(16).Value = 1280
args215(17).Name = "SearchItem.Command"
args215(17).Value = 3
args215(18).Name = "Quiet"
args215(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args215())

rem ----------------------------------------------------------------------
dim args216(18) as new com.sun.star.beans.PropertyValue
args216(0).Name = "SearchItem.StyleFamily"
args216(0).Value = 2
args216(1).Name = "SearchItem.CellType"
args216(1).Value = 0
args216(2).Name = "SearchItem.RowDirection"
args216(2).Value = true
args216(3).Name = "SearchItem.AllTables"
args216(3).Value = false
args216(4).Name = "SearchItem.Backward"
args216(4).Value = false
args216(5).Name = "SearchItem.Pattern"
args216(5).Value = false
args216(6).Name = "SearchItem.Content"
args216(6).Value = false
args216(7).Name = "SearchItem.AsianOptions"
args216(7).Value = false
args216(8).Name = "SearchItem.AlgorithmType"
args216(8).Value = 0
args216(9).Name = "SearchItem.SearchFlags"
args216(9).Value = 65536
args216(10).Name = "SearchItem.SearchString"
args216(10).Value = "Pacifika Insulana"
args216(11).Name = "SearchItem.ReplaceString"
args216(11).Value = "Pacifik insulani"
args216(12).Name = "SearchItem.Locale"
args216(12).Value = 255
args216(13).Name = "SearchItem.ChangedChars"
args216(13).Value = 2
args216(14).Name = "SearchItem.DeletedChars"
args216(14).Value = 2
args216(15).Name = "SearchItem.InsertedChars"
args216(15).Value = 2
args216(16).Name = "SearchItem.TransliterateFlags"
args216(16).Value = 1280
args216(17).Name = "SearchItem.Command"
args216(17).Value = 3
args216(18).Name = "Quiet"
args216(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args216())

rem ----------------------------------------------------------------------
dim args217(18) as new com.sun.star.beans.PropertyValue

dim NOVEMB1, NOVEMB2 as string
For I = 1 to 30
NOVEMB1 = Ltrim(Str$(I))+" di novem"
NOVEMB2 = Ltrim(Str$(I))+"ma di novem"

args217(0).Name = "SearchItem.StyleFamily"
args217(0).Value = 2
args217(1).Name = "SearchItem.CellType"
args217(1).Value = 0
args217(2).Name = "SearchItem.RowDirection"
args217(2).Value = true
args217(3).Name = "SearchItem.AllTables"
args217(3).Value = false
args217(4).Name = "SearchItem.Backward"
args217(4).Value = false
args217(5).Name = "SearchItem.Pattern"
args217(5).Value = false
args217(6).Name = "SearchItem.Content"
args217(6).Value = false
args217(7).Name = "SearchItem.AsianOptions"
args217(7).Value = false
args217(8).Name = "SearchItem.AlgorithmType"
args217(8).Value = 0
args217(9).Name = "SearchItem.SearchFlags"
args217(9).Value = 65536
args217(10).Name = "SearchItem.SearchString"
args217(10).Value = NOVEMB1
args217(11).Name = "SearchItem.ReplaceString"
args217(11).Value = NOVEMB2
args217(12).Name = "SearchItem.Locale"
args217(12).Value = 255
args217(13).Name = "SearchItem.ChangedChars"
args217(13).Value = 2
args217(14).Name = "SearchItem.DeletedChars"
args217(14).Value = 2
args217(15).Name = "SearchItem.InsertedChars"
args217(15).Value = 2
args217(16).Name = "SearchItem.TransliterateFlags"
args217(16).Value = 1280
args217(17).Name = "SearchItem.Command"
args217(17).Value = 3
args217(18).Name = "Quiet"
args217(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args217())

Next I

rem ----------------------------------------------------------------------
dim args218(18) as new com.sun.star.beans.PropertyValue
args218(0).Name = "SearchItem.StyleFamily"
args218(0).Value = 2
args218(1).Name = "SearchItem.CellType"
args218(1).Value = 0
args218(2).Name = "SearchItem.RowDirection"
args218(2).Value = true
args218(3).Name = "SearchItem.AllTables"
args218(3).Value = false
args218(4).Name = "SearchItem.Backward"
args218(4).Value = false
args218(5).Name = "SearchItem.Pattern"
args218(5).Value = false
args218(6).Name = "SearchItem.Content"
args218(6).Value = false
args218(7).Name = "SearchItem.AsianOptions"
args218(7).Value = false
args218(8).Name = "SearchItem.AlgorithmType"
args218(8).Value = 0
args218(9).Name = "SearchItem.SearchFlags"
args218(9).Value = 65536
args218(10).Name = "SearchItem.SearchString"
args218(10).Value = "rejado di"
args218(11).Name = "SearchItem.ReplaceString"
args218(11).Value = "regno di"
args218(12).Name = "SearchItem.Locale"
args218(12).Value = 255
args218(13).Name = "SearchItem.ChangedChars"
args218(13).Value = 2
args218(14).Name = "SearchItem.DeletedChars"
args218(14).Value = 2
args218(15).Name = "SearchItem.InsertedChars"
args218(15).Value = 2
args218(16).Name = "SearchItem.TransliterateFlags"
args218(16).Value = 1280
args218(17).Name = "SearchItem.Command"
args218(17).Value = 3
args218(18).Name = "Quiet"
args218(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args218())

rem ----------------------------------------------------------------------
dim args219(18) as new com.sun.star.beans.PropertyValue
args219(0).Name = "SearchItem.StyleFamily"
args219(0).Value = 2
args219(1).Name = "SearchItem.CellType"
args219(1).Value = 0
args219(2).Name = "SearchItem.RowDirection"
args219(2).Value = true
args219(3).Name = "SearchItem.AllTables"
args219(3).Value = false
args219(4).Name = "SearchItem.Backward"
args219(4).Value = false
args219(5).Name = "SearchItem.Pattern"
args219(5).Value = false
args219(6).Name = "SearchItem.Content"
args219(6).Value = false
args219(7).Name = "SearchItem.AsianOptions"
args219(7).Value = false
args219(8).Name = "SearchItem.AlgorithmType"
args219(8).Value = 0
args219(9).Name = "SearchItem.SearchFlags"
args219(9).Value = 65536
args219(10).Name = "SearchItem.SearchString"
args219(10).Value = " greca "
args219(11).Name = "SearchItem.ReplaceString"
args219(11).Value = " Greka "
args219(12).Name = "SearchItem.Locale"
args219(12).Value = 255
args219(13).Name = "SearchItem.ChangedChars"
args219(13).Value = 2
args219(14).Name = "SearchItem.DeletedChars"
args219(14).Value = 2
args219(15).Name = "SearchItem.InsertedChars"
args219(15).Value = 2
args219(16).Name = "SearchItem.TransliterateFlags"
args219(16).Value = 1280
args219(17).Name = "SearchItem.Command"
args219(17).Value = 3
args219(18).Name = "Quiet"
args219(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args219())

rem ----------------------------------------------------------------------
dim args220(18) as new com.sun.star.beans.PropertyValue
args220(0).Name = "SearchItem.StyleFamily"
args220(0).Value = 2
args220(1).Name = "SearchItem.CellType"
args220(1).Value = 0
args220(2).Name = "SearchItem.RowDirection"
args220(2).Value = true
args220(3).Name = "SearchItem.AllTables"
args220(3).Value = false
args220(4).Name = "SearchItem.Backward"
args220(4).Value = false
args220(5).Name = "SearchItem.Pattern"
args220(5).Value = false
args220(6).Name = "SearchItem.Content"
args220(6).Value = false
args220(7).Name = "SearchItem.AsianOptions"
args220(7).Value = false
args220(8).Name = "SearchItem.AlgorithmType"
args220(8).Value = 0
args220(9).Name = "SearchItem.SearchFlags"
args220(9).Value = 65536
args220(10).Name = "SearchItem.SearchString"
args220(10).Value = "menac"
args220(11).Name = "SearchItem.ReplaceString"
args220(11).Value = "minac"
args220(12).Name = "SearchItem.Locale"
args220(12).Value = 255
args220(13).Name = "SearchItem.ChangedChars"
args220(13).Value = 2
args220(14).Name = "SearchItem.DeletedChars"
args220(14).Value = 2
args220(15).Name = "SearchItem.InsertedChars"
args220(15).Value = 2
args220(16).Name = "SearchItem.TransliterateFlags"
args220(16).Value = 1280
args220(17).Name = "SearchItem.Command"
args220(17).Value = 3
args220(18).Name = "Quiet"
args220(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args220())

rem ----------------------------------------------------------------------
dim args221(18) as new com.sun.star.beans.PropertyValue
args221(0).Name = "SearchItem.StyleFamily"
args221(0).Value = 2
args221(1).Name = "SearchItem.CellType"
args221(1).Value = 0
args221(2).Name = "SearchItem.RowDirection"
args221(2).Value = true
args221(3).Name = "SearchItem.AllTables"
args221(3).Value = false
args221(4).Name = "SearchItem.Backward"
args221(4).Value = false
args221(5).Name = "SearchItem.Pattern"
args221(5).Value = false
args221(6).Name = "SearchItem.Content"
args221(6).Value = false
args221(7).Name = "SearchItem.AsianOptions"
args221(7).Value = false
args221(8).Name = "SearchItem.AlgorithmType"
args221(8).Value = 0
args221(9).Name = "SearchItem.SearchFlags"
args221(9).Value = 65536
args221(10).Name = "SearchItem.SearchString"
args221(10).Value = "kardenali"
args221(11).Name = "SearchItem.ReplaceString"
args221(11).Value = "kardinali"
args221(12).Name = "SearchItem.Locale"
args221(12).Value = 255
args221(13).Name = "SearchItem.ChangedChars"
args221(13).Value = 2
args221(14).Name = "SearchItem.DeletedChars"
args221(14).Value = 2
args221(15).Name = "SearchItem.InsertedChars"
args221(15).Value = 2
args221(16).Name = "SearchItem.TransliterateFlags"
args221(16).Value = 1280
args221(17).Name = "SearchItem.Command"
args221(17).Value = 3
args221(18).Name = "Quiet"
args221(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args221())

rem ----------------------------------------------------------------------
dim args222(18) as new com.sun.star.beans.PropertyValue
args222(0).Name = "SearchItem.StyleFamily"
args222(0).Value = 2
args222(1).Name = "SearchItem.CellType"
args222(1).Value = 0
args222(2).Name = "SearchItem.RowDirection"
args222(2).Value = true
args222(3).Name = "SearchItem.AllTables"
args222(3).Value = false
args222(4).Name = "SearchItem.Backward"
args222(4).Value = false
args222(5).Name = "SearchItem.Pattern"
args222(5).Value = false
args222(6).Name = "SearchItem.Content"
args222(6).Value = false
args222(7).Name = "SearchItem.AsianOptions"
args222(7).Value = false
args222(8).Name = "SearchItem.AlgorithmType"
args222(8).Value = 0
args222(9).Name = "SearchItem.SearchFlags"
args222(9).Value = 65536
args222(10).Name = "SearchItem.SearchString"
args222(10).Value = "de plus"
args222(11).Name = "SearchItem.ReplaceString"
args222(11).Value = "pose"
args222(12).Name = "SearchItem.Locale"
args222(12).Value = 255
args222(13).Name = "SearchItem.ChangedChars"
args222(13).Value = 2
args222(14).Name = "SearchItem.DeletedChars"
args222(14).Value = 2
args222(15).Name = "SearchItem.InsertedChars"
args222(15).Value = 2
args222(16).Name = "SearchItem.TransliterateFlags"
args222(16).Value = 1280
args222(17).Name = "SearchItem.Command"
args222(17).Value = 3
args222(18).Name = "Quiet"
args222(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args222())

rem ----------------------------------------------------------------------
dim args223(18) as new com.sun.star.beans.PropertyValue
args223(0).Name = "SearchItem.StyleFamily"
args223(0).Value = 2
args223(1).Name = "SearchItem.CellType"
args223(1).Value = 0
args223(2).Name = "SearchItem.RowDirection"
args223(2).Value = true
args223(3).Name = "SearchItem.AllTables"
args223(3).Value = false
args223(4).Name = "SearchItem.Backward"
args223(4).Value = false
args223(5).Name = "SearchItem.Pattern"
args223(5).Value = false
args223(6).Name = "SearchItem.Content"
args223(6).Value = false
args223(7).Name = "SearchItem.AsianOptions"
args223(7).Value = false
args223(8).Name = "SearchItem.AlgorithmType"
args223(8).Value = 0
args223(9).Name = "SearchItem.SearchFlags"
args223(9).Value = 65536
args223(10).Name = "SearchItem.SearchString"
args223(10).Value = "distanco"
args223(11).Name = "SearchItem.ReplaceString"
args223(11).Value = "disto"
args223(12).Name = "SearchItem.Locale"
args223(12).Value = 255
args223(13).Name = "SearchItem.ChangedChars"
args223(13).Value = 2
args223(14).Name = "SearchItem.DeletedChars"
args223(14).Value = 2
args223(15).Name = "SearchItem.InsertedChars"
args223(15).Value = 2
args223(16).Name = "SearchItem.TransliterateFlags"
args223(16).Value = 1280
args223(17).Name = "SearchItem.Command"
args223(17).Value = 3
args223(18).Name = "Quiet"
args223(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args223())

rem ----------------------------------------------------------------------
dim args224(18) as new com.sun.star.beans.PropertyValue
args224(0).Name = "SearchItem.StyleFamily"
args224(0).Value = 2
args224(1).Name = "SearchItem.CellType"
args224(1).Value = 0
args224(2).Name = "SearchItem.RowDirection"
args224(2).Value = true
args224(3).Name = "SearchItem.AllTables"
args224(3).Value = false
args224(4).Name = "SearchItem.Backward"
args224(4).Value = false
args224(5).Name = "SearchItem.Pattern"
args224(5).Value = false
args224(6).Name = "SearchItem.Content"
args224(6).Value = false
args224(7).Name = "SearchItem.AsianOptions"
args224(7).Value = false
args224(8).Name = "SearchItem.AlgorithmType"
args224(8).Value = 0
args224(9).Name = "SearchItem.SearchFlags"
args224(9).Value = 65536
args224(10).Name = "SearchItem.SearchString"
args224(10).Value = "Navarre"
args224(11).Name = "SearchItem.ReplaceString"
args224(11).Value = "Navara"
args224(12).Name = "SearchItem.Locale"
args224(12).Value = 255
args224(13).Name = "SearchItem.ChangedChars"
args224(13).Value = 2
args224(14).Name = "SearchItem.DeletedChars"
args224(14).Value = 2
args224(15).Name = "SearchItem.InsertedChars"
args224(15).Value = 2
args224(16).Name = "SearchItem.TransliterateFlags"
args224(16).Value = 1280
args224(17).Name = "SearchItem.Command"
args224(17).Value = 3
args224(18).Name = "Quiet"
args224(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args224())

rem ----------------------------------------------------------------------
dim args225(18) as new com.sun.star.beans.PropertyValue
args225(0).Name = "SearchItem.StyleFamily"
args225(0).Value = 2
args225(1).Name = "SearchItem.CellType"
args225(1).Value = 0
args225(2).Name = "SearchItem.RowDirection"
args225(2).Value = true
args225(3).Name = "SearchItem.AllTables"
args225(3).Value = false
args225(4).Name = "SearchItem.Backward"
args225(4).Value = false
args225(5).Name = "SearchItem.Pattern"
args225(5).Value = false
args225(6).Name = "SearchItem.Content"
args225(6).Value = false
args225(7).Name = "SearchItem.AsianOptions"
args225(7).Value = false
args225(8).Name = "SearchItem.AlgorithmType"
args225(8).Value = 0
args225(9).Name = "SearchItem.SearchFlags"
args225(9).Value = 65536
args225(10).Name = "SearchItem.SearchString"
args225(10).Value = "Navarra"
args225(11).Name = "SearchItem.ReplaceString"
args225(11).Value = "Navara"
args225(12).Name = "SearchItem.Locale"
args225(12).Value = 255
args225(13).Name = "SearchItem.ChangedChars"
args225(13).Value = 2
args225(14).Name = "SearchItem.DeletedChars"
args225(14).Value = 2
args225(15).Name = "SearchItem.InsertedChars"
args225(15).Value = 2
args225(16).Name = "SearchItem.TransliterateFlags"
args225(16).Value = 1280
args225(17).Name = "SearchItem.Command"
args225(17).Value = 3
args225(18).Name = "Quiet"
args225(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args225())

rem ----------------------------------------------------------------------
dim args226(18) as new com.sun.star.beans.PropertyValue
args226(0).Name = "SearchItem.StyleFamily"
args226(0).Value = 2
args226(1).Name = "SearchItem.CellType"
args226(1).Value = 0
args226(2).Name = "SearchItem.RowDirection"
args226(2).Value = true
args226(3).Name = "SearchItem.AllTables"
args226(3).Value = false
args226(4).Name = "SearchItem.Backward"
args226(4).Value = false
args226(5).Name = "SearchItem.Pattern"
args226(5).Value = false
args226(6).Name = "SearchItem.Content"
args226(6).Value = false
args226(7).Name = "SearchItem.AsianOptions"
args226(7).Value = false
args226(8).Name = "SearchItem.AlgorithmType"
args226(8).Value = 0
args226(9).Name = "SearchItem.SearchFlags"
args226(9).Value = 65536
args226(10).Name = "SearchItem.SearchString"
args226(10).Value = "pluveso"
args226(11).Name = "SearchItem.ReplaceString"
args226(11).Value = "pluvozeso"
args226(12).Name = "SearchItem.Locale"
args226(12).Value = 255
args226(13).Name = "SearchItem.ChangedChars"
args226(13).Value = 2
args226(14).Name = "SearchItem.DeletedChars"
args226(14).Value = 2
args226(15).Name = "SearchItem.InsertedChars"
args226(15).Value = 2
args226(16).Name = "SearchItem.TransliterateFlags"
args226(16).Value = 1280
args226(17).Name = "SearchItem.Command"
args226(17).Value = 3
args226(18).Name = "Quiet"
args226(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args226())

rem ----------------------------------------------------------------------
dim args227(18) as new com.sun.star.beans.PropertyValue
args227(0).Name = "SearchItem.StyleFamily"
args227(0).Value = 2
args227(1).Name = "SearchItem.CellType"
args227(1).Value = 0
args227(2).Name = "SearchItem.RowDirection"
args227(2).Value = true
args227(3).Name = "SearchItem.AllTables"
args227(3).Value = false
args227(4).Name = "SearchItem.Backward"
args227(4).Value = false
args227(5).Name = "SearchItem.Pattern"
args227(5).Value = false
args227(6).Name = "SearchItem.Content"
args227(6).Value = false
args227(7).Name = "SearchItem.AsianOptions"
args227(7).Value = false
args227(8).Name = "SearchItem.AlgorithmType"
args227(8).Value = 0
args227(9).Name = "SearchItem.SearchFlags"
args227(9).Value = 65536
args227(10).Name = "SearchItem.SearchString"
args227(10).Value = "quanto di pluvo"
args227(11).Name = "SearchItem.ReplaceString"
args227(11).Value = "pluvozeso"
args227(12).Name = "SearchItem.Locale"
args227(12).Value = 255
args227(13).Name = "SearchItem.ChangedChars"
args227(13).Value = 2
args227(14).Name = "SearchItem.DeletedChars"
args227(14).Value = 2
args227(15).Name = "SearchItem.InsertedChars"
args227(15).Value = 2
args227(16).Name = "SearchItem.TransliterateFlags"
args227(16).Value = 1280
args227(17).Name = "SearchItem.Command"
args227(17).Value = 3
args227(18).Name = "Quiet"
args227(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args227())

rem ----------------------------------------------------------------------
dim args228(18) as new com.sun.star.beans.PropertyValue
args228(0).Name = "SearchItem.StyleFamily"
args228(0).Value = 2
args228(1).Name = "SearchItem.CellType"
args228(1).Value = 0
args228(2).Name = "SearchItem.RowDirection"
args228(2).Value = true
args228(3).Name = "SearchItem.AllTables"
args228(3).Value = false
args228(4).Name = "SearchItem.Backward"
args228(4).Value = false
args228(5).Name = "SearchItem.Pattern"
args228(5).Value = false
args228(6).Name = "SearchItem.Content"
args228(6).Value = false
args228(7).Name = "SearchItem.AsianOptions"
args228(7).Value = false
args228(8).Name = "SearchItem.AlgorithmType"
args228(8).Value = 0
args228(9).Name = "SearchItem.SearchFlags"
args228(9).Value = 65536
args228(10).Name = "SearchItem.SearchString"
args228(10).Value = "speciala"
args228(11).Name = "SearchItem.ReplaceString"
args228(11).Value = "specala"
args228(12).Name = "SearchItem.Locale"
args228(12).Value = 255
args228(13).Name = "SearchItem.ChangedChars"
args228(13).Value = 2
args228(14).Name = "SearchItem.DeletedChars"
args228(14).Value = 2
args228(15).Name = "SearchItem.InsertedChars"
args228(15).Value = 2
args228(16).Name = "SearchItem.TransliterateFlags"
args228(16).Value = 1280
args228(17).Name = "SearchItem.Command"
args228(17).Value = 3
args228(18).Name = "Quiet"
args228(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args228())

rem ----------------------------------------------------------------------
dim args229(18) as new com.sun.star.beans.PropertyValue
args229(0).Name = "SearchItem.StyleFamily"
args229(0).Value = 2
args229(1).Name = "SearchItem.CellType"
args229(1).Value = 0
args229(2).Name = "SearchItem.RowDirection"
args229(2).Value = true
args229(3).Name = "SearchItem.AllTables"
args229(3).Value = false
args229(4).Name = "SearchItem.Backward"
args229(4).Value = false
args229(5).Name = "SearchItem.Pattern"
args229(5).Value = false
args229(6).Name = "SearchItem.Content"
args229(6).Value = false
args229(7).Name = "SearchItem.AsianOptions"
args229(7).Value = false
args229(8).Name = "SearchItem.AlgorithmType"
args229(8).Value = 0
args229(9).Name = "SearchItem.SearchFlags"
args229(9).Value = 65536
args229(10).Name = "SearchItem.SearchString"
args229(10).Value = "relacion"
args229(11).Name = "SearchItem.ReplaceString"
args229(11).Value = "relat"
args229(12).Name = "SearchItem.Locale"
args229(12).Value = 255
args229(13).Name = "SearchItem.ChangedChars"
args229(13).Value = 2
args229(14).Name = "SearchItem.DeletedChars"
args229(14).Value = 2
args229(15).Name = "SearchItem.InsertedChars"
args229(15).Value = 2
args229(16).Name = "SearchItem.TransliterateFlags"
args229(16).Value = 1280
args229(17).Name = "SearchItem.Command"
args229(17).Value = 3
args229(18).Name = "Quiet"
args229(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args229())

rem ----------------------------------------------------------------------
dim args230(18) as new com.sun.star.beans.PropertyValue
args230(0).Name = "SearchItem.StyleFamily"
args230(0).Value = 2
args230(1).Name = "SearchItem.CellType"
args230(1).Value = 0
args230(2).Name = "SearchItem.RowDirection"
args230(2).Value = true
args230(3).Name = "SearchItem.AllTables"
args230(3).Value = false
args230(4).Name = "SearchItem.Backward"
args230(4).Value = false
args230(5).Name = "SearchItem.Pattern"
args230(5).Value = false
args230(6).Name = "SearchItem.Content"
args230(6).Value = false
args230(7).Name = "SearchItem.AsianOptions"
args230(7).Value = false
args230(8).Name = "SearchItem.AlgorithmType"
args230(8).Value = 0
args230(9).Name = "SearchItem.SearchFlags"
args230(9).Value = 65536
args230(10).Name = "SearchItem.SearchString"
args230(10).Value = "senator"
args230(11).Name = "SearchItem.ReplaceString"
args230(11).Value = "senatan"
args230(12).Name = "SearchItem.Locale"
args230(12).Value = 255
args230(13).Name = "SearchItem.ChangedChars"
args230(13).Value = 2
args230(14).Name = "SearchItem.DeletedChars"
args230(14).Value = 2
args230(15).Name = "SearchItem.InsertedChars"
args230(15).Value = 2
args230(16).Name = "SearchItem.TransliterateFlags"
args230(16).Value = 1280
args230(17).Name = "SearchItem.Command"
args230(17).Value = 3
args230(18).Name = "Quiet"
args230(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args230())

rem ----------------------------------------------------------------------
dim args231(18) as new com.sun.star.beans.PropertyValue
args231(0).Name = "SearchItem.StyleFamily"
args231(0).Value = 2
args231(1).Name = "SearchItem.CellType"
args231(1).Value = 0
args231(2).Name = "SearchItem.RowDirection"
args231(2).Value = true
args231(3).Name = "SearchItem.AllTables"
args231(3).Value = false
args231(4).Name = "SearchItem.Backward"
args231(4).Value = false
args231(5).Name = "SearchItem.Pattern"
args231(5).Value = false
args231(6).Name = "SearchItem.Content"
args231(6).Value = false
args231(7).Name = "SearchItem.AsianOptions"
args231(7).Value = false
args231(8).Name = "SearchItem.AlgorithmType"
args231(8).Value = 0
args231(9).Name = "SearchItem.SearchFlags"
args231(9).Value = 65536
args231(10).Name = "SearchItem.SearchString"
args231(10).Value = "mortita"
args231(11).Name = "SearchItem.ReplaceString"
args231(11).Value = "mortinta"
args231(12).Name = "SearchItem.Locale"
args231(12).Value = 255
args231(13).Name = "SearchItem.ChangedChars"
args231(13).Value = 2
args231(14).Name = "SearchItem.DeletedChars"
args231(14).Value = 2
args231(15).Name = "SearchItem.InsertedChars"
args231(15).Value = 2
args231(16).Name = "SearchItem.TransliterateFlags"
args231(16).Value = 1280
args231(17).Name = "SearchItem.Command"
args231(17).Value = 3
args231(18).Name = "Quiet"
args231(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args231())

rem ----------------------------------------------------------------------
dim args232(18) as new com.sun.star.beans.PropertyValue

dim FORMAC(2), SUBIS(2) as string
FORMAC(0) = "ata da ": SUBIS(0) = "ata per "
FORMAC(1) = "ita da ": SUBIS(1) = "ita per "
FORMAC(2) = "esas da ": SUBIS(2) = "esas per "

For I = 0 to 2

args232(0).Name = "SearchItem.StyleFamily"
args232(0).Value = 2
args232(1).Name = "SearchItem.CellType"
args232(1).Value = 0
args232(2).Name = "SearchItem.RowDirection"
args232(2).Value = true
args232(3).Name = "SearchItem.AllTables"
args232(3).Value = false
args232(4).Name = "SearchItem.Backward"
args232(4).Value = false
args232(5).Name = "SearchItem.Pattern"
args232(5).Value = false
args232(6).Name = "SearchItem.Content"
args232(6).Value = false
args232(7).Name = "SearchItem.AsianOptions"
args232(7).Value = false
args232(8).Name = "SearchItem.AlgorithmType"
args232(8).Value = 0
args232(9).Name = "SearchItem.SearchFlags"
args232(9).Value = 65536
args232(10).Name = "SearchItem.SearchString"
args232(10).Value = "formac"+FORMAC(I)
args232(11).Name = "SearchItem.ReplaceString"
args232(11).Value = "formac"+SUBIS(I)
args232(12).Name = "SearchItem.Locale"
args232(12).Value = 255
args232(13).Name = "SearchItem.ChangedChars"
args232(13).Value = 2
args232(14).Name = "SearchItem.DeletedChars"
args232(14).Value = 2
args232(15).Name = "SearchItem.InsertedChars"
args232(15).Value = 2
args232(16).Name = "SearchItem.TransliterateFlags"
args232(16).Value = 1280
args232(17).Name = "SearchItem.Command"
args232(17).Value = 3
args232(18).Name = "Quiet"
args232(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args232())

Next I

rem ----------------------------------------------------------------------
dim args233(18) as new com.sun.star.beans.PropertyValue
args233(0).Name = "SearchItem.StyleFamily"
args233(0).Value = 2
args233(1).Name = "SearchItem.CellType"
args233(1).Value = 0
args233(2).Name = "SearchItem.RowDirection"
args233(2).Value = true
args233(3).Name = "SearchItem.AllTables"
args233(3).Value = false
args233(4).Name = "SearchItem.Backward"
args233(4).Value = false
args233(5).Name = "SearchItem.Pattern"
args233(5).Value = false
args233(6).Name = "SearchItem.Content"
args233(6).Value = false
args233(7).Name = "SearchItem.AsianOptions"
args233(7).Value = false
args233(8).Name = "SearchItem.AlgorithmType"
args233(8).Value = 0
args233(9).Name = "SearchItem.SearchFlags"
args233(9).Value = 65536
args233(10).Name = "SearchItem.SearchString"
args233(10).Value = "nomesita"
args233(11).Name = "SearchItem.ReplaceString"
args233(11).Value = "nomizita"
args233(12).Name = "SearchItem.Locale"
args233(12).Value = 255
args233(13).Name = "SearchItem.ChangedChars"
args233(13).Value = 2
args233(14).Name = "SearchItem.DeletedChars"
args233(14).Value = 2
args233(15).Name = "SearchItem.InsertedChars"
args233(15).Value = 2
args233(16).Name = "SearchItem.TransliterateFlags"
args233(16).Value = 1280
args233(17).Name = "SearchItem.Command"
args233(17).Value = 3
args233(18).Name = "Quiet"
args233(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args233())

rem ----------------------------------------------------------------------
dim args234(18) as new com.sun.star.beans.PropertyValue
args234(0).Name = "SearchItem.StyleFamily"
args234(0).Value = 2
args234(1).Name = "SearchItem.CellType"
args234(1).Value = 0
args234(2).Name = "SearchItem.RowDirection"
args234(2).Value = true
args234(3).Name = "SearchItem.AllTables"
args234(3).Value = false
args234(4).Name = "SearchItem.Backward"
args234(4).Value = false
args234(5).Name = "SearchItem.Pattern"
args234(5).Value = false
args234(6).Name = "SearchItem.Content"
args234(6).Value = false
args234(7).Name = "SearchItem.AsianOptions"
args234(7).Value = false
args234(8).Name = "SearchItem.AlgorithmType"
args234(8).Value = 0
args234(9).Name = "SearchItem.SearchFlags"
args234(9).Value = 65536
args234(10).Name = "SearchItem.SearchString"
args234(10).Value = "Sao Tome"
args234(11).Name = "SearchItem.ReplaceString"
args234(11).Value = "San-Tome"
args234(12).Name = "SearchItem.Locale"
args234(12).Value = 255
args234(13).Name = "SearchItem.ChangedChars"
args234(13).Value = 2
args234(14).Name = "SearchItem.DeletedChars"
args234(14).Value = 2
args234(15).Name = "SearchItem.InsertedChars"
args234(15).Value = 2
args234(16).Name = "SearchItem.TransliterateFlags"
args234(16).Value = 1280
args234(17).Name = "SearchItem.Command"
args234(17).Value = 3
args234(18).Name = "Quiet"
args234(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args234())

rem ----------------------------------------------------------------------
dim args235(18) as new com.sun.star.beans.PropertyValue
args235(0).Name = "SearchItem.StyleFamily"
args235(0).Value = 2
args235(1).Name = "SearchItem.CellType"
args235(1).Value = 0
args235(2).Name = "SearchItem.RowDirection"
args235(2).Value = true
args235(3).Name = "SearchItem.AllTables"
args235(3).Value = false
args235(4).Name = "SearchItem.Backward"
args235(4).Value = false
args235(5).Name = "SearchItem.Pattern"
args235(5).Value = false
args235(6).Name = "SearchItem.Content"
args235(6).Value = false
args235(7).Name = "SearchItem.AsianOptions"
args235(7).Value = false
args235(8).Name = "SearchItem.AlgorithmType"
args235(8).Value = 0
args235(9).Name = "SearchItem.SearchFlags"
args235(9).Value = 65536
args235(10).Name = "SearchItem.SearchString"
args235(10).Value = "konstitucata da"
args235(11).Name = "SearchItem.ReplaceString"
args235(11).Value = "konstitucata per"
args235(12).Name = "SearchItem.Locale"
args235(12).Value = 255
args235(13).Name = "SearchItem.ChangedChars"
args235(13).Value = 2
args235(14).Name = "SearchItem.DeletedChars"
args235(14).Value = 2
args235(15).Name = "SearchItem.InsertedChars"
args235(15).Value = 2
args235(16).Name = "SearchItem.TransliterateFlags"
args235(16).Value = 1280
args235(17).Name = "SearchItem.Command"
args235(17).Value = 3
args235(18).Name = "Quiet"
args235(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args235())

rem ----------------------------------------------------------------------
dim args236(18) as new com.sun.star.beans.PropertyValue
args236(0).Name = "SearchItem.StyleFamily"
args236(0).Value = 2
args236(1).Name = "SearchItem.CellType"
args236(1).Value = 0
args236(2).Name = "SearchItem.RowDirection"
args236(2).Value = true
args236(3).Name = "SearchItem.AllTables"
args236(3).Value = false
args236(4).Name = "SearchItem.Backward"
args236(4).Value = false
args236(5).Name = "SearchItem.Pattern"
args236(5).Value = false
args236(6).Name = "SearchItem.Content"
args236(6).Value = false
args236(7).Name = "SearchItem.AsianOptions"
args236(7).Value = false
args236(8).Name = "SearchItem.AlgorithmType"
args236(8).Value = 0
args236(9).Name = "SearchItem.SearchFlags"
args236(9).Value = 65536
args236(10).Name = "SearchItem.SearchString"
args236(10).Value = "Nederlanda "
args236(11).Name = "SearchItem.ReplaceString"
args236(11).Value = "Nederlandana "
args236(12).Name = "SearchItem.Locale"
args236(12).Value = 255
args236(13).Name = "SearchItem.ChangedChars"
args236(13).Value = 2
args236(14).Name = "SearchItem.DeletedChars"
args236(14).Value = 2
args236(15).Name = "SearchItem.InsertedChars"
args236(15).Value = 2
args236(16).Name = "SearchItem.TransliterateFlags"
args236(16).Value = 1280
args236(17).Name = "SearchItem.Command"
args236(17).Value = 3
args236(18).Name = "Quiet"
args236(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args236())

rem ----------------------------------------------------------------------
dim args237(18) as new com.sun.star.beans.PropertyValue
args237(0).Name = "SearchItem.StyleFamily"
args237(0).Value = 2
args237(1).Name = "SearchItem.CellType"
args237(1).Value = 0
args237(2).Name = "SearchItem.RowDirection"
args237(2).Value = true
args237(3).Name = "SearchItem.AllTables"
args237(3).Value = false
args237(4).Name = "SearchItem.Backward"
args237(4).Value = false
args237(5).Name = "SearchItem.Pattern"
args237(5).Value = false
args237(6).Name = "SearchItem.Content"
args237(6).Value = false
args237(7).Name = "SearchItem.AsianOptions"
args237(7).Value = false
args237(8).Name = "SearchItem.AlgorithmType"
args237(8).Value = 0
args237(9).Name = "SearchItem.SearchFlags"
args237(9).Value = 65536
args237(10).Name = "SearchItem.SearchString"
args237(10).Value = " en quo "
args237(11).Name = "SearchItem.ReplaceString"
args237(11).Value = " ube "
args237(12).Name = "SearchItem.Locale"
args237(12).Value = 255
args237(13).Name = "SearchItem.ChangedChars"
args237(13).Value = 2
args237(14).Name = "SearchItem.DeletedChars"
args237(14).Value = 2
args237(15).Name = "SearchItem.InsertedChars"
args237(15).Value = 2
args237(16).Name = "SearchItem.TransliterateFlags"
args237(16).Value = 1280
args237(17).Name = "SearchItem.Command"
args237(17).Value = 3
args237(18).Name = "Quiet"
args237(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args237())

rem ----------------------------------------------------------------------
dim args238(18) as new com.sun.star.beans.PropertyValue
args238(0).Name = "SearchItem.StyleFamily"
args238(0).Value = 2
args238(1).Name = "SearchItem.CellType"
args238(1).Value = 0
args238(2).Name = "SearchItem.RowDirection"
args238(2).Value = true
args238(3).Name = "SearchItem.AllTables"
args238(3).Value = false
args238(4).Name = "SearchItem.Backward"
args238(4).Value = false
args238(5).Name = "SearchItem.Pattern"
args238(5).Value = false
args238(6).Name = "SearchItem.Content"
args238(6).Value = false
args238(7).Name = "SearchItem.AsianOptions"
args238(7).Value = false
args238(8).Name = "SearchItem.AlgorithmType"
args238(8).Value = 0
args238(9).Name = "SearchItem.SearchFlags"
args238(9).Value = 65536
args238(10).Name = "SearchItem.SearchString"
args238(10).Value = "dekyar"
args238(11).Name = "SearchItem.ReplaceString"
args238(11).Value = "yardek"
args238(12).Name = "SearchItem.Locale"
args238(12).Value = 255
args238(13).Name = "SearchItem.ChangedChars"
args238(13).Value = 2
args238(14).Name = "SearchItem.DeletedChars"
args238(14).Value = 2
args238(15).Name = "SearchItem.InsertedChars"
args238(15).Value = 2
args238(16).Name = "SearchItem.TransliterateFlags"
args238(16).Value = 1280
args238(17).Name = "SearchItem.Command"
args238(17).Value = 3
args238(18).Name = "Quiet"
args238(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args238())

rem ----------------------------------------------------------------------
dim args239(18) as new com.sun.star.beans.PropertyValue
args239(0).Name = "SearchItem.StyleFamily"
args239(0).Value = 2
args239(1).Name = "SearchItem.CellType"
args239(1).Value = 0
args239(2).Name = "SearchItem.RowDirection"
args239(2).Value = true
args239(3).Name = "SearchItem.AllTables"
args239(3).Value = false
args239(4).Name = "SearchItem.Backward"
args239(4).Value = false
args239(5).Name = "SearchItem.Pattern"
args239(5).Value = false
args239(6).Name = "SearchItem.Content"
args239(6).Value = false
args239(7).Name = "SearchItem.AsianOptions"
args239(7).Value = false
args239(8).Name = "SearchItem.AlgorithmType"
args239(8).Value = 0
args239(9).Name = "SearchItem.SearchFlags"
args239(9).Value = 65536
args239(10).Name = "SearchItem.SearchString"
args239(10).Value = "okazis"
args239(11).Name = "SearchItem.ReplaceString"
args239(11).Value = "eventis"
args239(12).Name = "SearchItem.Locale"
args239(12).Value = 255
args239(13).Name = "SearchItem.ChangedChars"
args239(13).Value = 2
args239(14).Name = "SearchItem.DeletedChars"
args239(14).Value = 2
args239(15).Name = "SearchItem.InsertedChars"
args239(15).Value = 2
args239(16).Name = "SearchItem.TransliterateFlags"
args239(16).Value = 1280
args239(17).Name = "SearchItem.Command"
args239(17).Value = 3
args239(18).Name = "Quiet"
args239(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args239())

rem ----------------------------------------------------------------------
dim args240(18) as new com.sun.star.beans.PropertyValue
args240(0).Name = "SearchItem.StyleFamily"
args240(0).Value = 2
args240(1).Name = "SearchItem.CellType"
args240(1).Value = 0
args240(2).Name = "SearchItem.RowDirection"
args240(2).Value = true
args240(3).Name = "SearchItem.AllTables"
args240(3).Value = false
args240(4).Name = "SearchItem.Backward"
args240(4).Value = false
args240(5).Name = "SearchItem.Pattern"
args240(5).Value = false
args240(6).Name = "SearchItem.Content"
args240(6).Value = false
args240(7).Name = "SearchItem.AsianOptions"
args240(7).Value = false
args240(8).Name = "SearchItem.AlgorithmType"
args240(8).Value = 0
args240(9).Name = "SearchItem.SearchFlags"
args240(9).Value = 65536
args240(10).Name = "SearchItem.SearchString"
args240(10).Value = "perzono"
args240(11).Name = "SearchItem.ReplaceString"
args240(11).Value = "persono"
args240(12).Name = "SearchItem.Locale"
args240(12).Value = 255
args240(13).Name = "SearchItem.ChangedChars"
args240(13).Value = 2
args240(14).Name = "SearchItem.DeletedChars"
args240(14).Value = 2
args240(15).Name = "SearchItem.InsertedChars"
args240(15).Value = 2
args240(16).Name = "SearchItem.TransliterateFlags"
args240(16).Value = 1280
args240(17).Name = "SearchItem.Command"
args240(17).Value = 3
args240(18).Name = "Quiet"
args240(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args240())

rem ----------------------------------------------------------------------
dim args241(18) as new com.sun.star.beans.PropertyValue
args241(0).Name = "SearchItem.StyleFamily"
args241(0).Value = 2
args241(1).Name = "SearchItem.CellType"
args241(1).Value = 0
args241(2).Name = "SearchItem.RowDirection"
args241(2).Value = true
args241(3).Name = "SearchItem.AllTables"
args241(3).Value = false
args241(4).Name = "SearchItem.Backward"
args241(4).Value = false
args241(5).Name = "SearchItem.Pattern"
args241(5).Value = false
args241(6).Name = "SearchItem.Content"
args241(6).Value = false
args241(7).Name = "SearchItem.AsianOptions"
args241(7).Value = false
args241(8).Name = "SearchItem.AlgorithmType"
args241(8).Value = 0
args241(9).Name = "SearchItem.SearchFlags"
args241(9).Value = 65536
args241(10).Name = "SearchItem.SearchString"
args241(10).Value = "nomita"
args241(11).Name = "SearchItem.ReplaceString"
args241(11).Value = "nomizita"
args241(12).Name = "SearchItem.Locale"
args241(12).Value = 255
args241(13).Name = "SearchItem.ChangedChars"
args241(13).Value = 2
args241(14).Name = "SearchItem.DeletedChars"
args241(14).Value = 2
args241(15).Name = "SearchItem.InsertedChars"
args241(15).Value = 2
args241(16).Name = "SearchItem.TransliterateFlags"
args241(16).Value = 1280
args241(17).Name = "SearchItem.Command"
args241(17).Value = 3
args241(18).Name = "Quiet"
args241(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args241())

rem ----------------------------------------------------------------------
dim args242(18) as new com.sun.star.beans.PropertyValue
args242(0).Name = "SearchItem.StyleFamily"
args242(0).Value = 2
args242(1).Name = "SearchItem.CellType"
args242(1).Value = 0
args242(2).Name = "SearchItem.RowDirection"
args242(2).Value = true
args242(3).Name = "SearchItem.AllTables"
args242(3).Value = false
args242(4).Name = "SearchItem.Backward"
args242(4).Value = false
args242(5).Name = "SearchItem.Pattern"
args242(5).Value = false
args242(6).Name = "SearchItem.Content"
args242(6).Value = false
args242(7).Name = "SearchItem.AsianOptions"
args242(7).Value = false
args242(8).Name = "SearchItem.AlgorithmType"
args242(8).Value = 0
args242(9).Name = "SearchItem.SearchFlags"
args242(9).Value = 65536
args242(10).Name = "SearchItem.SearchString"
args242(10).Value = "nomata"
args242(11).Name = "SearchItem.ReplaceString"
args242(11).Value = "nomizita"
args242(12).Name = "SearchItem.Locale"
args242(12).Value = 255
args242(13).Name = "SearchItem.ChangedChars"
args242(13).Value = 2
args242(14).Name = "SearchItem.DeletedChars"
args242(14).Value = 2
args242(15).Name = "SearchItem.InsertedChars"
args242(15).Value = 2
args242(16).Name = "SearchItem.TransliterateFlags"
args242(16).Value = 1280
args242(17).Name = "SearchItem.Command"
args242(17).Value = 3
args242(18).Name = "Quiet"
args242(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args242())

rem ----------------------------------------------------------------------
dim args243(18) as new com.sun.star.beans.PropertyValue
args243(0).Name = "SearchItem.StyleFamily"
args243(0).Value = 2
args243(1).Name = "SearchItem.CellType"
args243(1).Value = 0
args243(2).Name = "SearchItem.RowDirection"
args243(2).Value = true
args243(3).Name = "SearchItem.AllTables"
args243(3).Value = false
args243(4).Name = "SearchItem.Backward"
args243(4).Value = false
args243(5).Name = "SearchItem.Pattern"
args243(5).Value = false
args243(6).Name = "SearchItem.Content"
args243(6).Value = false
args243(7).Name = "SearchItem.AsianOptions"
args243(7).Value = false
args243(8).Name = "SearchItem.AlgorithmType"
args243(8).Value = 0
args243(9).Name = "SearchItem.SearchFlags"
args243(9).Value = 65536
args243(10).Name = "SearchItem.SearchString"
args243(10).Value = "proponis"
args243(11).Name = "SearchItem.ReplaceString"
args243(11).Value = "propozis"
args243(12).Name = "SearchItem.Locale"
args243(12).Value = 255
args243(13).Name = "SearchItem.ChangedChars"
args243(13).Value = 2
args243(14).Name = "SearchItem.DeletedChars"
args243(14).Value = 2
args243(15).Name = "SearchItem.InsertedChars"
args243(15).Value = 2
args243(16).Name = "SearchItem.TransliterateFlags"
args243(16).Value = 1280
args243(17).Name = "SearchItem.Command"
args243(17).Value = 3
args243(18).Name = "Quiet"
args243(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args243())

rem ----------------------------------------------------------------------
dim args244(18) as new com.sun.star.beans.PropertyValue
args244(0).Name = "SearchItem.StyleFamily"
args244(0).Value = 2
args244(1).Name = "SearchItem.CellType"
args244(1).Value = 0
args244(2).Name = "SearchItem.RowDirection"
args244(2).Value = true
args244(3).Name = "SearchItem.AllTables"
args244(3).Value = false
args244(4).Name = "SearchItem.Backward"
args244(4).Value = false
args244(5).Name = "SearchItem.Pattern"
args244(5).Value = false
args244(6).Name = "SearchItem.Content"
args244(6).Value = false
args244(7).Name = "SearchItem.AsianOptions"
args244(7).Value = false
args244(8).Name = "SearchItem.AlgorithmType"
args244(8).Value = 0
args244(9).Name = "SearchItem.SearchFlags"
args244(9).Value = 65536
args244(10).Name = "SearchItem.SearchString"
args244(10).Value = "Pacifik insulani"
args244(11).Name = "SearchItem.ReplaceString"
args244(11).Value = "Pacifik-insulani"
args244(12).Name = "SearchItem.Locale"
args244(12).Value = 255
args244(13).Name = "SearchItem.ChangedChars"
args244(13).Value = 2
args244(14).Name = "SearchItem.DeletedChars"
args244(14).Value = 2
args244(15).Name = "SearchItem.InsertedChars"
args244(15).Value = 2
args244(16).Name = "SearchItem.TransliterateFlags"
args244(16).Value = 1280
args244(17).Name = "SearchItem.Command"
args244(17).Value = 3
args244(18).Name = "Quiet"
args244(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args244())

rem ----------------------------------------------------------------------
dim args245(18) as new com.sun.star.beans.PropertyValue
args245(0).Name = "SearchItem.StyleFamily"
args245(0).Value = 2
args245(1).Name = "SearchItem.CellType"
args245(1).Value = 0
args245(2).Name = "SearchItem.RowDirection"
args245(2).Value = true
args245(3).Name = "SearchItem.AllTables"
args245(3).Value = false
args245(4).Name = "SearchItem.Backward"
args245(4).Value = false
args245(5).Name = "SearchItem.Pattern"
args245(5).Value = false
args245(6).Name = "SearchItem.Content"
args245(6).Value = false
args245(7).Name = "SearchItem.AsianOptions"
args245(7).Value = false
args245(8).Name = "SearchItem.AlgorithmType"
args245(8).Value = 0
args245(9).Name = "SearchItem.SearchFlags"
args245(9).Value = 65536
args245(10).Name = "SearchItem.SearchString"
args245(10).Value = "Latinal Amerikani"
args245(11).Name = "SearchItem.ReplaceString"
args245(11).Value = "Latin-Amerikani"
args245(12).Name = "SearchItem.Locale"
args245(12).Value = 255
args245(13).Name = "SearchItem.ChangedChars"
args245(13).Value = 2
args245(14).Name = "SearchItem.DeletedChars"
args245(14).Value = 2
args245(15).Name = "SearchItem.InsertedChars"
args245(15).Value = 2
args245(16).Name = "SearchItem.TransliterateFlags"
args245(16).Value = 1280
args245(17).Name = "SearchItem.Command"
args245(17).Value = 3
args245(18).Name = "Quiet"
args245(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args245())

rem ----------------------------------------------------------------------
dim args246(18) as new com.sun.star.beans.PropertyValue
args246(0).Name = "SearchItem.StyleFamily"
args246(0).Value = 2
args246(1).Name = "SearchItem.CellType"
args246(1).Value = 0
args246(2).Name = "SearchItem.RowDirection"
args246(2).Value = true
args246(3).Name = "SearchItem.AllTables"
args246(3).Value = false
args246(4).Name = "SearchItem.Backward"
args246(4).Value = false
args246(5).Name = "SearchItem.Pattern"
args246(5).Value = false
args246(6).Name = "SearchItem.Content"
args246(6).Value = false
args246(7).Name = "SearchItem.AsianOptions"
args246(7).Value = false
args246(8).Name = "SearchItem.AlgorithmType"
args246(8).Value = 0
args246(9).Name = "SearchItem.SearchFlags"
args246(9).Value = 65536
args246(10).Name = "SearchItem.SearchString"
args246(10).Value = "domo-maestro"
args246(11).Name = "SearchItem.ReplaceString"
args246(11).Value = "domo-chefo"
args246(12).Name = "SearchItem.Locale"
args246(12).Value = 255
args246(13).Name = "SearchItem.ChangedChars"
args246(13).Value = 2
args246(14).Name = "SearchItem.DeletedChars"
args246(14).Value = 2
args246(15).Name = "SearchItem.InsertedChars"
args246(15).Value = 2
args246(16).Name = "SearchItem.TransliterateFlags"
args246(16).Value = 1280
args246(17).Name = "SearchItem.Command"
args246(17).Value = 3
args246(18).Name = "Quiet"
args246(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args246())

rem ----------------------------------------------------------------------
dim args247(18) as new com.sun.star.beans.PropertyValue
args247(0).Name = "SearchItem.StyleFamily"
args247(0).Value = 2
args247(1).Name = "SearchItem.CellType"
args247(1).Value = 0
args247(2).Name = "SearchItem.RowDirection"
args247(2).Value = true
args247(3).Name = "SearchItem.AllTables"
args247(3).Value = false
args247(4).Name = "SearchItem.Backward"
args247(4).Value = false
args247(5).Name = "SearchItem.Pattern"
args247(5).Value = false
args247(6).Name = "SearchItem.Content"
args247(6).Value = false
args247(7).Name = "SearchItem.AsianOptions"
args247(7).Value = false
args247(8).Name = "SearchItem.AlgorithmType"
args247(8).Value = 0
args247(9).Name = "SearchItem.SearchFlags"
args247(9).Value = 65536
args247(10).Name = "SearchItem.SearchString"
args247(10).Value = "senatori"
args247(11).Name = "SearchItem.ReplaceString"
args247(11).Value = "senatan"
args247(12).Name = "SearchItem.Locale"
args247(12).Value = 255
args247(13).Name = "SearchItem.ChangedChars"
args247(13).Value = 2
args247(14).Name = "SearchItem.DeletedChars"
args247(14).Value = 2
args247(15).Name = "SearchItem.InsertedChars"
args247(15).Value = 2
args247(16).Name = "SearchItem.TransliterateFlags"
args247(16).Value = 1024
args247(17).Name = "SearchItem.Command"
args247(17).Value = 3
args247(18).Name = "Quiet"
args247(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args247())

rem ----------------------------------------------------------------------
dim args248(18) as new com.sun.star.beans.PropertyValue
args248(0).Name = "SearchItem.StyleFamily"
args248(0).Value = 2
args248(1).Name = "SearchItem.CellType"
args248(1).Value = 0
args248(2).Name = "SearchItem.RowDirection"
args248(2).Value = true
args248(3).Name = "SearchItem.AllTables"
args248(3).Value = false
args248(4).Name = "SearchItem.Backward"
args248(4).Value = false
args248(5).Name = "SearchItem.Pattern"
args248(5).Value = false
args248(6).Name = "SearchItem.Content"
args248(6).Value = false
args248(7).Name = "SearchItem.AsianOptions"
args248(7).Value = false
args248(8).Name = "SearchItem.AlgorithmType"
args248(8).Value = 0
args248(9).Name = "SearchItem.SearchFlags"
args248(9).Value = 65536
args248(10).Name = "SearchItem.SearchString"
args248(10).Value = "California"
args248(11).Name = "SearchItem.ReplaceString"
args248(11).Value = "Kalifornia"
args248(12).Name = "SearchItem.Locale"
args248(12).Value = 255
args248(13).Name = "SearchItem.ChangedChars"
args248(13).Value = 2
args248(14).Name = "SearchItem.DeletedChars"
args248(14).Value = 2
args248(15).Name = "SearchItem.InsertedChars"
args248(15).Value = 2
args248(16).Name = "SearchItem.TransliterateFlags"
args248(16).Value = 1024
args248(17).Name = "SearchItem.Command"
args248(17).Value = 3
args248(18).Name = "Quiet"
args248(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args248())

rem ----------------------------------------------------------------------
dim args249(18) as new com.sun.star.beans.PropertyValue
args249(0).Name = "SearchItem.StyleFamily"
args249(0).Value = 2
args249(1).Name = "SearchItem.CellType"
args249(1).Value = 0
args249(2).Name = "SearchItem.RowDirection"
args249(2).Value = true
args249(3).Name = "SearchItem.AllTables"
args249(3).Value = false
args249(4).Name = "SearchItem.Backward"
args249(4).Value = false
args249(5).Name = "SearchItem.Pattern"
args249(5).Value = false
args249(6).Name = "SearchItem.Content"
args249(6).Value = false
args249(7).Name = "SearchItem.AsianOptions"
args249(7).Value = false
args249(8).Name = "SearchItem.AlgorithmType"
args249(8).Value = 0
args249(9).Name = "SearchItem.SearchFlags"
args249(9).Value = 65536
args249(10).Name = "SearchItem.SearchString"
args249(10).Value = "reprizent"
args249(11).Name = "SearchItem.ReplaceString"
args249(11).Value = "reprezent"
args249(12).Name = "SearchItem.Locale"
args249(12).Value = 255
args249(13).Name = "SearchItem.ChangedChars"
args249(13).Value = 2
args249(14).Name = "SearchItem.DeletedChars"
args249(14).Value = 2
args249(15).Name = "SearchItem.InsertedChars"
args249(15).Value = 2
args249(16).Name = "SearchItem.TransliterateFlags"
args249(16).Value = 1024
args249(17).Name = "SearchItem.Command"
args249(17).Value = 3
args249(18).Name = "Quiet"
args249(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args249())

rem ----------------------------------------------------------------------
dim args250(18) as new com.sun.star.beans.PropertyValue
args250(0).Name = "SearchItem.StyleFamily"
args250(0).Value = 2
args250(1).Name = "SearchItem.CellType"
args250(1).Value = 0
args250(2).Name = "SearchItem.RowDirection"
args250(2).Value = true
args250(3).Name = "SearchItem.AllTables"
args250(3).Value = false
args250(4).Name = "SearchItem.Backward"
args250(4).Value = false
args250(5).Name = "SearchItem.Pattern"
args250(5).Value = false
args250(6).Name = "SearchItem.Content"
args250(6).Value = false
args250(7).Name = "SearchItem.AsianOptions"
args250(7).Value = false
args250(8).Name = "SearchItem.AlgorithmType"
args250(8).Value = 0
args250(9).Name = "SearchItem.SearchFlags"
args250(9).Value = 65536
args250(10).Name = "SearchItem.SearchString"
args250(10).Value = "habitesis per singla individui"
args250(11).Name = "SearchItem.ReplaceString"
args250(11).Value = "de omna hemanari vivis unika sola individuo, e"
args250(12).Name = "SearchItem.Locale"
args250(12).Value = 255
args250(13).Name = "SearchItem.ChangedChars"
args250(13).Value = 2
args250(14).Name = "SearchItem.DeletedChars"
args250(14).Value = 2
args250(15).Name = "SearchItem.InsertedChars"
args250(15).Value = 2
args250(16).Name = "SearchItem.TransliterateFlags"
args250(16).Value = 1024
args250(17).Name = "SearchItem.Command"
args250(17).Value = 3
args250(18).Name = "Quiet"
args250(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args250())

rem ----------------------------------------------------------------------
dim args251(18) as new com.sun.star.beans.PropertyValue
args251(0).Name = "SearchItem.StyleFamily"
args251(0).Value = 2
args251(1).Name = "SearchItem.CellType"
args251(1).Value = 0
args251(2).Name = "SearchItem.RowDirection"
args251(2).Value = true
args251(3).Name = "SearchItem.AllTables"
args251(3).Value = false
args251(4).Name = "SearchItem.Backward"
args251(4).Value = false
args251(5).Name = "SearchItem.Pattern"
args251(5).Value = false
args251(6).Name = "SearchItem.Content"
args251(6).Value = false
args251(7).Name = "SearchItem.AsianOptions"
args251(7).Value = false
args251(8).Name = "SearchItem.AlgorithmType"
args251(8).Value = 0
args251(9).Name = "SearchItem.SearchFlags"
args251(9).Value = 65536
args251(10).Name = "SearchItem.SearchString"
args251(10).Value = "Segun [[Usana"
args251(11).Name = "SearchItem.ReplaceString"
args251(11).Value = "Segun l'[[Usana"
args251(12).Name = "SearchItem.Locale"
args251(12).Value = 255
args251(13).Name = "SearchItem.ChangedChars"
args251(13).Value = 2
args251(14).Name = "SearchItem.DeletedChars"
args251(14).Value = 2
args251(15).Name = "SearchItem.InsertedChars"
args251(15).Value = 2
args251(16).Name = "SearchItem.TransliterateFlags"
args251(16).Value = 1280
args251(17).Name = "SearchItem.Command"
args251(17).Value = 3
args251(18).Name = "Quiet"
args251(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args251())

rem ----------------------------------------------------------------------
dim args252(18) as new com.sun.star.beans.PropertyValue
args252(0).Name = "SearchItem.StyleFamily"
args252(0).Value = 2
args252(1).Name = "SearchItem.CellType"
args252(1).Value = 0
args252(2).Name = "SearchItem.RowDirection"
args252(2).Value = true
args252(3).Name = "SearchItem.AllTables"
args252(3).Value = false
args252(4).Name = "SearchItem.Backward"
args252(4).Value = false
args252(5).Name = "SearchItem.Pattern"
args252(5).Value = false
args252(6).Name = "SearchItem.Content"
args252(6).Value = false
args252(7).Name = "SearchItem.AsianOptions"
args252(7).Value = false
args252(8).Name = "SearchItem.AlgorithmType"
args252(8).Value = 0
args252(9).Name = "SearchItem.SearchFlags"
args252(9).Value = 65536
args252(10).Name = "SearchItem.SearchString"
args252(10).Value = "diferenta"
args252(11).Name = "SearchItem.ReplaceString"
args252(11).Value = "diferanta"
args252(12).Name = "SearchItem.Locale"
args252(12).Value = 255
args252(13).Name = "SearchItem.ChangedChars"
args252(13).Value = 2
args252(14).Name = "SearchItem.DeletedChars"
args252(14).Value = 2
args252(15).Name = "SearchItem.InsertedChars"
args252(15).Value = 2
args252(16).Name = "SearchItem.TransliterateFlags"
args252(16).Value = 1024
args252(17).Name = "SearchItem.Command"
args252(17).Value = 3
args252(18).Name = "Quiet"
args252(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args252())

rem ----------------------------------------------------------------------
dim args253(18) as new com.sun.star.beans.PropertyValue
args253(0).Name = "SearchItem.StyleFamily"
args253(0).Value = 2
args253(1).Name = "SearchItem.CellType"
args253(1).Value = 0
args253(2).Name = "SearchItem.RowDirection"
args253(2).Value = true
args253(3).Name = "SearchItem.AllTables"
args253(3).Value = false
args253(4).Name = "SearchItem.Backward"
args253(4).Value = false
args253(5).Name = "SearchItem.Pattern"
args253(5).Value = false
args253(6).Name = "SearchItem.Content"
args253(6).Value = false
args253(7).Name = "SearchItem.AsianOptions"
args253(7).Value = false
args253(8).Name = "SearchItem.AlgorithmType"
args253(8).Value = 0
args253(9).Name = "SearchItem.SearchFlags"
args253(9).Value = 65536
args253(10).Name = "SearchItem.SearchString"
args253(10).Value = "kontinu"
args253(11).Name = "SearchItem.ReplaceString"
args253(11).Value = "dur"
args253(12).Name = "SearchItem.Locale"
args253(12).Value = 255
args253(13).Name = "SearchItem.ChangedChars"
args253(13).Value = 2
args253(14).Name = "SearchItem.DeletedChars"
args253(14).Value = 2
args253(15).Name = "SearchItem.InsertedChars"
args253(15).Value = 2
args253(16).Name = "SearchItem.TransliterateFlags"
args253(16).Value = 1280
args253(17).Name = "SearchItem.Command"
args253(17).Value = 3
args253(18).Name = "Quiet"
args253(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args253())


end sub
