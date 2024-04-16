Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub MasterPartsListReset(control As IRibbonControl)
'Created for Pason by Dragon Wood (August 2015)
'Resets the Master Parts List page so it can be used again.

    Application.ScreenUpdating = False
    
    'Clear all data on the page.
    Call MasterPartsListClear
    
    Application.ScreenUpdating = True
    Sheets("System Selection").Select

End Sub

Function MasterPartsListClear()
'Created for Pason by Dragon Wood (August 2015)
'Resets the Master Parts List Page to the original unused state.

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Reset the Yes/No field to No on the Master Parts List Sheet and clear the Quantity Needed Column
    With ActiveWorkbook.Worksheets("Master Parts List")
        .Range("D3:D1958").ClearContents
        .Range("A3").Value = strNo
        .Range("A5:A9").Value = strNo
        .Range("A11").Value = strNo
        .Range("A13:A29").Value = strNo
        .Range("A31:A34").Value = strNo
        .Range("A36:A49").Value = strNo
        .Range("A51:A88").Value = strNo
        .Range("A90:A98").Value = strNo
        .Range("A100:A118").Value = strNo
        .Range("A120").Value = strNo
        .Range("A122:A134").Value = strNo
        .Range("A136:A154").Value = strNo
        .Range("A156:A166").Value = strNo
        .Range("A168:A172").Value = strNo
        .Range("A174:A181").Value = strNo
        .Range("A183:A205").Value = strNo
        .Range("A207:A208").Value = strNo
        .Range("A210:A241").Value = strNo
        .Range("A243:A279").Value = strNo
        .Range("A281:A283").Value = strNo
        .Range("A290:A304").Value = strNo
        .Range("A306:A333").Value = strNo
        .Range("A335:A337").Value = strNo
        .Range("A339:A401").Value = strNo
        .Range("A403:A439").Value = strNo
        .Range("A441:A451").Value = strNo
        .Range("A453").Value = strNo
        .Range("A455:A463").Value = strNo
        .Range("A465:A471").Value = strNo
        .Range("A473:A544").Value = strNo
        .Range("A546:A556").Value = strNo
        .Range("A558:A572").Value = strNo
        .Range("A574:A620").Value = strNo
        .Range("A622:A634").Value = strNo
        .Range("A636:A649").Value = strNo
        .Range("A651:A652").Value = strNo
        .Range("A654:A664").Value = strNo
        .Range("A666:A687").Value = strNo
        .Range("A689:A713").Value = strNo
        .Range("A715:A723").Value = strNo
        .Range("A725:A738").Value = strNo
        .Range("A740:A764").Value = strNo
        .Range("A766:A789").Value = strNo
        .Range("A791:A793").Value = strNo
        .Range("A795:A821").Value = strNo
        .Range("A823:A916").Value = strNo
        .Range("A918:A988").Value = strNo
        .Range("A990").Value = strNo
        .Range("A992:A994").Value = strNo
        .Range("A996:A1003").Value = strNo
        .Range("A1005:A1222").Value = strNo
        .Range("A1235:A1252").Value = strNo
        .Range("A1254:A1282").Value = strNo
        .Range("A1284:A1297").Value = strNo
        .Range("A1299:A1308").Value = strNo
        .Range("A1310:A1340").Value = strNo
        .Range("A1342:A1348").Value = strNo
        .Range("A1399:A1419").Value = strNo
        .Range("A1421:A1425").Value = strNo
        .Range("A1427:A1448").Value = strNo
        .Range("A1450:A1496").Value = strNo
        .Range("A1498:A1531").Value = strNo
        .Range("A1533:A1550").Value = strNo
        .Range("A1552:A1558").Value = strNo
        .Range("A1560:A1579").Value = strNo
        .Range("A1581:A1590").Value = strNo
        .Range("A1592:A1598").Value = strNo
        .Range("A1600:A1644").Value = strNo
        .Range("A1646:A1711").Value = strNo
        .Range("A1713:A1741").Value = strNo
        .Range("A1743:A1762").Value = strNo
        .Range("A1764:A1767").Value = strNo
        .Range("A1769:A1773").Value = strNo
        .Range("A1775").Value = strNo
        .Range("A1777:A1784").Value = strNo
        .Range("A1786:A1791").Value = strNo
        .Range("A1793:A1801").Value = strNo
        .Range("A1803:A1809").Value = strNo
        .Range("A1811:A1816").Value = strNo
        .Range("A1818:A1831").Value = strNo
        .Range("A1833:A1841").Value = strNo
        .Range("A1843:A1846").Value = strNo
        .Range("A1848:A1854").Value = strNo
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
       
    Application.GoTo Sheets("Master Parts List").Range("A1"), True
    Application.GoTo Sheets("System Selection").Range("A1"), True

End Function