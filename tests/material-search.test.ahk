#Requires AutoHotkey v2.0 32-bit
#NoTrayIcon
#Include %A_ScriptDir%\..\Lib\SapMaterialQuery.ahk
cases := [
    ["차량용배터리,DC200,12V", "배터리 12V", true],
    ["차량용배터리,DC200,24V", "배터리 12V", false],
    ["전원공급기,12V", "배터리 12V", false],
    ["차량용배터리,DC200,12V", "12v 배터리", true],
    ["차량용배터리,DC200,12V", "  배터리`t 12V  ", true],
    ["차량용배터리,DC200,12V", "배터리 배터리", true],
    ["차량용배터리,DC200,12V", "", true],
    ["차량용배터리,DC200,12V", " `t ", true],
    ["배터리[12V]", "배터리 [12V]", true],
    ["배터리12V", "배터리 [12V]", false]
]
for testCase in cases
    if SapMaterialMatches(testCase[1], testCase[2]) != testCase[3]
        ExitApp(1)
FileAppend("PASS: AND search, reversed order, case, whitespace, duplicate words, empty query, literal symbols`n", "*")
ExitApp(0)
