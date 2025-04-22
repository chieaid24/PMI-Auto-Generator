#Requires AutoHotkey v2.0
#SingleInstance Force
CoordMode "Pixel", "Screen"
CoordMode "Mouse", "Screen"

; INCLUDE:
; Popups for instructions during each step

    text := "ⓘ Open up completed bubble drawing in MBDVidia, enter default tolerances, and press SHIFT + ENTER when ready `n (make sure the display with MBDVidia is 1920x1080p, don't touch mouse or keyboard unless prompted, Alt + X to exit)"
    ; Create GUI
    myGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    myGui.BackColor := "White"
    myGui.TransColor := "White"
    myGui.SetFont("s10 Bold", "Calibri")
    txtCtrl := myGui.Add("Text", "c000000 w700 Center", text)

    ; Show GUI

    x := 2
    y := A_ScreenHeight - 120
    myGui.Show("x" x " y" y " NoActivate")

    ; Make it click-through
    hwnd := myGui.Hwnd
    WS_EX_LAYERED := 0x80000
    WS_EX_TRANSPARENT := 0x20
    DllCall("SetWindowLongPtr", "Ptr", hwnd, "Int", -20, "Ptr", WS_EX_LAYERED | WS_EX_TRANSPARENT)


~+Enter:: {
    global myGui
    WinWait("MBDVidia")
    WinActivate("MBDVidia")
    myGui.Hide()
    WinWaitActive("MBDVidia")
    Sleep(300)
    Send("{Ctrl down}w{Ctrl up}")
    Sleep(500)

    WinGetPos(&x, &y, &width, &height, "Select Reports")
    Sleep(500)
    if (ImageSearch(&X2, &Y2, x, y, (x+ width - 1), (y + height - 1), "*100 images/basic_only.png")) {
        if (ImageSearch(&X5, &Y5, x, y, (x + width - 1), (y + height - 1), "*50 images/export.png")) {
            MouseMove(X5 + 5, Y5 + 5)
            Sleep(200)
            Click()
        }
    }

    else {
        Loop
            {
                ; Get the position and size of the active window
                ; Find the blue check mark image within the active window
                if(!(ImageSearch(&X1, &Y1, x, y, (x + width - 1), (y + height - 1), "*100 images/blue_3.png"))) {
                    break
                }
        
                ; Click the blue check mark
                Click(X1 + 4, Y1 + 4)
                Sleep(100)
            }
        
            Click(x + 35, y + 65)
            Sleep(800)

            if (ImageSearch(&X5, &Y5, x, y, (x + width - 1), (y + height - 1), "*50 images/export.png")) {
                MouseMove(X5 + 5, Y5 + 5)
                Sleep(1000)
                Click()
            }
            else {
                Send("{Tab}")
                Sleep(800)
                Send("{Enter}")
            }
    }

    txtCtrl.Text := "ⓘ Save into the 'INPUT FILES' folder located next to this script's file `n*If the Excel doc is not being saved, hit Ctrl + W and export with just the 'Basic' report checked"
    myGui.Show()
    Loop {
        WinGetPos(&x, &y, &width, &height, "MBDVidia") 
        if (ImageSearch(&X3, &Y3, x, y, (x + width - 1), (y + height - 1), "*100 images/report_finished.png")) {
            break
        }
        Sleep(300)
    }
    myGui.Hide()
    Sleep(300)
    WinActivate("MBDVidia")
    Send("{Ctrl down}o{Ctrl up}")
    txtCtrl.Text := "ⓘ Open up the part's corresponding .STP file"
    myGui.Show()

    ; wait until the stp file is opened up
    stpFound := false
    Loop {
        if (stpFound) {
            break
        }
        for hwnd in WinGetList() {
            title := WinGetTitle(hwnd)
            if (RegExMatch(title, "i)(MBDVidia.*\.stp|\.stp.*MBDVidia)") || RegExMatch(title, "i)(MBDVidia.*\.STEP|.STEP.*MBDVidia)")) {
                stpFound := true
                break
            }
        }
    }
    myGui.Hide()
    Sleep(2000)
    WinActivate("MBDVidia")
    Send("^+s")
    txtCtrl.Text := "ⓘ Again, save into the 'INPUT FILES' folder located next to this script's file`n *Ctrl + Shift + S to manually Save As"
    myGui.Show()

    ; Wait until the Document is saved correctly
    Loop {
        if WinExist("MBDVidia") {
            WinGetPos(&x, &y, &width, &height, "MBDVidia") 
            if (ImageSearch(&X3, &Y3, x, y, (x + width - 1), (y + height - 1), "*100 images/document_saved.png")) {
                break
            }
        }
        Sleep(300)
    }
    myGui.Hide()
    Sleep(100)
    txtCtrl.Text := "ⓘ The annotated QIF will be opened shortly! The program is finished."
    myGui.Show()
    Sleep(1000)
    ExitApp()
}

!x:: {
    ExitApp()
}

; o:: {
;     Send("{Ctrl down}w{Ctrl up}")
;     Sleep(500)

;     WinGetPos(&x, &y, &width, &height, "Select Reports")  ; "A" refers to the active window
;     Sleep(500)
;     MouseGetPos(&xm, &ym)
;     ; MouseMove()
;     if (ImageSearch(&X2, &Y2, x, y, (x+ width - 1), (y + height - 1), "*100 images/basic_only.png")) {
;         if (ImageSearch(&X5, &Y5, x, y, (x + width - 1), (y + height - 1), "*50 images/export.png")) {
;             MouseMove(X5 + 5, Y5 + 5)
;             Sleep(200)
;             Click()
;         }
;     }

;     else {
;         Loop
;             {
;                 ; Get the position and size of the active window
;                 ; Find the blue check mark image within the active window
;                 if(!(ImageSearch(&X1, &Y1, x, y, (x + width - 1), (y + height - 1), "*100 images/blue_3.png"))) {
;                     break
;                 }
        
;                 ; Click the blue check mark
;                 Click(X1 + 4, Y1 + 4)
;                 Sleep(100)
;             }
        
;             Click(x + 35, y + 65)
;             Sleep(800)

;             if (ImageSearch(&X5, &Y5, x, y, (x + width - 1), (y + height - 1), "*50 images/export.png")) {
;                 MouseMove(X5 + 5, Y5 + 5)
;                 Sleep(1000)
;                 Click()
;             }
;             else {
;                 Send("{Tab}")
;                 Sleep(800)
;                 Send("{Enter}")
;             }
;     }

; }