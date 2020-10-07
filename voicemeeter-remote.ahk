#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force

WinWait, ahk_exe voicemeeterpro.exe  ; wait for voicemeeter

DllLoad := DllCall("LoadLibrary", "Str", "C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll")

VMLogin()
OnExit("VMLogout")

; set initial state
global headphone_out = 1
Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Bus[0].Gain", "Float", 3.0)
Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Bus[1].Gain", "Float", 0.0)
Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[3].A1", "Float", 1.0)
Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[3].A2", "Float", 0.0)

VMLogin() {
    Login := DllCall("VoicemeeterRemote64\VBVMR_Login")
}

VMLogout() {
    Logout := DllCall("VoicemeeterRemote64\VBVMR_Logout")
}

SetTimer, GetStatu, 100

ApplyVolume(vol_lvl) {
    if (vol_lvl > 12.0){
        vol_lvl = 12.0
    } else if (vol_lvl < -60.0) {
        vol_lvl = -60.0
    }
    Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Bus[0].Gain", "Float", vol_lvl)
    Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Bus[1].Gain", "Float", vol_lvl)
    DllCall("VoicemeeterRemote64\VBVMR_IsParametersDirty")
}



; redirect mic 1 to main output, mute mic 2
CapsLock::
    Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[0].B1", "Float", 0.1)
    Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[0].Mute", "Float", 0.0)
    
return

CapsLock Up::
    Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[0].B1", "Float", 0.0)
    Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[0].Mute", "Float", 1.0)
return


Volume_Up::
    DllCall("VoicemeeterRemote64\VBVMR_IsParametersDirty")
    Result := DllCall("VoicemeeterRemote64\VBVMR_GetParameterFloat", "AStr", "Bus[0].Gain", "Ptr", &vol_lvl)
    vol_lvl := NumGet(vol_lvl, 0, "Float")
    OutputDebug, %vol_lvl%
    vol_lvl += 1.0
    ApplyVolume(vol_lvl)
return

Volume_Down::
    DllCall("VoicemeeterRemote64\VBVMR_IsParametersDirty")
    Result := DllCall("VoicemeeterRemote64\VBVMR_GetParameterFloat", "AStr", "Bus[0].Gain", "Ptr", &vol_lvl)
    
    vol_lvl := NumGet(vol_lvl, 0, "Float")
    vol_lvl -= 1.0
    ApplyVolume(vol_lvl)
return


GetStatu(){
    DllCall("VoicemeeterRemote64\VBVMR_IsParametersDirty")
    Result := DllCall("VoicemeeterRemote64\VBVMR_GetParameterFloat", "AStr", "Strip[0].Mute", "Ptr", &vol_lv2)
    vol_lv2 := NumGet(vol_lv2, 0, "Float")
    DllCall("VoicemeeterRemote64\VBVMR_IsParametersDirty")
    vol_lv2 += 0.0
    OutputDebug, %vol_lv2%
    if (vol_lv2 == 1.0){
        TaskBar_SetAttr(0) 
    } else {
        TaskBar_SetAttr(2, 0xff0f0f91)   
    }
    return
}



; switch audio devices
XButton2::
    if (headphone_out = 1){
        Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[3].A1", "Float", 0.0)
        Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[3].A2", "Float", 1.0)
        headphone_out = 0
    } else {
        Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[3].A1", "Float", 1.0)
        Result := DllCall("VoicemeeterRemote64\VBVMR_SetParameterFloat", "AStr", "Strip[3].A2", "Float", 0.0)
        headphone_out = 1
    }
return

TaskBar_SetAttr(accent_state := 0, gradient_color := "0x01000000")
{
    static init, hTrayWnd, ver := DllCall("GetVersion") & 0xff < 10
    static pad := A_PtrSize = 8 ? 4 : 0, WCA_ACCENT_POLICY := 19

    if !(init) {
        if (ver)
            throw Exception("Minimum support client: Windows 10", -1)
        if !(hTrayWnd := DllCall("user32\FindWindow", "str", "Shell_TrayWnd", "ptr", 0, "ptr"))
            throw Exception("Failed to get the handle", -1)
        init := 1
    }

    accent_size := VarSetCapacity(ACCENT_POLICY, 16, 0)
    NumPut((accent_state > 0 && accent_state < 4) ? accent_state : 0, ACCENT_POLICY, 0, "int")

    if (accent_state >= 1) && (accent_state <= 2) && (RegExMatch(gradient_color, "0x[[:xdigit:]]{8}"))
        NumPut(gradient_color, ACCENT_POLICY, 8, "int")

    VarSetCapacity(WINCOMPATTRDATA, 4 + pad + A_PtrSize + 4 + pad, 0)
    && NumPut(WCA_ACCENT_POLICY, WINCOMPATTRDATA, 0, "int")
    && NumPut(&ACCENT_POLICY, WINCOMPATTRDATA, 4 + pad, "ptr")
    && NumPut(accent_size, WINCOMPATTRDATA, 4 + pad + A_PtrSize, "uint")
    if !(DllCall("user32\SetWindowCompositionAttribute", "ptr", hTrayWnd, "ptr", &WINCOMPATTRDATA))
        throw Exception("Failed to set transparency / blur", -1)
    return true
}

; ===============================================================================================================================

/*
Shell_TrayWnd             -> Main TaskBar
Shell_SecondaryTrayWnd    -> 2nd  TaskBar (on multiple monitors)
*/

/* C++ ==========================================================================================================================
BOOL GetWindowCompositionAttribute(
    _In_    HWND hWnd,
    _Inout_ WINDOWCOMPOSITIONATTRIBDATA* pAttrData
);
BOOL SetWindowCompositionAttribute(
    _In_    HWND hWnd,
    _Inout_ WINDOWCOMPOSITIONATTRIBDATA* pAttrData
);
typedef struct _WINDOWCOMPOSITIONATTRIBDATA {
    WINDOWCOMPOSITIONATTRIB Attrib;
    PVOID                   pvData;
    SIZE_T                  cbData;
} WINDOWCOMPOSITIONATTRIBDATA;
typedef enum _WINDOWCOMPOSITIONATTRIB {
    WCA_UNDEFINED = 0,
    WCA_NCRENDERING_ENABLED = 1,
    WCA_NCRENDERING_ENABLED = 1,
    WCA_NCRENDERING_POLICY = 2,
    WCA_TRANSITIONS_FORCEDISABLED = 3,
    WCA_ALLOW_NCPAINT = 4,
    WCA_CAPTION_BUTTON_BOUNDS = 5,
    WCA_NONCLIENT_RTL_LAYOUT = 6,
    WCA_FORCE_ICONIC_REPRESENTATION = 7,
    WCA_EXTENDED_FRAME_BOUNDS = 8,
    WCA_HAS_ICONIC_BITMAP = 9,
    WCA_THEME_ATTRIBUTES = 10,
    WCA_NCRENDERING_EXILED = 11,
    WCA_NCADORNMENTINFO = 12,
    WCA_EXCLUDED_FROM_LIVEPREVIEW = 13,
    WCA_VIDEO_OVERLAY_ACTIVE = 14,
    WCA_FORCE_ACTIVEWINDOW_APPEARANCE = 15,
    WCA_DISALLOW_PEEK = 16,
    WCA_CLOAK = 17,
    WCA_CLOAKED = 18,
    WCA_ACCENT_POLICY = 19,
    WCA_FREEZE_REPRESENTATION = 20,
    WCA_EVER_UNCLOAKED = 21,
    WCA_VISUAL_OWNER = 22,
    WCA_LAST = 23
} WINDOWCOMPOSITIONATTRIB;
typedef struct _ACCENT_POLICY {
    ACCENT_STATE AccentState;
    DWORD        AccentFlags;
    DWORD        GradientColor;
    DWORD        AnimationId;
} ACCENT_POLICY;
typedef enum _ACCENT_STATE {
    ACCENT_DISABLED = 0,
    ACCENT_ENABLE_GRADIENT = 1,
    ACCENT_ENABLE_TRANSPARENTGRADIENT = 2,
    ACCENT_ENABLE_BLURBEHIND = 3,
    ACCENT_INVALID_STATE = 4
} ACCENT_STATE;
_ACCENT_FLAGS {
    DrawLeftBorder = 0x20,
    DrawTopBorder = 0x40,
    DrawRightBorder = 0x80,
    DrawBottomBorder = 0x100,
    DrawAllBorders = (DrawLeftBorder | DrawTopBorder | DrawRightBorder | DrawBottomBorder)
}
============================================================================================================================== */

