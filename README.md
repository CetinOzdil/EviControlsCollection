# Evi Controls Collection for VB6 1.2
Revised and updated version of Visual Basic 6 Evi Controls Collection by Evi Indra Effendi <effendi24@gmail.com> at [Github](https://github.com/Planet-Source-Code/evi-indra-effendi-evi-collection-control-xp-update-v1-1__1-69723)

![Different OSes](https://github.com/CetinOzdil/EviControlsCollection/blob/main/EviStyles.png)

_Only made changes on EviButton, rest of controls are untouched from version 1.1_

#### Changes 
* Fixed a GDI leak which lead to a crash
* Added Windows Vista/7 and Windows 8/10 button styles
* Added option to adapt OS button style automatically (`ButtonStyleOS`)
* Made changes to `DefaultColors` behavior to preserve custom colors
* Made changes to `DefaultColors`and `BackColor` behavior to auto generate `ColorHover` and `ColorPressed` automatically
* Removed fancy XP style tooltip, which was causing owner form to loose focus (Classic tooltip is still working)
* Removed `Bevel` and `BevelDepth` options, which feels out of place

#### Notes
* I didn't bother to implement transition animations of new Windows systems
* `InitCommonControlsEx` caused crashes on Windows 7 systems when exiting from application. I found this mush simpler snippet which works ok so far

``` vba
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub InitCommonControlsVB()
    IsUserAnAdmin
    InitCommonControls
End Sub

Public Sub Main()
    InitCommonControlsVB
    ' Rest of startup
End Sub

```
