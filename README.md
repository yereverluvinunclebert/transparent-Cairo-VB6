# Transparent-Cairo-VB6

A test project using non-Windows-traditional methods of painting PNGs to VB6 picboxes using a transparent VB6 form and a wrapper around Cairo to paint a PNG.
This program is a test for replacing RichClient PNG graphic support with an alternative Cairo drop-in replacement and also testing improved image handling within TwinBasic native image controls.

This program runs using VB6 or TwinBasic. It uses the Cairo wrapper DLL and TLB found here https://github.com/VBForumsCommunity/VbCairo

At the moment I am getting the program to operate with a combination of GDI & GDI+ functions (https://github.com/yereverluvinunclebert/SteamyDock) 

![cogs](https://github.com/yereverluvinunclebert/SteamyDock/assets/2788342/ba617c24-0c77-4577-b211-47e1c05a4a5e)

but the idea is to move away from low level Windows specific frameworks and attempt the same using multi-platform capable graphics framework and using TwinBasic's native controls in the hope that soon these will be able to achieve what I am looking for in order to be able to create layered transparent image apps with events, properties and controlled click-through - https://github.com/twinbasic/twinbasic/issues/2167

![picBoTrans002](https://github.com/user-attachments/assets/bb179608-3331-4488-aa6f-09b1f69198d9)

When run in TwinBasic, the two image boxes are loaded using PNGs as TwinBasic has PNG support out of the box and the result is a great improvement with the images nicely layered on top of each other, the background being completely transparent. In addition the picbox on the left displays the PNG perfectly when painted with Cairo. Note the imageboxes do not have a .hDC so we cannot use Cairo to write a PNG to an imagebox, instead we use TwinBasic's native PNG support.

![picBoTrans001](https://github.com/user-attachments/assets/24e2f986-9514-4486-9e0c-8542b8a0a57f)

A double-click on each image will cause a msgbox to appear with the clicked-control named so it will become clear which transparent area corresponds to which control.

1. The CAIRO.DLL used is still a 32bit DLL and so converting from RichClient to this new Cairo implementation will not help much with a 64bit implementation. The code to the Cairo wrapper is available though and in time should be compileable to a 64bit version using TwinBasic.
