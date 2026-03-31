# Transparent-Cairo-VB6

A test project using non-Windows-traditional methods of painting PNGs to a transparent VB6 form using a wrapper around Cairo/GDI+.
This program is a test for replacing RichClient PNG graphic support with an alternative GDI+/Cairo drop-in replacement whilst also testing improved modern image handling within TwinBasic's native image collections.

This program runs using VB6 or TwinBasic. Eventually it will use the Cairo wrapper DLL and TLB found here https://github.com/VBForumsCommunity/VbCairo

At the moment I am getting the program to operate with a combination of GDI+ functions as per my older program, Steamydock - (https://github.com/yereverluvinunclebert/SteamyDock) 

![cogs](https://github.com/yereverluvinunclebert/SteamyDock/assets/2788342/ba617c24-0c77-4577-b211-47e1c05a4a5e)
Fig01. Showing Steamydock, images placed on a transparent form using GDI+. 

However, the idea is to move away from low level Windows specific graphic frameworks and attempt the same using multi-platform capable graphics framework in the hope that soon these will be able to achieve what I am looking for in order to be able to create layered transparent image apps with events, properties and controlled click-through - https://github.com/twinbasic/twinbasic/issues/2167

<img width="1026" height="900" alt="cpu-XML-GDIP" src="https://github.com/user-attachments/assets/e089bdaa-f48a-4f25-b4a7-2c138339b8b6" />
Fig02. This program displaying PNG images placed from an imagelist onto a transparent form using GDI+. 

Image above shows good progress replicating an existing RichClient program in looks (if not yet functionality), success in placing multiple images onto the desktop using transparent PNGs, the size and location details read from a PSD file via an extracted XML descriptor, similarly extracted images are then read into a dictionary/imagelist, then placed lightning-fast onto the display using GDI+.

Working with GDI+ as a proof of concept until the Cairo elements are working as efficiently as GDI+

<img width="1252" height="699" alt="trinket0001" src="https://github.com/user-attachments/assets/9be0e165-acee-4041-846a-504ae2bcabd4" />
Fig03. This program showing a mouseDown on a layer with transparency, raising and handling events.

The CAIRO.DLL used is still a 32bit DLL and so converting from RichClient to this new Cairo implementation will not help much with a 64bit implementation. The code to the Cairo wrapper is available though and in time should be compileable to a 64bit version using TwinBasic.

Status:

BASIC test program - Puts multiple images on screen using GDI+ with a right click menu on an invisible VB6 form using an image class that encapsulates the properties of each image. 
Each layer image is now identified via hit-testing and a click through any transparent area is passed to the layer below. Event raising for each layer is enabled and events handled.

Currently ironing some bugs and then will be creating a widget class that will provide a 'home' for the images.

 
