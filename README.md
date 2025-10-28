# Transparent-Cairo-VB6

A test project using non-Windows-traditional methods of painting PNGs to VB6 picboxes using a transparent VB6 form and a wrapper around Cairo to paint a PNG.

This program runs using VB6 or TwinBasic.

When run in VB6, the two image boxes are loaded using JPGs but the result is blocky whilst the background is a solid fill. Note that VB6 imageboxes cannot handle PNGs at all. The picbox displays the PNG correctly when painted with Cairo.

![picBoTrans002](https://github.com/user-attachments/assets/bb179608-3331-4488-aa6f-09b1f69198d9)

When run in TwinBasic, the two image boxes are loaded using PNGs and the result is a great improvement with the PNGs nicely layered on top of each other, the background being completely transparent. The picbox displays the PNG perfectly when painted with Cairo.

![picBoTrans001](https://github.com/user-attachments/assets/24e2f986-9514-4486-9e0c-8542b8a0a57f)


There are some artifacts left on the screen using both language variants. This is due to neither environment having native capability to display a transparent form. Instead we have to use a series of Windows APIs to change the form layer attributes to add a vbCyan colour mask to make the coloured elements disappear. This process does not work with alpha blended areas of PNGs such as shadows or 'glows' leaving them with a distinct cyan hue. 

This is not a mistake of TwinBasic but a hangover from the inadequate method of making a form transparent that windows provides.

Another artifact left over is a single thin line that marks the boundary edge of the form. This only shows in the VB6 version.

The TwinBasic version handles transparencies very well. There is a problem though with mouse events. Although the PNGs are displayed correctly in the TwinBasic version, in areas where images overlap, a click on a transparent area around the upper image layer does not respond exactly as you would expect, instead reverting to the layer above, the 'topmost' control. A double-click on each image and the transparent areas that surround them will show which control responds.

Note that the form cannot be dragged in the traditional way (no title bar) so instead you click on an image to drag it to a new location. 

This program is a test for replacing RichClient PNG graphic support with an alternative Cairo drop in replacement and TwinBasic native image controls.

The problems are that the CAIRO.DLL used is still a 32bit DLL and so converting from RichClient to this new Cairo implementation will not help much with a 64bit implementation. The code to the Cairo wrapper is available though and in time should be compileable to a 64bit version using TwinBasic. Unfortunately TwinBasic cannot compile the Cairo component successfully yet.
