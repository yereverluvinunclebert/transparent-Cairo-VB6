# Transparent-Cairo-VB6

A test project using non-Windows-traditional methods of painting PNGs to VB6 picboxes using a transparent VB6 form and a wrapper around Cairo to paint a PNG.

This program runs using VB6 or TwinBasic.

When run in VB6, the two image boxes are loaded using JPGs but the result is blocky whilst the background is a solid fill. Note that VB6 imageboxes cannot handle PNGs at all. The picbox displays the PNG correctly when painted with Cairo.

When run in TwinBasic, the two image boxes are loaded using PNGs and the result is a great improvement with the PNGs nicely layered on top of each other, the background being completely transparent. The picbox displays the PNG perfectly when painted with Cairo.

There are some artifacts left on the screen using both language variants. This is due to neither environment having native capability to display a transparent form. Instead we have to use a series of Windows APIs to change the form layer attributes to add a vbCyan colour mask to make the coloured elements disappear. This process does not work with alpha blended areas of PNGs such as shadows or 'glows' leaving them with a distinct cyan hue. 

This is not a mistake of TwinBasic but a hangover from the inadequate method of making a form transparent that windows provides.

Another artifact left over is a single thin line that marks the boundary edge of the form. This only shows in the VB6 version.

The TwinBasic version handles transparencies very well.
