# Transparent-Cairo-VB6

A test project using non-Windows-traditional methods of painting PNGs to VB6 picboxes using a transparent VB6 form and a wrapper around Cairo to paint a PNG.
This program is a test for replacing RichClient PNG graphic support with an alternative Cairo drop-in replacement and also testing improved image handling within TwinBasic native image controls.

Note that the form cannot be dragged in the traditional way having no title bar, so instead you click on an image to drag it to a new location. 

This program runs using VB6 or TwinBasic.

When run in VB6, the two image boxes on the right are loaded using JPGs but the result is blocky whilst the background is a solid fill. Note that VB6 imageboxes cannot handle PNGs at all. The picbox on the left displays the PNG correctly when painted with Cairo giving VB6 an ability it previously did not have.

![picBoTrans002](https://github.com/user-attachments/assets/bb179608-3331-4488-aa6f-09b1f69198d9)

When run in TwinBasic, the two image boxes are loaded using PNGs as TwinBasic has PNG support out of the box and the result is a great improvement with the images nicely layered on top of each other, the background being completely transparent. In addition the picbox on the left displays the PNG perfectly when painted with Cairo. Note the imageboxes do not have a .hDC so we cannot use Cairo to write a PNG to an imagebox, instead we use TwinBasic's native PNG support.

![picBoTrans001](https://github.com/user-attachments/assets/24e2f986-9514-4486-9e0c-8542b8a0a57f)

There are some artifacts left on the screen using both language variants. This is due to neither environment having native capability to display a transparent form. Instead, we have to use a series of Windows APIs to change the form layer attributes to add a vbCyan colour mask to make any Cyan-coloured elements disappear. This process does not work with alpha blended areas of PNGs such as shadows or 'glows' leaving them with a distinct cyan hue. This is not a mistake of TwinBasic but a hangover from the inadequate method of making a form transparent that windows provides using APIs.

Another artifact left over in the VB6 version alone, is a single thin line that marks the boundary edge of the form.

The TwinBasic version handles transparencies very well. There is a problem though with mouse events with both environments that only really make an appearance in TwinBasic. It is related to the imageboxes. In VB6 a click on the white block around an image is clearly associated at the imagebox containing the image. In TwinBasic, although the PNGs are displayed correctly, in areas where images overlap, you might expect a click on a transparent area to clikc-through to the image layered underneath. A click in this area does not respond exactly as you would intuitively expect as the surrounding border is still present even though it is now transparent. Any click is taken by the layer above, the 'topmost' control. 

A double-click on each image will cause a msgbox to appear with the clicked-control named so it will become clear which transparent area corresponds to which control.

The problems are as follows:

1. The form though transparent does not allow click-through to underlying elements.
2. The form can be made to allow click-through to underlying elements but when this flag is implemented all the controls on the form become click-though too.
3. The imageboxes showing PNGs seem to hijack a click to an underlying image as the transparent areas do not allow click-through to the control below.
4. The CAIRO.DLL used is still a 32bit DLL and so converting from RichClient to this new Cairo implementation will not help much with a 64bit implementation. The code to the Cairo wrapper is available though and in time should be compileable to a 64bit version using TwinBasic.
5. Unfortunately TwinBasic cannot compile the Cairo component successfully yet.
