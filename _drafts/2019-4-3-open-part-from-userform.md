---
categories: Solidworks-macros
title:  Solidworks Macros - Open new Part document from VBA Userforms 
---

In this post, we learn how can we *open a part document* from a Visual Basic for Application's *Userform*.

For this please we need to do following steps as described below.

## Create a new macro

1st we need to create a **new macro** in *Solidworks 3D CAD Software*.

If you don't know how to create a new macro in Solidworks, please go to [VBA in Solidworks](/solidworks-macros/vba-in-solidworks) post for this.

This will open a new macro in Visual Basic Editor with some code as shown in below image.

![new_macro_window](/assets/Solidworks_Images/Open new part from userform/1.new_macro_window.png)

## Insert userform in the macro

After this we need to insert a userform in our macro.

For this, select the button shown in below image.

![insert-userform-button](/assets/Solidworks_Images/Open new part from userform/2.insert-userform-buttonw.png)

This button is called ***insert userform***. 

As the name suggest, function of this button is *inserting a userform*.

> Please note that in a macro we can insert any number of userform as we like. But for this example we insert only 1 userform.

After clicking the ***insert userform*** button we get the userform window as shown in below image.

![userform-window](/assets/Solidworks_Images/Open new part from userform/3.userform-window.png)

## Adding a Button

Now in our userform window, we add a `Command Button` at center of window.

You can find `Command Button` highlighted in red in below image.

![add-command-buton](/assets/Solidworks_Images/Open new part from userform/4.add-command-button.png)

You can place command button at your desire. I placed it at center of the userform window as shown in below image.

![userform-window-after-adding-command-button](/assets/Solidworks_Images/Open new part from userform/5.userform-window-after-adding-command-button.png)

## Updating Properties of Command Button and Userform Windows

Now we update some properties of Command Button and Userform Windows for our use.

> It is not necessary to update properties but it is a good habit to update them for our purpose. 




