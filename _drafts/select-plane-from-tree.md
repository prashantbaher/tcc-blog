---
categories: Solidworks-macros
title:  Solidworks Macros - Select Planes
---

In this post, we select **Default planes** with following methods:

1. By `SelectByID` method

2. By *Traversing* through *Feature Manager Tree* and then select plane

## By SelectByID method

`SelectByID` is the easiest method for selecting Default plane.

I will explain the use of this method in 2 different scenerio as follows:

1. Using this method in the previous example of creating a new document and then select a Plane.

2. Using this method in an open document.

### Using this method in previous example

In the previous 2 posts, we learned how to create *a new part document, an assembly document, and a drawing document*.

Now we use the same code and *extended* it for using selecting planes.

```vb
Option Explicit

' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
' Boolean Variable
Dim BoolStatus As Boolean

' Main function of our VBA program
Sub main()

    ' Setting Solidworks variable to Solidworks application
    Set swApp = Application.SldWorks
    
    ' Creating string type variable for storing default part location
    Dim defaultTemplate As String
    ' Setting value of this string type variable to "Default part template"
    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)

    ' Setting Solidworks document to new part document
    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)

    ' Selecting Front Plane
    BoolStatus = swDoc.SelectByID("Front Plane", "PLANE", 0, 0, 0)

End Sub
```

Above code, 1st create a new part document and then select "**Front Plane**" in VBA macro.

To select the plane, I have to add 2 lines. 1st I create a boolean varible above main function as shown in below code:

```vb
' Boolean Variable
Dim BoolStatus As Boolean
```

Then I use this `BoolStatus` variable to selecting *Front Plane* as shown in below code:

```vb
' Selecting Front Plane
BoolStatus = swDoc.SelectByID("Front Plane", "PLANE", 0, 0, 0)
```

`SelectByID` is takes following parameters:

*Name* : Name of the object or an empty string

*Type* : Type of object in uppercase or an empty string

*X* : X selection location

*Y* : Y selection location

*Z* : Z selection location

**Return Value** - If the item is select then this method returns `True` and otherwise `False`.

Since this method returns `True` or `False`, hence we use boolean variable to perfom this method.

If we want to select **Right Plane** then we just need to replace `"Front Plane"` -> `"Right Plane"` in previous code sample.

Similar for selecting **Top Plane**, we need to replace `"Front Plane"` -> `"Top Plane"` in previous code sample.

### Using this method in an Open document

For using this method in an open document we use differnet code sample.

The code sample is given below:

```vb
Option Explicit

' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
' Boolean Variable
Dim BoolStatus As Boolean


' Main function of our VBA program
Sub main()

    ' Setting Solidworks variable to Solidworks application
    Set swApp = Application.SldWorks
        
    ' Setting Solidworks document to active open document
    Set swDoc = swApp.ActiveDoc
        
    ' Selecting Front Plane
    BoolStatus = swDoc.SelectByID("Front Plane", "PLANE", 0, 0, 0)

End Sub
```

Most of the things in this code sample is already explained [this post](/solidworks-macros/open-new-document) and in previous section of this very post.

In this code I have set the *Solidworks document variable* `swDoc` to active open document.

And then we use similar method to select **"Front Plane"**.

As explained in previous section we can select **Right Plane** and **Top Plane**.

