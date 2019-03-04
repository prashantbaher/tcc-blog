---
categories: Solidworks-macros
title:  Solidworks Sketch Macros - Create Line 
---

In this post, I tell you about *how to create 2D Line through Solidworks VBA Macros* in a sketch.

For this, I take the example from previous [Solidworks Macros - Open Assembly and Drawing document](/solidworks-macros/select-plane-from-tree) post.

In this post, I tell you about `CreateLine` method from **Solidworks** `SketchManager` object.

This method is ***most updated*** method, I found in Solidworks API Help. 

So ***use this method*** if you want to create a new line.

Below is the `code` sample for creating lines.

```vb
Option Explicit

' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
' Boolean Variable
Dim BoolStatus As Boolean
' Creating variable for Solidworks Sketch Manager
Dim swSketchManager As SldWorks.SketchManager

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
    BoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
    
    ' Setting Sketch manager for our sketch
    Set swSketchManager = swDoc.SketchManager
    
    ' Creating Variable for Solidworks Sketch segment
    Dim mySketchSegment As SketchSegment
    
    ' Inserting a sketch into selected plane
    swSketchManager.InsertSketch True
    
    ' Creating an horizontal line
    Set mySketchSegment = swSketchManager.CreateLine(0, 0, 0, 2, 0, 0)
    
    ' De-select the line after creation
    swDoc.ClearSelection2 True

End Sub
```

Now let us walk through *each line* in the above code, and **understand** the meaning of every line.

```vb
Option Explicit
```

This line forces us to define every variable we are going to use. 

For more information please visit [Solidworks Macros - Open new Part document](/solidworks-macros/open-new-document) post.

```vb
' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
```

In this line, we are creating a variable which we named as `swApp` and the type of this `swApp` variable is `SldWorks.SldWorks`.

<!-- Amazon ad for audible -->
{%- include amazon-audible-promotion.html -%}

```vb
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
```

In this line, we are creating a variable which we named as `swDoc` and the type of this `swDoc` variable is `SldWorks.ModelDoc2`.

Next is our `Sub` procedure named `main`. This procedure hold all the *statements (instructions)* we give to computer.

```vb
' Setting Solidworks variable to Solidworks application
Set swApp = Application.SldWorks
```

In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application.

```vb
' Creating string type variable for storing default part location
Dim defaultTemplate As String
' Setting value of this string type variable to "Default part template"
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
```

In 1st statement of above example, we are defining a variable of `string` type and named it as `defaultTemplate`.

This variable `defaultTemplate`, hold the location the location of **Default Part Template**.

In 2nd line of above example. we assign value to our newly define `defaultTemplate` variable.

We assign the value by using a *Method* named `GetUserPreferenceStringValue()`. This method is a part of our main Solidworks variable `swApp`.

