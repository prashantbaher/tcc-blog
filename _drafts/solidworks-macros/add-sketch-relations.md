---
categories: Solidworks-macro
title:  Solidworks Macro - Add Sketch Relations (Constraints) From VBA Macro
image:  post-image.jpg
tags:   [Solidworks Macro]
---

In this post, I tell you about **how to Add Sketch Relations (Constraints) using Solidworks VBA Macros** in a Sketch.

In this post, I explain about `SketchAddConstraints` method from **Solidworks**'s `ModelDoc2` object.

This method is ***most updated*** method, I found in *Solidworks API Help*.

## Add Fix Sketch Relation to sketch segment

```vb
Option Explicit

' Create variable for Solidworks application
Dim swApp As SldWorks.SldWorks

' Create variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2

' Boolean Variable
Dim BoolStatus As Boolean

Sub main()

  ' Set Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  
  ' Set Solidworks document to new part document
  Set swDoc = swApp.ActiveDoc
  
  ' Select Circle
  BoolStatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
  
  ' Add Fix sketch relation
  swDoc.SketchAddConstraints ("sgFIXED")
  
  ' Clear selection after adding relation
  swDoc.ClearSelection2 True

End Sub
```