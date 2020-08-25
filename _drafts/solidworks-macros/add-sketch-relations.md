---
categories: Solidworks-macro
title:  Solidworks Macro - Add Sketch Relations (Constraints) From VBA Macro
image:  post-image.jpg
tags:   [Solidworks Macro]
---

In this post, I tell you about **how to Add Sketch Relations (Constraints) using Solidworks VBA Macros** in a Sketch.

In this post, I explain about `SketchAddConstraints` method from **Solidworks**'s `ModelDoc2` object.

This method is ***most updated*** method, I found in *Solidworks API Help*.

This post will utilize the methods explained in earlier posts, hence knowledge to those is required but it is not mandatory.

An absolute beginner can follow what is written here.

---

## Content

This post is divided into below sections:

  - *[Code Sample](#code-sample)*
  
  - *[Understanding the Code](#understanding-the-code)*

Feel free to select the section you want to go!

---

## Add Sketch Relations (Constraints) method

For adding relations to a sketch segment, we use `SketchAddConstraints` method from **Solidworks**'s `ModelDoc2` object.

This `SketchAddConstraints` method takes following parameters as explained:

  - **Constraint** : *ID of constraint as given on **[this page](http://help.solidworks.com/2020/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDoc2~SketchAddConstraints.html?verRedirect=1)**.*

**Return Value**: There are no return value for this method.

In following sections we add different sketch constraints to sketch segments.

---

## Add Fix Sketch Relation to a sketch segment

Here we learn how to add `Fixed` *sketch relation* to a sketch segment through **VBA**.

We need *an unconstraint sketch segment*.

In this post, I use a `circle` as shown in below image:

**Before Add *Fix* Sketch Relation to Circle**

![circle-before-fixed-relation](/assets/Solidworks_Images/sketch-relations/circle-before-fixed-relation.png)

**Code to add `Fix` sketch relation**

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
  
  ' Set Solidworks document to active part document
  Set swDoc = swApp.ActiveDoc
  
  ' Select Circle
  BoolStatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
  
  ' Add Fix sketch relation
  swDoc.SketchAddConstraints ("sgFIXED")
  
  ' Clear selection after adding relation
  swDoc.ClearSelection2 True

End Sub
```

**After Add *Fix* Sketch Relation to Circle**

![circle-after-fixed-relation](/assets/Solidworks_Images/sketch-relations/circle-after-fixed-relation.png)

I have added comments to each line `code sample`, hence it is easy to understand.

---

## Add Coincident Sketch Relation to a sketch segment

Here we learn how to add `Coincident` *sketch relation* to a sketch segment through **VBA**.

We need *an unconstraint sketch segment*.

In this post, I use a `circle` as shown in below image:

**Before Add *Coincident* Sketch Relation to Circle**

![circle-before-coincident-relation](/assets/Solidworks_Images/sketch-relations/circle-before-coincident-relation.png)

**Code to add `Coincident` sketch relation**

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
  
  ' Select Circle center point
  BoolStatus = swDoc.Extension.SelectByID2("Point2", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
  
  ' Select Origin
  BoolStatus = swDoc.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
  
  ' Add Coincident sketch relation
  swDoc.SketchAddConstraints ("sgCOINCIDENT")
  
  ' Clear selection after adding relation
  swDoc.ClearSelection2 True

End Sub
```

**After Add *Coincident* Sketch Relation to Circle**

![circle-after-Coincident-relation](/assets/Solidworks_Images/sketch-relations/circle-after-Coincident-relation.png)

I have added comments to each line `code sample`, hence it is easy to understand.