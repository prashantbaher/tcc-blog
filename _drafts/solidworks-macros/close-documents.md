---
categories: Solidworks-macro
title:  Solidworks Macro - Close Active Document From VBA Macro
image:  post-image.jpg
tags:   [Solidworks Macro]
---

## Content

This post is divided into below sections:

  - *[Introduction](#introduction)*

  - *[Code Sample](#code-sample)*

  - *[Understanding the Code](#understanding-the-code)*

  - *[Multiple Cases](#cases)*

    - *[Update Arc Radius](#update-arc-radius)*

    - *[Update Arc Angle](#update-arc-angle)*

    - *[Update Number of Instances](#update-number-of-instances)*

    - *[Update Pattern Spacing](#update-pattern-spacing)*

    - *[Update Display Rotation of Pattern](#update-rotation-of-pattern)*

    - *[Update Number of Instances to Delete](#update-number-of-instance-to-delete)*

    - *[Update Display Radius Dimension](#update-display-radius-dimension)*

    - *[Update Display Angle Dimension](#update-display-angle-dimension)*

    - *[Update Display Number of Instances](#update-display-number-of-instance)*

Feel free to select the section you want to go!

---

## Introduction

In this post, I tell you about **how to Close Active Document from Solidworks VBA Macros**.

In this post, I explain about `CloseDoc` method from **Solidworks** `SldWorks` object.

This method is ***most updated*** method, I found in *Solidworks API Help*. 

So ***use this methods*** if you want to *Close Active Document*.

---
<!--
## Video of Code on YouTube

Please see below video how we can *Edit Circular Sketch Pattern* in **Solidworks VBA macro**.

<iframe src="https://www.youtube.com/embed/4pLUprIxXHU" frameborder="0" allowfullscreen></iframe>
<br>

Please note that there is **no explanation** in the video. 

Why we write our code this way is **explained** in this post.

--- 

## For Experience Macro Developer - Edit Circular Sketch Pattern From VBA Macro

If you are an experienced **Solidworks Macro developer**, then you are looking for a specific code sample.

Below is the code for **Edit Circular Sketch Pattern** from **Solidworks VBA Macro**.

```vb
' Boolean Variable
Dim BoolStatus As Boolean

' Edit a Circular Sketch Pattern
BoolStatus = swSketchManager.EditCircularSketchStepAndRepeat(5, 4, 1, 0.75, 0.785, 1.5708, "(3,2)(2,1)", True, True, False, False, True, "Arc1_")
```

**Method Name**: `EditCircularSketchStepAndRepeat`

**Description**: Edit *Circular Sketch Pattern*.

**Prerequisites**: To *edit* a **Circular Sketch Pattern** a Solidworks Sketch entity or entities, first, we need the following things:

  1. Existing Circular Sketch Pattern

**How it works**:

  - For **Edit a Circular Sketch Pattern**, first, we need to **create** a variable of `Boolean` type.

  - After creating variable, we need to set the value of this `Boolean` variable.

  - For this, we used `EditCircularSketchStepAndRepeat` method from **Solidworks Sketch Manager**.

  - This `EditCircularSketchStepAndRepeat` method set the value of `Boolean` type variable.

  - If the editing of *Circular Sketch Pattern* is **successful** then `EditCircularSketchStepAndRepeat` method return **True** value otherwise `EditCircularSketchStepAndRepeat` returns **False** value.

This `EditCircularSketchStepAndRepeat` method takes the following parameters as explained:

  - **NumX**: *Total number of instances along the **x** axis, including the seed i.e. original entity/entities.*

  - **NumY**: *Total number of instances along the **y** axis, including the seed i.e. original entity/entities.*

  - **SpacingX**: *Spacing between instances along the **x** axis.*

  - **SpacingY**: *Spacing between instances along the **y** axis.*

  - **AngleX**: *Angle for direction 1 relative to the **x** axis.*

  - **AngleY**: *Angle for direction 1 relative to the **y** axis.*

  - **DeleteInstances**: *Number of instances to delete, passed as a string in the format: "(a) (b) (c)".*

  - **XSpacingDim**: *True to display the spacing between instances dimension along the **x** axis in the graphics area, false to not*

  - **YSpacingDim**: *True to display the spacing between instances dimension along the **y** axis in the graphics area, false to not*
  
  - **AngleDim**: *True to display the angle dimension between axes in the graphics area, false to not.*

  - **CreateNumOfInstancesDimInXDir**: *True to display the number of instances in the **x** direction dimension in the graphics area, false to not.*

  - **CreateNumOfInstancesDimInYDir**: *True to display the number of instances in the **y** direction dimension in the graphics area, false to not.*

  - **Seed**: *List of the names of the entities, separated by the underscore character (_), that comprise the seed pattern (e.g., Line1_Line2_Line3_Line4 for a rectangular-shaped seed pattern).*

**Return Value**:

  - **True**: *If Editing of Circular Sketch Pattern is "Success".*

  - **False**: *If Editing of Circular Sketch Pattern is "Fail".*

If you want more detailed explaination then please read further otherwise this will help you to *edit* a **Circular Sketch Pattern From VBA Macro**.

--- -->

## Code Sample

Below is the `code` sample to *Close Document active document* from a number of opened documents.

```vb
Option Explicit

' Create variable for Solidworks application
Dim swApp As SldWorks.SldWorks

' Create variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2

' Boolean Variable
Dim BoolStatus As Boolean

' Main function of our VBA program
Sub main()

  ' Set Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  
  ' Create string type variable for storing default part location
  Dim defaultTemplate As String

  ' Set value of this string type variable to "Default part template"
  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)

  ' Local variable to define number of documents we want to open
  Dim numberOfDocuments As Integer
  
  ' Create a loop for creating 3 new part documents
  For numberOfDocuments = 1 To 3
    ' Set Solidworks document to new part document
    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
  Next
  
  ' Set solidworks document to currently active document
  Set swDoc = swApp.ActiveDoc

  ' close active document only
  swApp.CloseDoc (swDoc.GetPathName)
  
End Sub
```

---

### Understanding the Code

Now let us walk through **each line** in the above code, and **understand** the meaning and purpose of every line.

I also give some link so that you can go through them if there are anything I explained in **previous posts**.

```vb
Option Explicit
```

This line forces us to define every variable we are going to use. 

For more information please visit [Solidworks Macros - Open new Part document](/solidworks-macro/open-new-document) post.

```vb
' Create variable for Solidworks application
Dim swApp As SldWorks.SldWorks
```

In this line, we create a variable which we named as `swApp` and the type of this `swApp` variable is `SldWorks.SldWorks`.

```vb
' Create variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
```

In this line, we create a variable which we named as `swDoc` and the type of this `swDoc` variable is `SldWorks.ModelDoc2`.

```vb
' Boolean Variable
Dim BoolStatus As Boolean
```

In this line, we create a variable named `BoolStatus` as `Boolean` object type.

These all are our global variables.

As you can see in code sample, they are **Solidworks API Objects**.

So basically I group all the **Solidworks API Objects** in one place.

I have also place `boolean` type object at top also, because after certain point we will *need* this variable frequently.

Thus, I have started placing it here.

Next is our `Sub` procedure which has name of `main`. 

This procedure hold all the *statements (instructions)* we give to computer.

```vb
' Set Solidworks variable to Solidworks application
Set swApp = Application.SldWorks
```

In this line, we set the value of our Solidworks variable `swApp`; which we define earlier; to Solidworks application.

```vb
' Create string type variable for storing default part location
Dim defaultTemplate As String
' Set value of this string type variable to "Default part template"
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
```

In 1st statement of above example, we are defining a variable of `string` type and named it as `defaultTemplate`.

This variable `defaultTemplate`, hold the location the location of **Default Part Template**.

In 2nd line of above example. we assign value to our newly define `defaultTemplate` variable.

We assign the value by using a *Method* named `GetUserPreferenceStringValue()`. 

This `GetUserPreferenceStringValue()` method is a part of our main Solidworks variable `swApp`.

```vb
' Local variable to define number of documents we want to open
Dim numberOfDocuments As Integer
```

In above line we create a local variable of `Integer` type to *define number of documents we want to open* in Solidworks.

```vb
' Create a loop for creating 3 new part documents
For numberOfDocuments = 1 To 3
Next
```

In above line we create a loop for creating 3 new part documents in Solidworks.

Inside the loop we add below code line (statement).

```vb
' Set Solidworks document to new part document
Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
```

In this line, we set the value of our `swDoc` variable to new document.

For **detailed information** about these lines please visit [Solidworks Macros - Open new Part document](/solidworks-macro/open-new-document) post.

I have discussed them **thoroghly** in [Solidworks Macros - Open new Part document](/solidworks-macro/open-new-document) post, so do checkout that post if you want to understand above code in more detail.

```vb
' Set solidworks document to currently active document
Set swDoc = swApp.ActiveDoc
```

In above code line we again set the value of `swDoc` variable to currently active document in Solidworks.

```vb
' close active document only by passing active document full path
swApp.CloseDoc (swDoc.GetPathName)
```

For "**close active document**" a Circular Sketch pattern, we need `CloseDoc` method from *Solidworks* object/variable.

This `CloseDoc` method takes following parameters as explained:

  - **Name** : *Name of the document we want to close.*

After the function complete following are the results:

**Return Value**:

  This method did not return anything.

In above code, I pass the `swDoc.GetPathName` method **directly** as parameter.

`GetPathName` method is part of *Solidworks Document* object/variable.

This `GetPathName` method above return full path of active document.

Then we pass this path into `CloseDoc` method as parameter.

*By doing this we shorten the code.*

Below image shows before and after we update **Arc Radius**.

**Before Closing the Active documents**

![total-number-of-documents-open](/assets/Solidworks_Images/close-document/total-number-of-documents-open.png)

**After Closing the Active documents**

![after-update-arc-radius](/assets/Solidworks_Images/close-document/after-update-arc-radius.png)

---

### **Cases**

In this section, we will go through different cases by 

  - *Modifying different parameters*

  - *See images, before and after parameter modification*

---
  
#### CASE 1 : Update Arc Radius

In our code, if we want to update Arc Radius, then we need to update `arcRadius` variable only.

```vb
' Update Arc Radius
arcRadius = 20 * LengthConversionFactor
```

In above line we **Update Arc Radius** to new value. of 20 inch.

***Example Images:***

Below image shows before and after we update **Arc Radius**.

**Before Update Arc Radius**

![before-edit-circular-pattern](/assets/Solidworks_Images/sketch-patterns/before-edit-circular-pattern.png)

**After Update Arc Radius**

![after-update-arc-radius](/assets/Solidworks_Images/sketch-patterns/after-update-arc-radius.png)

#### CASE 2 : Update Arc Angle

In our code, if we want to update Arc Angle, then we need to update `arcAngle` variable only.

```vb
' Update Arc Angle
arcAngle = 30 * AngleConversionFactor
```

In above line we **Update Arc Angle** to new value of 30 inch.

***Example Images:***

Below image shows before and after we update **Arc Angle**.

**Before Update Arc Angle**

![after-update-arc-radius](/assets/Solidworks_Images/sketch-patterns/after-update-arc-radius.png)

**After Update Arc Angle**

![after-update-arc-angle](/assets/Solidworks_Images/sketch-patterns/after-update-arc-angle.png)

#### CASE 3 : Update Number of Instances

In our code, if we want to update Number of Instances, then we need to update `numberOfInstance` variable only.

```vb
' Update Number of Instances
numberOfInstance = 5
```

In above line we **Update Number of Instances** to new value of 5 number of instances.

***Example Images:***

Below image shows before and after we update **Number of Instances**.

**Before Update Number of Instances**

![after-update-arc-angle](/assets/Solidworks_Images/sketch-patterns/after-update-arc-angle.png)

**After Update Number of Instances**

![after-update-number-of-instances](/assets/Solidworks_Images/sketch-patterns/after-update-number-of-instances.png)

#### CASE 4 : Update Pattern Spacing

In our code, if we want to update Number of Instances, then we need to update `patternSpacing` variable only.

```vb
' Update Pattern Spacing
patternSpacing = 10 * AngleConversionFactor
```

In above line we **Update Pattern Spacing** to new value of 10 degree.

***Example Images:***

Below image shows before and after we update **Pattern Spacing**.

**Before Update Pattern Spacing**

![after-update-number-of-instances](/assets/Solidworks_Images/sketch-patterns/after-update-number-of-instances.png)

**After Update Pattern Spacing**

![after-update-pattern-spacing](/assets/Solidworks_Images/sketch-patterns/after-update-pattern-spacing.png)

#### CASE 5 : Update Display Rotation of Pattern

If we want to update Display Rotation of Pattern, then we need to update value to either `True` or `False`.

In our code, we set this value to `True` which means we are displaying the rotation of pattern.

We update our code for not displaying the rotation of pattern as given in below code sample.

```vb
' Edit a Circular Sketch Pattern
BoolStatus = swSketchManager.EditCircularSketchStepAndRepeat(arcRadius, arcAngle, numberOfInstance, patternSpacing, False, "", True, True, True, "Arc1_")
```

***Example Images:***

Below image shows before and after we update **Display Rotation of Pattern**.

**Before Update Display Rotation of Pattern**

![after-update-pattern-spacing](/assets/Solidworks_Images/sketch-patterns/after-update-pattern-spacing.png)

**After Update Display Rotation of Pattern**

![after-update-rotation-of-pattern](/assets/Solidworks_Images/sketch-patterns/after-update-rotation-of-pattern.png)

#### CASE 6 : Update Number of Instances to Delete

If we want to update Number of Instances to Delete, then we need to update value of `""` as given in below code sample.

```vb
' Edit a Circular Sketch Pattern
BoolStatus = swSketchManager.EditCircularSketchStepAndRepeat(arcRadius, arcAngle, numberOfInstance, patternSpacing, False, "(3)", True, True, True, "Arc1_")
```

In above code sample, we want to delete 3rd instance hence we pass the number **`3`** inside **`()`**.

> **Note: For delete any instance we need to pass its position in paranthesis (). Otherwise it won't work.**

***Example Images:***

Below image shows before and after we update **Number of Instances to Delete**.

**Before Update Number of Instances to Delete**

![after-update-pattern-spacing](/assets/Solidworks_Images/sketch-patterns/after-update-pattern-spacing.png)

**After Update Number of Instances to Delete**

![after-update-number-of-instance-to-delete](/assets/Solidworks_Images/sketch-patterns/after-update-number-of-instance-to-delete.png)

#### CASE 7 : Update Display Radius Dimension

If we want to update Display Radius Dimension, then we need to update value to either `True` or `False`.

In our code, we set this value to `True` which means we are displaying the Display Radius Dimension.

We update our code for not displaying the Display Radius Dimension as given in below code sample.

```vb
' Edit a Circular Sketch Pattern
BoolStatus = swSketchManager.EditCircularSketchStepAndRepeat(arcRadius, arcAngle, numberOfInstance, patternSpacing, False, "(3)", False, True, True, "Arc1_")
```

***Example Images:***

Below image shows before and after we update **Display Radius Dimension**.

**Before Update Display Radius Dimension**

![after-update-number-of-instance-to-delete](/assets/Solidworks_Images/sketch-patterns/after-update-number-of-instance-to-delete.png)

**After Update Display Radius Dimension**

![after-update-display-radius-dimension](/assets/Solidworks_Images/sketch-patterns/after-update-display-radius-dimension.png)

#### CASE 8 : Update Display Angle Dimension

If we want to update Display Angle Dimension, then we need to update value to either `True` or `False`.

In our code, we set this value to `True` which means we are displaying the Display Angle Dimension.

We update our code for not displaying the Display Angle Dimension as given in below code sample.

```vb
' Edit a Circular Sketch Pattern
BoolStatus = swSketchManager.EditCircularSketchStepAndRepeat(arcRadius, arcAngle, numberOfInstance, patternSpacing, False, "(3)", False, False, True, "Arc1_")
```

***Example Images:***

Below image shows before and after we update **Display Angle Dimension**.

**Before Update Display Angle Dimension**

![after-update-number-of-instance-to-delete](/assets/Solidworks_Images/sketch-patterns/after-update-number-of-instance-to-delete.png)

**After Update Display Angle Dimension**

![after-update-display-angle-dimension](/assets/Solidworks_Images/sketch-patterns/after-update-display-angle-dimension.png)

#### CASE 9 : Update Display Number of Instances

If we want to update Display Number of Instances, then we need to update value to either `True` or `False`.

In our code, we set this value to `True` which means we are displaying the Display Number of Instances.

We update our code for not displaying the Display Number of Instances as given in below code sample.

```vb
' Edit a Circular Sketch Pattern
BoolStatus = swSketchManager.EditCircularSketchStepAndRepeat(arcRadius, arcAngle, numberOfInstance, patternSpacing, False, "(3)", False, False, False, "Arc1_")
```

***Example Images:***

Below image shows before and after we update **Display Number of Instances**.

**Before Update Display Number of Instances**

![after-update-number-of-instance-to-delete](/assets/Solidworks_Images/sketch-patterns/after-update-number-of-instance-to-delete.png)

**After Update Display Number of Instances**

![after-update-display-number-of-instance](/assets/Solidworks_Images/sketch-patterns/after-update-display-number-of-instance.png)

---

**This is it !!!**

*It is indeed a very LONG post. But I try to update the code and move into the direction where we were able to use these code samples in UserForms.*

*I hope you like my effort!!!*

If you found anything to add or update, please let me know on my e-mail.

Hope this post helps you to *Edit a Circular Sketch Pattern* with Solidworks VBA Macros.

For more such tutorials on **Solidworks VBA Macro**, do come to this blog after sometime.

*If you like the post then please share it with your friends also.*

*Do let me know by you like this post or not!*

*Till then, Happy learning!!!*
