---
id: sw-sketch-macro-extend-entities
title:  Extend Sketch Entities
---

In this post, I tell you about *how to Extend Sketch Entities using Solidworks VBA Macros* in a Sketch.

---

## Video of Code on YouTube

Please see below video how visually we *Extend Sketch Entities* in **Solidworks VBA macro**.

<div class="youtube-responsive-container">
<iframe src="https://www.youtube.com/embed/tINYT057tFM" frameborder="0" allowfullscreen></iframe>
</div>

Please note that there are **no explaination** given in the video. 

**Explaination** of each line and why we write code this way is given in this post.

---

## For Experience Macro Developer

If you are an experience **Solidworks Macro developer**, then you are looking for a specific code sample.

Below is the code for **Extend Sketch Entities** from **Solidworks VBA Macro**.

```vb
' Boolean Variable
Dim BoolStatus As Boolean

' Select Line 1
BoolStatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)

' Extend Solidworks Sketch segment by "SketchExtend" method from Solidworks sketch manager
BoolStatus = swSketchManager.SketchExtend(0, 0, 0)
```

**Method Name**: `SketchExtend`

**Description**: Add lengths to the selected Sketch Entity till next sketch entity.

For **Extend** Solidworks Sketch segment, first you need to **Create** a variable of `Boolean` type.

After creating variable, you need to set the value of this `Boolean` variable.

For this you used `SketchExtend` method from **Solidworks Sketch Manager**.

This `SketchExtend` method set the value of `Boolean` type variable.

If Extend is **successful** then `SketchExtend` method return **True** otherwise `SketchExtend` returns **False**.

This `SketchExtend` method takes following parameters as explained:

 - **X** : *Not used*

 - **Y** : *Not used*

 - **Z** : *Not used*

If you want a more detail explaination then please read further otherwise this will help you to **Extend Sketch Entities From VBA Macro**.

---

## For Beginners Macro Developers

In this post, I tell you about `SketchExtend` method from **Solidworks** `SketchManager` object.

This method is ***most updated*** method, I found in *Solidworks API Help*. 

So ***use this method*** if you want to *Extend Sketces*..

Below is the `code` sample for *Extend Sketces*.

```vb
Option Explicit

' Create variable for Solidworks application
Dim swApp As SldWorks.SldWorks

' Create variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2

' Boolean Variable
Dim BoolStatus As Boolean

' Create variable for Solidworks Sketch Manager
Dim swSketchManager As SldWorks.SketchManager

' Create Variable for Solidworks Sketch Segment
Dim swSketchSegment As SldWorks.SketchSegment

' Main function of our VBA program
Sub main()

  ' Set Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  
  ' Create string type variable for storing default part location
  Dim defaultTemplate As String
  ' Set value of this string type variable to "Default part template"
  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)

  ' Set Solidworks document to new part document
  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)

  ' Select Front Plane
  BoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)

  ' Set Sketch manager for our sketch
  Set swSketchManager = swDoc.SketchManager

  ' Insert a sketch into selected plane
  swSketchManager.InsertSketch True
  
  ' Set Sketch Segment value and Create Line 1
  Set swSketchSegment = swSketchManager.CreateLine(-0.5, 0.75, 0, -0.25, -0.5, 0)
  
  ' Set Sketch Segment value and Create Line 2
  Set swSketchSegment = swSketchManager.CreateLine(-0.75, -1.25, 0, 0.5, -1.25, 0)
  
  ' De-select the lines after creation
  swDoc.ClearSelection2 True
  
  ' Select Line 1
  BoolStatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)

  ' Extend selected Sketch Segments by "SketchExtend" method from Solidworks sketch manager
  BoolStatus = swSketchManager.SketchExtend(0.0, 0.0, 0.0)

  ' De-select the Sketch Segment after Extend
  swDoc.ClearSelection2 True
  
  ' Show Front View after Extend Sketch Segments
  swDoc.ShowNamedView2 "", swStandardViews_e.swFrontView
  
  ' Zoom to fit screen in Solidworks Window
  swDoc.ViewZoomtofit2

End Sub
```

---

### Understanding the Code

Now let us walk through *each line* in the above code, and **understand** the meaning of every line.

```vb
Option Explicit
```

This line forces us to define every variable we are going to use. 

For more information please visit **[Solidworks Macros - Open new Part document](sw-macro-open-part)** post.

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

```vb
' Create variable for Solidworks Sketch Manager
Dim swSketchManager As SldWorks.SketchManager
```

In above line, we create variable `swSketchManager` for **Solidworks Sketch Manager**.

As the name suggested, a **Sketch Manager** holds variours methods and properties to manage *Sketches*.

To see methods and properties related to `SketchManager` object, please visit **[this page](help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html)**.

```vb
' Create variable for Solidworks Sketch Segment
Dim swSketchSegment As SldWorks.SketchSegment
```

In this line, we Create a variable which we named as `swSketchSegment` and the type of this `swSketchSegment` variable is `SldWorks.SketchSegment`.

We create variable `swSketchSegment` for **Solidworks Sketch Segments**.

To see methods and properties related to `swSketchSegment` object, please visit **[this page](http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment_members.html)**.

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

In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application.

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
' Set Solidworks document to new part document
Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
```

In this line, we set the value of our `swDoc` variable to new document.

For **detailed information** about these lines please visit **[Solidworks Macros - Open new Part document](sw-macro-open-part)** post.

I have discussed them **thoroghly** in **[Solidworks Macros - Open new Part document](sw-macro-open-part)** post, so do checkout this post if you don't understand above code.

```vb
' Select Front Plane
BoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
```

In above line, we select the *front plane* by using `SelectByID2` method from `Extension` object.

For more information about selection method please visit **[Solidworks Macros - Selection Methods](sw-macro-selection-methods)** post.

```vb
' Set Sketch manager for our sketch
Set swSketchManager = swDoc.SketchManager
```

In above line, we set the sketch manager variable to current document's sketch manager.

```vb
' Insert a sketch into selected plane
swSketchManager.InsertSketch True
```

In above line, we use `InsertSketch` method of *SketchManager* and give `True` value.

This method allows us to insert a sketch in selected plane.

```vb
' Set Sketch Segment value and Create Line 1
Set swSketchSegment = swSketchManager.CreateLine(-0.5, 0.75, 0, -0.25, -0.5, 0)
```

In above line, we set the value of Solidworks Sketch Segment variable `swSketchSegment` by `CreateLine` method from *Solidworks Sketch Manager*.

This `CreateLine` method creates Lines between 2 given points.

For more information about `CreateLine` method, you can read my **[Solidworks Sketch Macros - Create Line](sw-sketch-macro-create-line)** post.

This post describe all the parameters we need for this `CreateLine` method.

In above line, we create a Line between **point (-0.5, 0.75, 0)** and **point (-0.25, -0.5, 0)**.

This line is **not parallel** to any axis.

This line is at an angle to *Vertical* axis.

```vb
' Set Sketch Segment value and Create Line 2
Set swSketchSegment = swSketchManager.CreateLine(-0.75, -1.25, 0, 0.5, -1.25, 0)
```

In above line we create Line 2 by using same `CreateLine` method from *Solidworks Sketch Manager*.

In above code, we create our 2nd line between **point (-0.75, -1.25, 0)** and **point (0.5, -1.25, 0)**.

This line is **parallel** to X-axis.

```vb
' De-select the lines after creation
swDoc.ClearSelection2 True
```

After creating both lines we de-select those lines.

:::note
We **don't need** to de-select the lines for **Extend operation** as I will select those lines agains in next 2 lines. I just want to show you how to select a **Sketch Segment** with `SelectById` Menthod.
:::

```vb
' Select Line 1
BoolStatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
```

In above line of code, we select **Line 1**.

```vb
' Extend selected Sketch Segments by "SketchExtend" method from Solidworks sketch manager
BoolStatus = swSketchManager.SketchExtend(0.0, 0.0, 0.0)
```

In above line, we **Extend** selected *Sketch Segments* by `SketchExtend` method from *Solidworks sketch manager*.

This `SketchExtend` method takes following parameters:

 - **X** : *Not used*

 - **Y** : *Not used*

 - **Z** : *Not used*

 - **Return Value**:

    - **True**: If Extend operation is *Success*.

    - **False**: If Extend operation is *Fail*.

In our code, I have used following values:

 - **X** : I have used 0.0 value for *X*.

 - **Y** : I have used 0.0 value for *Y*.

 - **Z** : I have used 0.0 value for *Z*.

Below image shows before and after Extend operation on the sketch.

**Before Extend Operation**

![before_extend](/assets/Solidworks_Images/trim_and_extend/before_extend.png)

**After Extend Operation**

![after_extend](/assets/Solidworks_Images/trim_and_extend/after_extend.png)

---

### NOTE

It is ***very important*** to remember that, when you give distance or any other numeric value in **Solidworks API**, Solidworks takes that numeric value in ***Meter only***.

Solidworks API does not care about your application's Unit systems.

For example, I works in **ANSI** system means inches for distance. But when I used **Solidworks API** through *VBA macros or C#*, I need to use converted numeric values.

Because Solidworks API output the distance in **Meter** which is not my requirement.

---

```vb
' De-select the Sketch after creation
swDoc.ClearSelection2 True
```

In the above line of code, we deselect the **Sketch** after the *Extend* operation.

For de-selecting, we use `ClearSelection2` method from our Solidworks document name `swDoc`.

```vb
' Show Front View after Sketch Extend
swDoc.ShowNamedView2 "", swStandardViews_e.swFrontView
```

In the above line of code, we update the *view orientation* to **Front View**.

In my machine, after inserting a sketch view orientation does not changed.

Because of this I have to update the view to **Front view**.

For showing **Front View** we used `ShowNamedView2` method from our Solidworks document name `swDoc`.

This method takes 2 parameter described as follows:

  - **VName** : Name of the view to display or an empty string to use ViewId instead

  - **ViewId** : ID of the view to display as defined by `swStandardViews_e` or -1 to use the **VName** argument instead.

:::note
If you specify both **VName** and **ViewId**, then **ViewId** takes precedence if the two arguments do not resolve to the same view.
:::

`swStandardViews_e` has following Standard View Types:

- *swBackView*

- *swBottomView*

- *swDimetricView*

- *swFrontView*

- *swIsometricView*

- *swLeftView*

- *swRightView*

- *swTopView*

- *swExtendetricView*

In our code, we did not use **VName** instead I used *empty string* in form of ***""*** symbol.

I used **ViewId** value to specify view and used `swStandardViews_e.swFrontView` value to use *Standard Front View*.

```vb
' Zoom to fit screen in Solidworks Window
swDoc.ViewZoomtofit
```

In this last line we use *zoom to fit* command.

For Zoom to fit, we use `ViewZoomtofit` method from our Solidworks document variable `swDoc`.

This is it !!!

If you found anything to add or update, please let me know on my e-mail.

---

## VBA Language feature used in this post

In this post used some features of **VBA programming language**.

This section of post, has some brief information about the VBA programming language specific features.

1. We use **Option Explicit** for capturing un-declared variables.

If you want to read more about **Option Explicit** then please visit **[Declaring and Scoping of Variables](../vba/vba-variables-decl)**.

2. Then we create **variable** for different data types.

If you don't know about them, then please visit **[Variables](../vba/vba-variables)** and **[Data-types](../vba/vba-prog-concept#data-types-in-vba)** posts of this blog.

These posts will help you to understand what **Variables** are and how to use them.

3. Then we create **main Sub procedure** for our macro.

If you don't know about the **Sub procedure**, then I suggest you to visit **[VBA Sub and Function Procedures](../vba/vba-procedures)** and **[Executing Sub and Function Procedures](../vba/vba-procedures-exec)** posts of this blog.

These posts will help you to understand what **Procedures** are and how to use them.

4. In most part we create some variables and set their values. We set those values by using some **functions** provided from objects.

If you don't know about the **functions**, then you should visit **[VBA Functions](../vba/vba-functions)** and **[VBA Functions that do more](../vba/vba-functions-adv)** posts of this blog.

These posts will help you to understand what **functions** are and how to use them.

---

## Solidworks API Objects

In this post of **Sketch Extend**, we use *Solidworks API objects and their methods*.

This section contains the list of all **Solidworks Objects** used in this post.

I have also attached links of these **Solidworks API Objects** in **API Help website**.

If you want to explore those objects, you can use these links.

These Solidworks API Objects are listed below:

- **Solidworks Application Object**

  If you want explore ***Properties and Methods/Functions*** of **Solidworks Application Object** object you can visit **[this link](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html)**.

- **Solidworks Document Object**

  If you want explore ***Properties and Methods/Functions*** of **Solidworks Document Object** object you can visit **[this link](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html)**.

- **Solidworks Sketch Manager Object**

  If you want explore ***Properties and Methods/Functions*** of **Solidworks Sketch Manager Object** you can visit **[this link](help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html)**.

- **Solidworks Sketch Segment Object**

  If you want explore ***Properties and Methods/Functions*** of **Solidworks Sketch Segment Object** you can visit **[this link](http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment_members.html)**.

---

Hope this post helps you to *Extend* Sketch Entities with Solidworks VB Macros.

For more such tutorials on **Solidworks VBA Macros**, do come to this blog after sometime.

Till then, Happy learning!!!