---
categories: Solidworks-macros
title:  Solidworks Macros - Create a Point From VBA Macro
---

In this post, I tell you about *how to create a Point through Solidworks VBA Macros* in a sketch.

The process is almost identical with previous [Sketch - Create Lines](/solidworks-macros/sketch-create-line) post.

In this post, I tell you about `CreatePoint` method from **Solidworks** `SketchManager` object.

By this method 1st we create *a simple point*, after that we create *a sequence of points*.

This method is ***most updated*** method, I found in *Solidworks API Help*. 

So ***use this method*** if you want to create a new **Point**.

Below is the `code` sample for creating *a Point*.

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
  
  ' Inserting a sketch into selected plane
  swSketchManager.InsertSketch True
  
  ' Creating Varient for Polygon
  Dim myPoint As SketchPoint
  
  ' Creating a Point
  Set myPoint = swSketchManager.CreatePoint(0, 1, 0)
  
  ' #########Creating a number of points##############
  
  ' Declaring integer type variable for loop
  Dim i As Integer
  
  ' Looping through 1 to 5
  For i = 0 To 5
  
    ' Declaring integer type variables for X, Y and Z cordinates of point
    Dim x, y, z As Integer
    
    ' Setting values of x, y and z
    x = i
    y = x + i
    z = 0
    
    ' Create points till loop continues
    Set myPoint = swSketchManager.CreatePoint(x, y, z)
    
  Next
  
  ' De-select the Polygon after creation
  swDoc.ClearSelection2 True
  
  ' Zoom to fit screen in Solidworks Window
  swDoc.ViewZoomtofit

End Sub
```

---

## Understanding the Code

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
<!--{%- include amazon-us-native-ad.html -%}-->

```vb
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
```

In this line, we are creating a variable which we named as `swDoc` and the type of this `swDoc` variable is `SldWorks.ModelDoc2`.

Next is our `Sub` procedure named as `main`. This procedure hold all the *statements (instructions)* we give to computer.

```vb
' Setting Solidworks variable to Solidworks application
Set swApp = Application.SldWorks
```

In this line, we are setting the value of our Solidworks variable `swApp` which we defined earlier to Solidworks application.

```vb
' Creating string type variable for storing default part location
Dim defaultTemplate As String
' Setting value of this string type variable to "Default part template"
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
```

In 1st statement of above example, we are defining a variable of `string` type and named it as `defaultTemplate`.

This variable `defaultTemplate`, holds the location the location of **Default Part Template**.

In 2nd line of above example. we assign value to our newly define `defaultTemplate` variable.

We assign the value by using a *Method* named `GetUserPreferenceStringValue()`. 

This method is a part of our main Solidworks variable `swApp`.

```vb
' Setting Solidworks document to new part document
Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
```

In this line, we set the value of our `swDoc` variable to new document.

For **more detailed information** about above lines please visit [Solidworks Macros - Open new Part document](/solidworks-macros/open-new-document) post. 

I have discussed them **thoroghly** in [Solidworks Macros - Open new Part document](/solidworks-macros/open-new-document) post, so do checkout this post if you don't understand above code.

```vb
' Boolean Variable
Dim BoolStatus As Boolean

' Selecting Front Plane
BoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
```

In 1st line, we create a variable named `BoolStatus` as `Boolean` object.

In next line, we select the *front plane* by using `SelectByID2` method from `Extension` object.

For more information about selection method please visit [Solidworks Macros - Selection Methods](/solidworks-macros/select-plane-from-tree) post.

I have discussed about different *Selection methods* in details in [Soldworks Macros - Selection Methods](/solidworks-macros/select-plane-from-tree) post, so do visit this post for more *Selection methods*.

```vb
' Creating variable for Solidworks Sketch Manager
Dim swSketchManager As SldWorks.SketchManager
```

In above line, we create variable `swSketchManager` for **Solidworks Sketch Manager**.

As the name suggested, a **Sketch Manager** holds variours methods and properties to manage *Sketches*.

To see methods and properties related to SketchManager object, please visit [this page](help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html)

```vb
' Setting Sketch manager for our sketch
Set swSketchManager = swDoc.SketchManager
```

In above line, we set the **Sketch manager** variable to current document's sketch manager.

```vb
' Inserting a sketch into selected plane
swSketchManager.InsertSketch True
```

In above line, we use `InsertSketch` method of *SketchManager* and give `True` value.

This method allows us to insert a sketch in selected plane.

```vb
' Creating Variable for Sketch Point
Dim myPoint As SketchPoint
      
' Creating a Point
Set myPoint = swSketchManager.CreatePoint(0, 1, 0)
```

In above sample code, we 1st create a variable named `myPoint` of type `SketchPoint`.

In 2nd line, we **set** the value of *SketchPoint* variable `myPoint`.

We get this value from `CreatePoint` method which is inside the `swSketchManager` variable.

`swSketchManager` variable is a type of **SketchManager**, hence we used `CreateSketchSlot` method from **SketchManager**.

This `CreatePoint` method takes following parameters as explained:

**X** : *X Location of Point*

**Y** : *Y Location of Point*

**Z** : *Z Location of Point*

For creating a *Sketch Point*, I used following parameter Values:

  * **X** : 0

  * **Y** : 1

  * **Z** : 0

This create a point in *Y - Direction* at the distance of 1.

```vb
' Declaring integer type variable for loop
Dim i As Integer

' Looping through 1 to 5
For i = 0 To 5

  ' Declaring integer type variables for X, Y and Z cordinates of point
  Dim x, y, z As Integer
  
  ' Setting values of x, y and z
  x = i
  y = x + i
  z = 0
  
  ' Create points till loop continues
  Set myPoint = swSketchManager.CreatePoint(x, y, z)
  
Next
```

Above Lines of code creates a number of points.

Below Image described **the Parameters for Centerpoint Arc Slot** in more detail.

![centerpoint-arc-slot-parameters](/assets/Solidworks_Images/slots/centerpoint-arc-slot-parameters.png)

This `CreateSketchSlot` method returns *Sketch Slot* interface i.e. `ISketchSlot` interface. 

This `ISketchSlot` interface has various **methods and properties** for *a Slot*.

For more detail about **methods and properties** of `ISketchSlot` interface you can visit [this page](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSlot_members.html)

### NOTE

It is ***very important*** to remember that, when you give distance or any other numeric value in **Solidworks API**, Solidworks takes that numeric value in ***Meter only***.

*Solidworks API* does not care about your application's Unit systems.

For example, I works in ANSI system means "inches" for distance. 

But when I used Solidworks API through *VBA macros* or *C#*, I have to use **converted** numeric values.

Because Solidworks API output the distance in **Meter** only; which is not my requirement.

```vb
' De-select the Slot after creation
swDoc.ClearSelection2 True
```

In the this line of code, we de-select the created Centerpointrc Slot.

For de-selecting, we use `ClearSelection2` method from our Solidworks document variable `swDoc`.

```vb
' Zoom to fit screen in Solidworks Window
swDoc.ViewZoomtofit
```

In this last line we use *zoom to fit* command.

For Zoom to fit, we use `ViewZoomtofit` method from our Solidworks document variable `swDoc`. 

Hope this post helps you to *create a Centerpoint Arc Slot* in Sketches with Solidworks VB Macros.

For more such tutorials on **Solidworks VBA Macros**, do come to this blog after sometime.

Till then, Happy learning!!! 

<!-- This is post navigation bar -->
<div class="w3-bar w3-margin-top w3-margin-bottom">
  <a href="/solidworks-macros/create-3point-arc-slot" class="w3-button w3-rose">&#10094; Previous</a>
  <a href="/solidworks-macros/create-centerpoint-arc-slot" class="w3-button w3-rose w3-right">Next &#10095;</a>
</div>
