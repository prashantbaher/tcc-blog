---
categories: VBA Userforms
title:  VBA Userforms - Fillet All Corners of a Polygon From VBA Macro
---

In this post, I tell you about *how to Fillet All Corners of a Polygon through Solidworks VBA Macros* in a sketch.

This post is an **Example type post** in which I show you following things:

1. How to use **Input Box** for input value as *number of sides*

2. How to create a **Polygon** of different number of sides *(User define the sides not our macro!!!)*

3. How to **Fillet** All corners of User-defined created **Polygon**

4. How to use ***Checks*** for *better code performance* and *avoid failures* in Code.

From my point of view, this is a **Very Important** post for a *Solidworks VBA macro beginner* because

* *1st I create the code for desired functionality*

* *Next I break those functionality*

* *After breaking I fix those functionality*

**This is the process I follow. It is personal to me.**

If you like this process do follow otherwise I believe every person is differnt.

Hence you can follow a different process.

*Since I am a self-taught programmer, please overlook my silly mistakes in this post!!!*

---

## Content

- [Code Demo Video on YouTube](#video-of-code-on-youtube)

- [For Experience Macro Developers](#for-experience-macro-developer---fillet-all-corners-of-a-polygon-from-vba-macro)

- [For Beginner Macro Developers](#for-beginners-macro-developers---fillet-all-corners-of-a-polygon-from-vba-macro)

  - [Understanding the Code](#understanding-the-code)

  - [NOTE](#note)

- [VBA Language feature used in this post](#vba-language-feature-used-in-this-post)

- [Solidworks API Objects](#solidworks-api-objects)

Feel free to select the topic you want to.

---

## Video of Code on YouTube

Please see below video how visually we can create *a Chamfer* from **Solidworks VBA macro**.

<div class="w3-card">
  <iframe class="w3-panel w3-mobile" height="500px" width="100%" src="https://www.youtube.com/embed/HobbXAv9zMI" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
</div>

Please note that there are **no explaination** given in the video. 

**Explaination** of each line and why we write code this way is given in this post.

---

## For Experience Macro Developer - Fillet All Corners of a Polygon From VBA Macro

If you are an experience **Solidworks Macro developer**, then you are looking for a specific code sample.

Below is the code for creating **A Chamfer** from **Solidworks VBA Macro**.

```vb
' Creating variable for Solidworks Sketch Segment
Dim swSketchSegment As SldWorks.SketchSegment
      
' Set the value of Solidworks Sketch segment by "CreateChamfer" method from Solidworks sketch manager
Set swSketchSegment = swSketchManager.CreateChamfer(swSketchChamferType_e.swSketchChamfer_DistanceEqual, 0.1, 0.2)
```

For creating a **Chamfer** first you need to **Create** a variable of `SketchSegment` type.

After creating variable, you need to set the value of this variable.

For this you used `CreateChamfer` method from **Solidworks Sketch Manager**.

This `CreateChamfer` method set the value of `SketchSegment` type variable.

This `CreateChamfer` method takes following parameters as explained:

**Type** : *Type of chamfer as defined in `swSketchChamferType_e`*

**Distance** : *Distance of the chamfer*

**AngleORdist** : *These are as follows*

* If Type = `swSketchChamfer_DistanceDistance`, then the second chamfer distance 

* If Type = `swSketchChamfer_DistanceAngle`, then the second chamfer angle 

* If Type = `swSketchChamfer_DistanceEqual`, then this argument is ignored because Distance
is used for both edges

If you want a more detail explaination then please read further otherwise this will help you to **Create a Chamfer From VBA Macro**.

---

## For Beginners Macro Developers - Fillet All Corners of a Polygon From VBA Macro

In this post, I tell you about `CreateChamfer` method from **Solidworks** `SketchManager` object.

This method is ***most updated*** method, I found in *Solidworks API Help*. 

So ***use this method*** if you want to create a Chamfer.

Below is the `code` sample for creating a Chamfer.

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

' Creating variable for Solidworks Selection Manager
Dim swSelectionManager As SldWorks.SelectionMgr

' Creating variable for Solidworks Select Data
Dim swSelectData As SldWorks.SelectData

' Creating variable for Solidworks Sketch
Dim swSketch As SldWorks.Sketch

' Creating variable for Solidworks Sketch Point
Dim swSketchPoint As SldWorks.SketchPoint

' Creating variable for Solidworks Sketch Segment
Dim swSketchSegment As SldWorks.SketchSegment

' Main function of our VBA program
Sub main()

  ' Set Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  
  ' String type variable for storing Input Box value
  Dim NumberSidesOfPolygon As String
  
  ' Ask the user to Enter the number of Sides s/he wants for polygon
  ' And store that value into our String type variable
  NumberSidesOfPolygon = InputBox("Please Enter number of Sides you want in Polygon: ", "TheCADCoder Macro Example")
  
  '''''''''''''''''''''' REQUIRED CHECKS '''''''''''''''''''''''''''''''''''''''''''''
  ' We need to check few things before we proceed
  '  1. There should be an input value and Input Value should be a "NUMBER", no Alphabet or Alpha-numeric
  '  2. Input number should be greater than number "2"
  '    Because for a polygon minimum number of side is 3. So we need to check for that also.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  ' CHECK 1: THERE SHOULD BE AN INPUT VALUE AND INPUT NUMBER SHOULD BE A NUMBER ONLY
  '----------------------------------------------
  ' Do while loop to check if the Input value is a Number or not
  ' In condition, we use "IsEmpty()" Build-in function of VBA.
  ' This function check, if the given value inside () is a Empty or not
  ' In condition, we use "IsNumeric()" Build-in function of VBA.
  ' This function check, if the given value inside () is a number or not
  ' If not then we inform user to given a Number (Numeric value)
  ' Then we ask again from User to give a value
  ' This process will goes on until user gives us a Numeric Value (Number)
  
  ' Do while loop to check condition if the Input value is a Number or not
  Do While Not (IsEmpty(NumberSidesOfPolygon) Or IsNumeric(NumberSidesOfPolygon))
  
    ' If not then we inform user to given a Number (Numeric value)
    MsgBox ("Please Enter numbers only.")
    
    ' Ask the user to Enter the number of Sides s/he wants for polygon
    ' And store that value into our String type variable
    NumberSidesOfPolygon = InputBox("Please Enter number of Sides you want in Polygon: ", "TheCADCoder Macro Example")
  Loop
  
  ' CHECK 2: INPUT NUMBER SHOULD BE GREATER THAN NUMBER "3"
  '--------------------------------------------------------
  ' Do while loop to check condition if the Input value is less than or equal to 2
  ' In condition, we use "Conversion.Int()" Build-in function of VBA.
  ' This function converts the string into Integer value, which we can use for mathematical operations
  ' If Input value is not less than or equal to 2, then we inform user with a required Message
  ' Then we ask again from User to give a value greater than 2
  ' This process will goes on until user gives us a Numeric Value (Number)
  
  ' Do while loop to check condition if the Input value is less than or equal to 2
  Do While (Conversion.Int(NumberSidesOfPolygon) <= 2)
  
    ' If Input value is not less than or equal to 2, then we inform user with a required Message
    MsgBox ("Please Enter number Greater than 2.")
    
    ' Ask the user to Enter the number of Sides s/he wants for polygon
    ' And store that value into our String type variable
    NumberSidesOfPolygon = InputBox("Please Enter number of Sides you want in Polygon: ", "TheCADCoder Macro Example")
  Loop
  
  ' Create a String type variable for storing default part location
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
  
  ' Creating a "Variant" type Variable for Polygon
  Dim vSketchLines As Variant
  
  ' Creating a Corner Polygon
  vSketchLines = swSketchManager.CreatePolygon(0, 0, 0, 1, 0, 0, NumberSidesOfPolygon, True)
  
  ' De-select the lines after creation
  swDoc.ClearSelection2 True
  
  ' Zoom to fit screen in Solidworks Window
  swDoc.ViewZoomtofit
  
  ' Set the value of Solidworks Selection Manager with document Selection manager
  Set swSelectionManager = swDoc.SelectionManager
  
  ' Set the value of Solidworks Sketch to Active Sketch
  Set swSketch = swSketchManager.ActiveSketch
  
  ' Create a Variant type Local variable for storing all sketch points
  Dim vSketchPointArray As Variant
  
  ' Get all the Sketch point in this active sketch and store them into Variant type variable
  vSketchPointArray = swSketch.GetSketchPoints2
  
  ' Set the value of Solidworks Select Data to new Select Data
  Set swSelectData = swSelectionManager.CreateSelectData
  
  ' Local Integer type variable for looping purpose
  Dim i As Integer
  
  ' For-Loop: from 0 to number of values in "vSketchPointArray" variable
  For i = 0 To UBound(vSketchPointArray)
    
    ' Set the value of Solidworks Sketch Point to Current Point
    Set swSketchPoint = vSketchPointArray(i)
    
    ' Select the Current Point using Select4 method from Solidworks Sketch Point
    BoolStatus = swSketchPoint.Select4(True, swSelectData)

    ' Create a Fillet
    Set swSketchSegment = swSketchManager.CreateFillet(0.1, swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry)
    
    ' De-select the Point after creation
    swDoc.ClearSelection2 True
    
  Next i
  
  ' De-select the Sketch
  swDoc.ClearSelection2 True
  
  ' Show "Front" view
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

For more information please visit [Solidworks Macros - Open new Part document](/solidworks-macro/open-new-document) post.

```vb
' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
```

In this line, we are creating a variable which we named as `swApp` and the type of this `swApp` variable is `SldWorks.SldWorks`.

```vb
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2
```

In this line, we are creating a variable which we named as `swDoc` and the type of this `swDoc` variable is `SldWorks.ModelDoc2`.

```vb
' Boolean Variable
Dim BoolStatus As Boolean
```

In this line, we create a variable named `BoolStatus` as `Boolean` object type.

```vb
' Creating variable for Solidworks Sketch Manager
Dim swSketchManager As SldWorks.SketchManager
```

In above line, we create variable `swSketchManager` for **Solidworks Sketch Manager**.

As the name suggested, a **Sketch Manager** holds variours methods and properties to manage *Sketches*.

To see methods and properties related to `SketchManager` object, please visit [this page](help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html)

```vb
' Creating variable for Solidworks Sketch Segment
Dim swSketchSegment As SldWorks.SketchSegment
```

In this line, we are creating a variable which we named as `swSketchSegment` and the type of this `swSketchSegment` variable is `SldWorks.SketchSegment`.

We create variable `swSketchSegment` for **Solidworks Sketch Segments**.

To see methods and properties related to `swSketchSegment` object, please visit [this page](http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment_members.html)

These all are our global variables.

As you can see in code sample, they are **Solidworks API Objects**.

So basically I group all the **Solidworks API Objects** in one place.

I have also place `boolean` type object at top also, because after certain point we will *need* this variable frequently.

Thus, I have started placing it here.

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

```vb
' Setting Solidworks document to new part document
Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
```

In this line, we set the value of our `swDoc` variable to new document.

For **detailed information** about these lines please visit [Solidworks Macros - Open new Part document](/solidworks-macro/open-new-document) post.

I have discussed them **thoroghly** in [Solidworks Macros - Open new Part document](/solidworks-macro/open-new-document) post, so do checkout this post if you don't understand above code.

```vb
' Selecting Front Plane
BoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
```

In above line, we select the *front plane* by using `SelectByID2` method from `Extension` object.

For more information about selection method please visit [Solidworks Macros - Selection Methods](/solidworks-macro/select-plane-from-tree) post.

```vb
' Setting Sketch manager for our sketch
Set swSketchManager = swDoc.SketchManager
```

In above line, we set the sketch manager variable to current document's sketch manager.

```vb
' Inserting a sketch into selected plane
swSketchManager.InsertSketch True
```

In above line, we use `InsertSketch` method of *SketchManager* and give `True` value.

This method allows us to insert a sketch in selected plane.

```vb
' Creating a "Variant" Variable which holds the values return by "CreateCornerRectangle" method
Dim vSketchLines As Variant
    
' Creating a Corner Rectangle
vSketchLines = swSketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0)
```

In above sample code, we 1st create a variable named `vSketchLines` of type `Variant`.

A `Variant` type variable can hold **any** type of value depends upon the use of variable.

In 2nd line, we set the value of variable `vSketchLines`.

Value of `vSketchLinesis` an array of lines. This array is send as return value when we use `CreateCornerRectangle` method.

This `CreateCornerRectangle` method is part of `swSketchManager` and it is the latest method to create a corner rectangle.

For detail explaination on `CreateCornerRectangle` method, please see [Sketch - Create Corner Rectangle](/solidworks-macro/create-corner-rectangle) post.

In the above code sample I have used (0, 1, 0) Upper-left point in *Y-direction*.

For Lower-right point I used (1, 0, 0) which is 1 point distance in *X-direction*.

```vb
' De-select the Rectangle after creation
swDoc.ClearSelection2 True
```

In above line, we de-select the ractangle we just create.

```vb
' Selecting Front Plane
BoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
```

In above line, we select the *front plane* by using `SelectByID2` method from `Extension` object.

For more information about selection method please visit [Solidworks Macros - Selection Methods](/solidworks-macro/select-plane-from-tree) post.

```vb
' Set the value of Solidworks Sketch segment by "CreateChamfer" method from Solidworks sketch manager
Set swSketchSegment = swSketchManager.CreateChamfer(swSketchChamferType_e.swSketchChamfer_DistanceEqual, 0.1, 0.2)
```

In above line, we set the value of Solidworks Sketch Segment variable `swSketchSegment` by `CreateChamfer` method from *Solidworks Sketch Manager*.

This `CreateChamfer` method takes following parameters:

**Type** : *Type of chamfer as defined in `swSketchChamferType_e`*.

The `swSketchChamferType_e` has 3 values for type of chamfers:

* `swSketchChamfer_DistanceAngle`

* `swSketchChamfer_DistanceDistance`

* `swSketchChamfer_DistanceEqual`

**Distance** : *Distance of the chamfer*

**AngleORdist** : *Angle or Distance for chamfer. These are as follows*

* If Type = `swSketchChamfer_DistanceDistance`, then the second chamfer distance 

* If Type = `swSketchChamfer_DistanceAngle`, then the second chamfer angle 

* If Type = `swSketchChamfer_DistanceEqual`, then this argument is ignored because Distance
is used for both edges.

Below Image described **the Parameters for a Chamfer**.

![fillet_parameters](/assets/Solidworks_Images/fillet and chamfer/fillet_parameters.png)

In our code, I have used following values:

**Type** : I have used `swSketchChamferType_e.swSketchChamfer_DistanceEqual` enumerator as value for type of Chamfer.

**Distance** : I have used 0.1 (This value is in meter) as the distance of Chamfer.

**AngleORdist** : I have used 0.2 (This value is in meter). But in our code **Type = `swSketchChamfer_DistanceEqual`**, then this argument is ignored because Distance
is used for both edges.

### NOTE

It is ***very important*** to remember that, when you give distance or any other numeric value in **Solidworks API**, Solidworks takes that numeric value in ***Meter only***.

Solidworks API does not care about your application's Unit systems.

For example, I works in **ANSI** system means inches for distance. But when I used **Solidworks API** through *VBA macros or C#*, I need to use converted numeric values.

Because Solidworks API output the distance in **Meter** which is not my requirement.

```vb
' De-select the Fillet after creation
swDoc.ClearSelection2 True
```

In the above line of code, we deselect the **Chamfer** we have created.

For de-selecting, we use `ClearSelection2` method from our Solidworks document name `swDoc`.

```vb
' Show Front View after creating Chamfer
swDoc.ShowNamedView2 "", swStandardViews_e.swFrontView
```

In the above line of code, we update the *view orientation* to **Front View**.

In my machine, after inserting a sketch view orientation does not changed.

Because of this I have to update the view to **Front view**.

For showing **Front View** we used `ShowNamedView2` method from our Solidworks document name `swDoc`.

This method takes 2 parameter described as follows:

**VName** : Name of the view to display or an empty string to use ViewId instead

**ViewId** : ID of the view to display as defined by `swStandardViews_e` or -1 to use the **VName** argument instead.

*NOTE:* If you specify both **VName** and **ViewId**, then **ViewId** takes precedence if the two arguments do not resolve to the same view.

`swStandardViews_e` has following Standard View Types:

- *swBackView*

- *swBottomView*

- *swDimetricView*

- *swFrontView*

- *swIsometricView*

- *swLeftView*

- *swRightView*

- *swTopView*

- *swTrimetricView*

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

If you want to read more about **Option Explicit** then please visit [Declaring and Scoping of Variables](/visual-basic/vba-declaring-and-scoping-of-variables).

2. Then we create **variable** for different data types.

If you don't know about them, then please visit [Variables](/visual-basic/vba-variables) and [Data-types](/visual-basic/vba-programming-concepts-comments-and-datatypes) posts of this blog.

These posts will help you to understand what **Variables** are and how to use them.

3. Then we create **main Sub procedure** for our macro.

If you don't know about the **Sub procedure**, then I suggest you to visit [VBA Sub and Function Procedures](/visual-basic/vba-sub-and-function-procedure) and [Executing Sub and Function Procedures](/visual-basic/vba-executing-procedures) posts of this blog.

These posts will help you to understand what **Procedures** are and how to use them.

4. In most part we create some variables and set their values. We set those values by using some **functions** provided from objects.

If you don't know about the **functions**, then you should visit [VBA Functions](/visual-basic/vba-functions) and [VBA Functions that do more](/visual-basic/vba-more-function) posts of this blog.

These posts will help you to understand what **functions** are and how to use them.

---

## Solidworks API Objects

In this post, for creating a **Fillet**, we use *Solidworks API objects and their methods*.

This section contains the list of all **Solidworks Objects** used in this post.

I have also attached links of these **Solidworks API Objects** in **API Help website**.

If you want to explore those objects, you can use these links.

These Solidworks API Objects are listed below:

- **Solidworks Application Object**

If you want explore ***Properties and Methods/Functions*** of **Solidworks Application Object** object you can visit [this link](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html).

- **Solidworks Document Object**

If you want explore ***Properties and Methods/Functions*** of **Solidworks Document Object** object you can visit [this link](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html).

- **Solidworks Sketch Manager Object**

If you want explore ***Properties and Methods/Functions*** of **Solidworks Sketch Manager Object** you can visit [this link](help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html).

- **Solidworks Sketch Segment Object**

If you want explore ***Properties and Methods/Functions*** of **Solidworks Sketch Segment Object** you can visit [this link](http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment_members.html).

---

Hope this post helps you to *create a Chamfer* in Sketches with Solidworks VB Macros.

For more such tutorials on **Solidworks VBA Macros**, do come to this blog after sometime.

Till then, Happy learning!!!