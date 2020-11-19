(window.webpackJsonp=window.webpackJsonp||[]).push([[42],{100:function(e,t,n){"use strict";n.r(t),n.d(t,"frontMatter",(function(){return l})),n.d(t,"metadata",(function(){return c})),n.d(t,"rightToc",(function(){return b})),n.d(t,"default",(function(){return i}));var a=n(2),r=n(7),o=(n(0),n(153)),l={id:"sw-sketch-macro-create-straight-slot",title:"Create a Straight Slot"},c={unversionedId:"solidworks-macros/sw-sketch-macro-create-straight-slot",id:"solidworks-macros/sw-sketch-macro-create-straight-slot",isDocsHomePage:!1,title:"Create a Straight Slot",description:"In this post, I tell you about how to create a Straight Slot through Solidworks VBA Macros in a sketch.",source:"@site/docs\\solidworks-macros\\2019-06-19-create-straight-slot.md",slug:"/solidworks-macros/sw-sketch-macro-create-straight-slot",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-straight-slot",version:"current",sidebar:"swvba",previous:{title:"Create Polygon",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-polygon"},next:{title:"Create a Centerpoint Straight Slot",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-centerpoint-straight-slot"}},b=[{value:"Understanding the Code",id:"understanding-the-code",children:[{value:"NOTE",id:"note",children:[]}]}],p={rightToc:b};function i(e){var t=e.components,l=Object(r.a)(e,["components"]);return Object(o.b)("wrapper",Object(a.a)({},p,l,{components:t,mdxType:"MDXLayout"}),Object(o.b)("p",null,"In this post, I tell you about ",Object(o.b)("em",{parentName:"p"},"how to create a Straight Slot through Solidworks VBA Macros")," in a sketch."),Object(o.b)("p",null,"The process is almost identical with previous ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-sketch-macro-create-tangent-arc"}),"Sketch - Create Tangent Arc"))," post."),Object(o.b)("p",null,"In this post, I tell you about ",Object(o.b)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method from ",Object(o.b)("strong",{parentName:"p"},"Solidworks")," ",Object(o.b)("inlineCode",{parentName:"p"},"SketchManager")," object."),Object(o.b)("p",null,"This method is ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"most updated"))," method, I found in ",Object(o.b)("em",{parentName:"p"},"Solidworks API Help"),". "),Object(o.b)("p",null,"So ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"use this method"))," if you want to create a new ",Object(o.b)("strong",{parentName:"p"},"Straight Slot"),"."),Object(o.b)("p",null,"Below is the ",Object(o.b)("inlineCode",{parentName:"p"},"code")," sample for creating ",Object(o.b)("em",{parentName:"p"},"a Straight Slot"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n  \n  ' Creating string type variable for storing default part location\n  Dim defaultTemplate As String\n  ' Setting value of this string type variable to \"Default part template\"\n  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n  ' Setting Solidworks document to new part document\n  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\n  ' Selecting Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  \n  ' Setting Sketch manager for our sketch\n  Set swSketchManager = swDoc.SketchManager\n  \n  ' Inserting a sketch into selected plane\n  swSketchManager.InsertSketch True\n  \n  ' Creating Variable for Solidworks Slot\n  Dim mySketchSlot As SketchSlot\n      \n  ' Creating a Straight slot\n  Set mySketchSlot = swSketchManager.CreateSketchSlot(swSketchSlotCreationType_e.swSketchSlotCreationType_line, swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter, 1, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, False)\n  \n  ' De-select the Slot after creation\n  swDoc.ClearSelection2 True\n  \n  ' Zoom to fit screen in Solidworks Window\n  swDoc.ViewZoomtofit\n\nEnd Sub\n")),Object(o.b)("hr",null),Object(o.b)("h2",{id:"understanding-the-code"},"Understanding the Code"),Object(o.b)("p",null,"Now let us walk through ",Object(o.b)("em",{parentName:"p"},"each line")," in the above code, and ",Object(o.b)("strong",{parentName:"p"},"understand")," the meaning of every line."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n")),Object(o.b)("p",null,"This line forces us to define every variable we are going to use. "),Object(o.b)("p",null,"For more information please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),Object(o.b)("p",null,"Next is our ",Object(o.b)("inlineCode",{parentName:"p"},"Sub")," procedure named as ",Object(o.b)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",Object(o.b)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),Object(o.b)("p",null,"In this line, we are setting the value of our Solidworks variable ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," which we defined earlier to Solidworks application."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),Object(o.b)("p",null,"In 1st statement of above example, we are defining a variable of ",Object(o.b)("inlineCode",{parentName:"p"},"string")," type and named it as ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),Object(o.b)("p",null,"This variable ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate"),", holds the location the location of ",Object(o.b)("strong",{parentName:"p"},"Default Part Template"),"."),Object(o.b)("p",null,"In 2nd line of above example. we assign value to our newly define ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),Object(o.b)("p",null,"We assign the value by using a ",Object(o.b)("em",{parentName:"p"},"Method")," named ",Object(o.b)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". "),Object(o.b)("p",null,"This method is a part of our main Solidworks variable ",Object(o.b)("inlineCode",{parentName:"p"},"swApp"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),Object(o.b)("p",null,"In this line, we set the value of our ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," variable to new document."),Object(o.b)("p",null,"For ",Object(o.b)("strong",{parentName:"p"},"more detailed information")," about above lines please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post. "),Object(o.b)("p",null,"I have discussed them ",Object(o.b)("strong",{parentName:"p"},"thoroghly")," in ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post, so do checkout this post if you don't understand above code."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Boolean Variable\nDim BoolStatus As Boolean\n\n\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),Object(o.b)("p",null,"In 1st line, we create a variable named ",Object(o.b)("inlineCode",{parentName:"p"},"BoolStatus")," as ",Object(o.b)("inlineCode",{parentName:"p"},"Boolean")," object."),Object(o.b)("p",null,"In next line, we select the ",Object(o.b)("em",{parentName:"p"},"front plane")," by using ",Object(o.b)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",Object(o.b)("inlineCode",{parentName:"p"},"Extension")," object."),Object(o.b)("p",null,"For more information about selection method please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-selection-methods"}),"Solidworks Macros - Selection Methods"))," post."),Object(o.b)("p",null,"I have discussed about different ",Object(o.b)("em",{parentName:"p"},"Selection methods")," in details in ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-selection-methods"}),"Soldworks Macros - Selection Methods"))," post, so do visit this post for more ",Object(o.b)("em",{parentName:"p"},"Selection methods"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n")),Object(o.b)("p",null,"In above line, we create variable ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," for ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),Object(o.b)("p",null,"As the name suggested, a ",Object(o.b)("strong",{parentName:"p"},"Sketch Manager")," holds variours methods and properties to manage ",Object(o.b)("em",{parentName:"p"},"Sketches"),"."),Object(o.b)("p",null,"To see methods and properties related to SketchManager object, please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"https://help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"}),"this page")),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Sketch manager for our sketch\nSet swSketchManager = swDoc.SketchManager\n")),Object(o.b)("p",null,"In above line, we set the ",Object(o.b)("strong",{parentName:"p"},"Sketch manager")," variable to current document's sketch manager."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Inserting a sketch into selected plane\nswSketchManager.InsertSketch True\n")),Object(o.b)("p",null,"In above line, we use ",Object(o.b)("inlineCode",{parentName:"p"},"InsertSketch")," method of ",Object(o.b)("em",{parentName:"p"},"SketchManager")," and give ",Object(o.b)("inlineCode",{parentName:"p"},"True")," value."),Object(o.b)("p",null,"This method allows us to insert a sketch in selected plane."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating Variable for Solidworks Slot\nDim mySketchSlot As SketchSlot\n      \n' Creating a Straight slot\nSet mySketchSlot = swSketchManager.CreateSketchSlot(swSketchSlotCreationType_e.swSketchSlotCreationType_line, swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter, 1, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, False)\n")),Object(o.b)("p",null,"In above sample code, we 1st create a variable named ",Object(o.b)("inlineCode",{parentName:"p"},"mySketchSlot")," of type ",Object(o.b)("inlineCode",{parentName:"p"},"SketchSlot"),"."),Object(o.b)("p",null,"In 2nd line, we ",Object(o.b)("strong",{parentName:"p"},"set")," the value of ",Object(o.b)("em",{parentName:"p"},"SketchSlot")," variable ",Object(o.b)("inlineCode",{parentName:"p"},"mySketchSlot"),"."),Object(o.b)("p",null,"We get this value from ",Object(o.b)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method which is inside the ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," variable."),Object(o.b)("p",null,Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," variable is a type of ",Object(o.b)("strong",{parentName:"p"},"SketchManager"),", hence we used ",Object(o.b)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method from ",Object(o.b)("strong",{parentName:"p"},"SketchManager"),"."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method takes following parameters as explained:"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"SlotCreationType")," : ",Object(o.b)("em",{parentName:"p"},"Type of sketch slot")," as defined in ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e"),"."),Object(o.b)("p",null,"  There are 4 Different types of Slots we can create."),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Straight Slot"))," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_line")," or ",Object(o.b)("strong",{parentName:"p"},"0"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Centerpoint straight Slot"))," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_center_line")," or ",Object(o.b)("strong",{parentName:"p"},"1"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Centerpoint arc Slot"))," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_arc")," or ",Object(o.b)("strong",{parentName:"p"},"2"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"3-point arc Slot"))," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_3pointarc")," or ",Object(o.b)("strong",{parentName:"p"},"4")))),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"SlotLengthType")," : ",Object(o.b)("em",{parentName:"p"},"Type of length of sketch slot")," as defined in ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e"),"."),Object(o.b)("p",null,"  There are 2 different types of Sketch slot length we can create."),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Center to Center"))," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter")," or ",Object(o.b)("strong",{parentName:"p"},"0"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Full Length"))," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e.swSketchSlotLengthType_FullLength")," or ",Object(o.b)("strong",{parentName:"p"},"1")))),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Width")," : Width of Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"X1")," : X coordinate of the point 1, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Y1")," : Y coordinate of the point 1, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Z1")," : Z coordinate of the point 1, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"X2")," : X coordinate of the point 2, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Y2")," : Y coordinate of the point 2, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Z2")," : Z coordinate of the point 2, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"X3")," : X coordinate of the point 3, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Y3")," : Y coordinate of the point 3, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Z3")," : Z coordinate of the point 3, of the Slot"),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"CenterArcDirection")," : We need to set the direction eiter Clockwise or Anti-Clockwise/Counterclockwise as follows:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Clockwise (CW)"))," : -1")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Anti-Clockwise/Counterclockwise (CCW)"))," : 1"))),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"AddDimension")," : ",Object(o.b)("inlineCode",{parentName:"p"},"True")," to automatically add dimensions, ",Object(o.b)("inlineCode",{parentName:"p"},"False")," to not."),Object(o.b)("p",null,"For ",Object(o.b)("strong",{parentName:"p"},"more details")," about ",Object(o.b)("em",{parentName:"p"},"Slot Parameter")," you can visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"http://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchmanager~createsketchslot.html"}),"this page")),"."),Object(o.b)("p",null,"For creating a ",Object(o.b)("em",{parentName:"p"},"Straight Slot"),", I used following parameter Values:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"SlotCreationType")," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_line")),Object(o.b)("p",{parentName:"li"},"Since we want to create a ",Object(o.b)("em",{parentName:"p"},"Straight Slot")," hence I select above value.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"SlotLengthType")," : ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter")),Object(o.b)("p",{parentName:"li"},"I want length of this Slot from ",Object(o.b)("em",{parentName:"p"},"Center to Center")," hence I select above value.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Width")," : ",Object(o.b)("strong",{parentName:"p"},"1"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"X1, Y1, Z1")," : ",Object(o.b)("inlineCode",{parentName:"p"},"0, 0, 0")),Object(o.b)("p",{parentName:"li"},"For Point 1, I use (0, 0, 0) values, which is ",Object(o.b)("em",{parentName:"p"},"origin")," of Sketch.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"X2, Y2, Z2")," : ",Object(o.b)("inlineCode",{parentName:"p"},"1, 0, 0")),Object(o.b)("p",{parentName:"li"},"For Point 2, I use (1, 0, 0) values, which is which is 1 point distance in X-direction.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"X3, Y3, Z3")," : ",Object(o.b)("inlineCode",{parentName:"p"},"1, 1, 0")),Object(o.b)("p",{parentName:"li"},"For Point 2, I use (1, 1, 0) values, which is which is 1 point distance in X-direction and 1 point distance in Y-direction.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"CenterArcDirection")," : ",Object(o.b)("strong",{parentName:"p"},"1")),Object(o.b)("p",{parentName:"li"},"I want to create Anti-Clockwise/Counterclockwise Slot.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"AddDimension")," : ",Object(o.b)("inlineCode",{parentName:"p"},"False")))),Object(o.b)("p",null,"Below Image described ",Object(o.b)("strong",{parentName:"p"},"the Parameters for Straight Slot")," in more detail."),Object(o.b)("p",null,Object(o.b)("img",{alt:"straight-slot-parameters",src:n(257).default})),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method returns ",Object(o.b)("em",{parentName:"p"},"Sketch Slot")," interface i.e. ",Object(o.b)("inlineCode",{parentName:"p"},"ISketchSlot")," interface. "),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"ISketchSlot")," interface has various ",Object(o.b)("strong",{parentName:"p"},"methods and properties")," for ",Object(o.b)("em",{parentName:"p"},"a Slot"),"."),Object(o.b)("p",null,"For more detail about ",Object(o.b)("strong",{parentName:"p"},"methods and properties")," of ",Object(o.b)("inlineCode",{parentName:"p"},"ISketchSlot")," interface you can visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSlot_members.html"}),"this page")),"."),Object(o.b)("hr",null),Object(o.b)("h3",{id:"note"},"NOTE"),Object(o.b)("p",null,"It is ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"very important"))," to remember that, when you give distance or any other numeric value in ",Object(o.b)("strong",{parentName:"p"},"Solidworks API"),", Solidworks takes that numeric value in ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Meter only")),"."),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"Solidworks API")," does not care about your application's Unit systems."),Object(o.b)("p",null,'For example, I works in ANSI system means "inches" for distance. '),Object(o.b)("p",null,"But when I used Solidworks API through ",Object(o.b)("em",{parentName:"p"},"VBA macros")," or ",Object(o.b)("em",{parentName:"p"},"C#"),", I have to use ",Object(o.b)("strong",{parentName:"p"},"converted")," numeric values."),Object(o.b)("p",null,"Because Solidworks API output the distance in ",Object(o.b)("strong",{parentName:"p"},"Meter")," only; which is not my requirement."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' De-select the Slot after creation\nswDoc.ClearSelection2 True\n")),Object(o.b)("p",null,"In the this line of code, we de-select the created Straight Slot."),Object(o.b)("p",null,"For de-selecting, we use ",Object(o.b)("inlineCode",{parentName:"p"},"ClearSelection2")," method from our Solidworks document variable ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Zoom to fit screen in Solidworks Window\nswDoc.ViewZoomtofit\n")),Object(o.b)("p",null,"In this last line we use ",Object(o.b)("em",{parentName:"p"},"zoom to fit")," command."),Object(o.b)("p",null,"For Zoom to fit, we use ",Object(o.b)("inlineCode",{parentName:"p"},"ViewZoomtofit")," method from our Solidworks document variable ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),". "),Object(o.b)("p",null,"Hope this post helps you to ",Object(o.b)("em",{parentName:"p"},"create a Straight Slot")," in Sketches with Solidworks VB Macros."),Object(o.b)("p",null,"For more such tutorials on ",Object(o.b)("strong",{parentName:"p"},"Solidworks VBA Macros"),", do come to this blog after sometime."),Object(o.b)("p",null,"Till then, Happy learning!!!"))}i.isMDXComponent=!0},153:function(e,t,n){"use strict";n.d(t,"a",(function(){return s})),n.d(t,"b",(function(){return d}));var a=n(0),r=n.n(a);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function l(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function c(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?l(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):l(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function b(e,t){if(null==e)return{};var n,a,r=function(e,t){if(null==e)return{};var n,a,r={},o=Object.keys(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var p=r.a.createContext({}),i=function(e){var t=r.a.useContext(p),n=t;return e&&(n="function"==typeof e?e(t):c(c({},t),e)),n},s=function(e){var t=i(e.components);return r.a.createElement(p.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return r.a.createElement(r.a.Fragment,{},t)}},O=r.a.forwardRef((function(e,t){var n=e.components,a=e.mdxType,o=e.originalType,l=e.parentName,p=b(e,["components","mdxType","originalType","parentName"]),s=i(n),O=a,d=s["".concat(l,".").concat(O)]||s[O]||m[O]||o;return n?r.a.createElement(d,c(c({ref:t},p),{},{components:n})):r.a.createElement(d,c({ref:t},p))}));function d(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var o=n.length,l=new Array(o);l[0]=O;var c={};for(var b in t)hasOwnProperty.call(t,b)&&(c[b]=t[b]);c.originalType=e,c.mdxType="string"==typeof e?e:a,l[1]=c;for(var p=2;p<o;p++)l[p]=n[p];return r.a.createElement.apply(null,l)}return r.a.createElement.apply(null,n)}O.displayName="MDXCreateElement"},257:function(e,t,n){"use strict";n.r(t),t.default=n.p+"assets/images/straight-slot-parameters-0370f625925d5ddf397ec055c03cb9fc.png"}}]);