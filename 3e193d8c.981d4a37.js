(window.webpackJsonp=window.webpackJsonp||[]).push([[23],{152:function(e,t,n){"use strict";n.d(t,"a",(function(){return p})),n.d(t,"b",(function(){return u}));var a=n(0),r=n.n(a);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function l(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function b(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?l(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):l(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function i(e,t){if(null==e)return{};var n,a,r=function(e,t){if(null==e)return{};var n,a,r={},o=Object.keys(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var s=r.a.createContext({}),c=function(e){var t=r.a.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):b(b({},t),e)),n},p=function(e){var t=c(e.components);return r.a.createElement(s.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return r.a.createElement(r.a.Fragment,{},t)}},d=r.a.forwardRef((function(e,t){var n=e.components,a=e.mdxType,o=e.originalType,l=e.parentName,s=i(e,["components","mdxType","originalType","parentName"]),p=c(n),d=a,u=p["".concat(l,".").concat(d)]||p[d]||m[d]||o;return n?r.a.createElement(u,b(b({ref:t},s),{},{components:n})):r.a.createElement(u,b({ref:t},s))}));function u(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var o=n.length,l=new Array(o);l[0]=d;var b={};for(var i in t)hasOwnProperty.call(t,i)&&(b[i]=t[i]);b.originalType=e,b.mdxType="string"==typeof e?e:a,l[1]=b;for(var s=2;s<o;s++)l[s]=n[s];return r.a.createElement.apply(null,l)}return r.a.createElement.apply(null,n)}d.displayName="MDXCreateElement"},167:function(e,t,n){"use strict";n.r(t),t.default=n.p+"assets/images/fillet_parameters-c914080196123507c1e3e6449604a6b8.png"},80:function(e,t,n){"use strict";n.r(t),n.d(t,"frontMatter",(function(){return l})),n.d(t,"metadata",(function(){return b})),n.d(t,"rightToc",(function(){return i})),n.d(t,"default",(function(){return c}));var a=n(2),r=n(6),o=(n(0),n(152)),l={id:"sw-sketch-macro-create-fillet",title:"Create a Fillet"},b={unversionedId:"solidworks-macros/sw-sketch-macro-create-fillet",id:"solidworks-macros/sw-sketch-macro-create-fillet",isDocsHomePage:!1,title:"Create a Fillet",description:"In this post, I tell you about how to create a Fillet through Solidworks VBA Macros in a sketch.",source:"@site/docs\\solidworks-macros\\2019-10-10-create-fillet.md",slug:"/solidworks-macros/sw-sketch-macro-create-fillet",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-fillet",version:"current",sidebar:"swvba",previous:{title:"Create a Spline",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-spline"},next:{title:"Create a Chamfer",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-chamfer"}},i=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",children:[]},{value:"For Experience Macro Developer",id:"for-experience-macro-developer",children:[]},{value:"For Beginners Macro Developers",id:"for-beginners-macro-developers",children:[{value:"Understanding the Code",id:"understanding-the-code",children:[]},{value:"NOTE",id:"note",children:[]}]},{value:"VBA Language feature used in this post",id:"vba-language-feature-used-in-this-post",children:[]},{value:"Solidworks API Objects",id:"solidworks-api-objects",children:[]}],s={rightToc:i};function c(e){var t=e.components,l=Object(r.a)(e,["components"]);return Object(o.b)("wrapper",Object(a.a)({},s,l,{components:t,mdxType:"MDXLayout"}),Object(o.b)("p",null,"In this post, I tell you about ",Object(o.b)("em",{parentName:"p"},"how to create a Fillet through Solidworks VBA Macros")," in a sketch."),Object(o.b)("p",null,"This post is an extension of ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-sketch-macro-create-corner-rec"}),"Sketch - Create Corner Rectangle"))," post."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),Object(o.b)("p",null,"Please see below video how visually we can create ",Object(o.b)("em",{parentName:"p"},"a Fillet")," from ",Object(o.b)("strong",{parentName:"p"},"Solidworks VBA macro"),"."),Object(o.b)("div",{class:"youtube-responsive-container"},Object(o.b)("iframe",{src:"https://www.youtube.com/embed/IMHM0_QF7HQ",frameborder:"0",allowfullscreen:!0})),Object(o.b)("p",null,"Please note that there are ",Object(o.b)("strong",{parentName:"p"},"no explaination")," given in the video. "),Object(o.b)("p",null,Object(o.b)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"for-experience-macro-developer"},"For Experience Macro Developer"),Object(o.b)("p",null,"If you are an experience ",Object(o.b)("strong",{parentName:"p"},"Solidworks Macro developer"),", then you are looking for a specific code sample."),Object(o.b)("p",null,"Below is the code for creating ",Object(o.b)("strong",{parentName:"p"},"A Fillet")," from ",Object(o.b)("strong",{parentName:"p"},"Solidworks VBA Macro"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks Sketch Segment\nDim swSketchSegment As SldWorks.SketchSegment\n      \n' Set the value of Solidworks Sketch segment by \"CreateFillet\" method from Solidworks sketch manager\nSet swSketchSegment = swSketchManager.CreateFillet(0.1, swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry)\n")),Object(o.b)("p",null,"For creating a ",Object(o.b)("strong",{parentName:"p"},"Fillet")," first you need to ",Object(o.b)("strong",{parentName:"p"},"Create")," a variable of ",Object(o.b)("inlineCode",{parentName:"p"},"SketchSegment")," type."),Object(o.b)("p",null,"After creating variable, you need to set the value of this variable."),Object(o.b)("p",null,"For this you used ",Object(o.b)("inlineCode",{parentName:"p"},"CreateFillet")," method from ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateFillet")," method set the value of ",Object(o.b)("inlineCode",{parentName:"p"},"SketchSegment")," type variable."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateFillet")," method takes following parameters as explained:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"Radius")," : ",Object(o.b)("em",{parentName:"p"},"Radius of the fillet in meters."))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"ConstrainedCorners")," : ",Object(o.b)("em",{parentName:"p"},"Action to take if the corner to be filleted is constrained or has a dimension.")))),Object(o.b)("p",null,"If you want a more detail explaination then please read further otherwise this will help you to ",Object(o.b)("strong",{parentName:"p"},"Create a Fillet From VBA Macro"),"."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"for-beginners-macro-developers"},"For Beginners Macro Developers"),Object(o.b)("p",null,"In this post, I tell you about ",Object(o.b)("inlineCode",{parentName:"p"},"CreateFillet")," method from ",Object(o.b)("strong",{parentName:"p"},"Solidworks")," ",Object(o.b)("inlineCode",{parentName:"p"},"SketchManager")," object."),Object(o.b)("p",null,"This method is ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"most updated"))," method, I found in ",Object(o.b)("em",{parentName:"p"},"Solidworks API Help"),". "),Object(o.b)("p",null,"So ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"use this method"))," if you want to create a Fillet."),Object(o.b)("p",null,"Below is the ",Object(o.b)("inlineCode",{parentName:"p"},"code")," sample for creating a Fillet."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Boolean Variable\nDim BoolStatus As Boolean\n\n' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Creating Variable for Solidworks Sketch Segment\nDim swSketchSegment As SldWorks.SketchSegment\n\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n  \n  ' Creating string type variable for storing default part location\n  Dim defaultTemplate As String\n  ' Setting value of this string type variable to \"Default part template\"\n  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n  ' Setting Solidworks document to new part document\n  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\n  ' Selecting Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n\n  ' Setting Sketch manager for our sketch\n  Set swSketchManager = swDoc.SketchManager\n\n  ' Inserting a sketch into selected plane\n  swSketchManager.InsertSketch True\n  \n  ' Creating a \"Variant\" Variable which holds the values return by \"CreateCornerRectangle\" method\n  Dim vSketchLines As Variant\n  \n  ' Creating a Corner Rectangle\n  vSketchLines = swSketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0)\n  \n  ' De-select the lines after creation\n  swDoc.ClearSelection2 True\n  \n  ' Selecting Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Point1\", \"SKETCHPOINT\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n\n  ' Set the value of Solidworks Sketch segment by \"CreateFillet\" method from Solidworks sketch manager\n  Set swSketchSegment = swSketchManager.CreateFillet(0.1, swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry)\n\n  ' De-select the Fillet after creation\n  swDoc.ClearSelection2 True\n  \n  ' Show Front View after creating Fillet\n  swDoc.ShowNamedView2 \"\", swStandardViews_e.swFrontView\n  \n  ' Zoom to fit screen in Solidworks Window\n  swDoc.ViewZoomtofit2\n\nEnd Sub\n")),Object(o.b)("hr",null),Object(o.b)("h3",{id:"understanding-the-code"},"Understanding the Code"),Object(o.b)("p",null,"Now let us walk through ",Object(o.b)("em",{parentName:"p"},"each line")," in the above code, and ",Object(o.b)("strong",{parentName:"p"},"understand")," the meaning of every line."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n")),Object(o.b)("p",null,"This line forces us to define every variable we are going to use. "),Object(o.b)("p",null,"For more information please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Boolean Variable\nDim BoolStatus As Boolean\n")),Object(o.b)("p",null,"In this line, we create a variable named ",Object(o.b)("inlineCode",{parentName:"p"},"BoolStatus")," as ",Object(o.b)("inlineCode",{parentName:"p"},"Boolean")," object type."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n")),Object(o.b)("p",null,"In above line, we create variable ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," for ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),Object(o.b)("p",null,"As the name suggested, a ",Object(o.b)("strong",{parentName:"p"},"Sketch Manager")," holds variours methods and properties to manage ",Object(o.b)("em",{parentName:"p"},"Sketches"),"."),Object(o.b)("p",null,"To see methods and properties related to ",Object(o.b)("inlineCode",{parentName:"p"},"SketchManager")," object, please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"}),"this page")),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks Sketch Segment\nDim swSketchSegment As SldWorks.SketchSegment\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSegment")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSegment")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.SketchSegment"),"."),Object(o.b)("p",null,"We create variable ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSegment")," for ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Segments"),"."),Object(o.b)("p",null,"To see methods and properties related to ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSegment")," object, please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment_members.html"}),"this page")),"."),Object(o.b)("p",null,"These all are our global variables."),Object(o.b)("p",null,"As you can see in code sample, they are ",Object(o.b)("strong",{parentName:"p"},"Solidworks API Objects"),"."),Object(o.b)("p",null,"So basically I group all the ",Object(o.b)("strong",{parentName:"p"},"Solidworks API Objects")," in one place."),Object(o.b)("p",null,"I have also place ",Object(o.b)("inlineCode",{parentName:"p"},"boolean")," type object at top also, because after certain point we will ",Object(o.b)("em",{parentName:"p"},"need")," this variable frequently."),Object(o.b)("p",null,"Thus, I have started placing it here."),Object(o.b)("p",null,"Next is our ",Object(o.b)("inlineCode",{parentName:"p"},"Sub")," procedure named ",Object(o.b)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",Object(o.b)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),Object(o.b)("p",null,"In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),Object(o.b)("p",null,"In 1st statement of above example, we are defining a variable of ",Object(o.b)("inlineCode",{parentName:"p"},"string")," type and named it as ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),Object(o.b)("p",null,"This variable ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate"),", hold the location the location of ",Object(o.b)("strong",{parentName:"p"},"Default Part Template"),"."),Object(o.b)("p",null,"In 2nd line of above example. we assign value to our newly define ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),Object(o.b)("p",null,"We assign the value by using a ",Object(o.b)("em",{parentName:"p"},"Method")," named ",Object(o.b)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". This method is a part of our main Solidworks variable ",Object(o.b)("inlineCode",{parentName:"p"},"swApp"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),Object(o.b)("p",null,"In this line, we set the value of our ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," variable to new document."),Object(o.b)("p",null,"For ",Object(o.b)("strong",{parentName:"p"},"detailed information")," about these lines please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post."),Object(o.b)("p",null,"I have discussed them ",Object(o.b)("strong",{parentName:"p"},"thoroghly")," in ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post, so do checkout this post if you don't understand above code."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),Object(o.b)("p",null,"In above line, we select the ",Object(o.b)("em",{parentName:"p"},"front plane")," by using ",Object(o.b)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",Object(o.b)("inlineCode",{parentName:"p"},"Extension")," object."),Object(o.b)("p",null,"For more information about selection method please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-selection-methods"}),"Solidworks Macros - Selection Methods"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Sketch manager for our sketch\nSet swSketchManager = swDoc.SketchManager\n")),Object(o.b)("p",null,"In above line, we set the sketch manager variable to current document's sketch manager."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Inserting a sketch into selected plane\nswSketchManager.InsertSketch True\n")),Object(o.b)("p",null,"In above line, we use ",Object(o.b)("inlineCode",{parentName:"p"},"InsertSketch")," method of ",Object(o.b)("em",{parentName:"p"},"SketchManager")," and give ",Object(o.b)("inlineCode",{parentName:"p"},"True")," value."),Object(o.b)("p",null,"This method allows us to insert a sketch in selected plane."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Creating a "Variant" Variable which holds the values return by "CreateCornerRectangle" method\nDim vSketchLines As Variant\n    \n\' Creating a Corner Rectangle\nvSketchLines = swSketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0)\n')),Object(o.b)("p",null,"In above sample code, we 1st create a variable named ",Object(o.b)("inlineCode",{parentName:"p"},"vSketchLines")," of type ",Object(o.b)("inlineCode",{parentName:"p"},"Variant"),"."),Object(o.b)("p",null,"A ",Object(o.b)("inlineCode",{parentName:"p"},"Variant")," type variable can hold ",Object(o.b)("strong",{parentName:"p"},"any")," type of value depends upon the use of variable."),Object(o.b)("p",null,"In 2nd line, we set the value of variable ",Object(o.b)("inlineCode",{parentName:"p"},"vSketchLines"),"."),Object(o.b)("p",null,"Value of ",Object(o.b)("inlineCode",{parentName:"p"},"vSketchLinesis")," an array of lines. This array is send as return value when we use ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method is part of ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," and it is the latest method to create a corner rectangle."),Object(o.b)("p",null,"For detail explaination on ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method, please see ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-sketch-macro-create-corner-rec"}),"Sketch - Create Corner Rectangle"))," post."),Object(o.b)("p",null,"In the above code sample I have used (0, 1, 0) Upper-left point in ",Object(o.b)("em",{parentName:"p"},"Y-direction"),"."),Object(o.b)("p",null,"For Lower-right point I used (1, 0, 0) which is 1 point distance in ",Object(o.b)("em",{parentName:"p"},"X-direction"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' De-select the Rectangle after creation\nswDoc.ClearSelection2 True\n")),Object(o.b)("p",null,"In above line, we de-select the ractangle we just create."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),Object(o.b)("p",null,"In above line, we select the ",Object(o.b)("em",{parentName:"p"},"front plane")," by using ",Object(o.b)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",Object(o.b)("inlineCode",{parentName:"p"},"Extension")," object."),Object(o.b)("p",null,"For more information about selection method please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-selection-methods"}),"Solidworks Macros - Selection Methods"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Set the value of Solidworks Sketch segment by "CreateFillet" method from Solidworks sketch manager\nSet swSketchSegment = swSketchManager.CreateFillet(0.1, swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry)\n')),Object(o.b)("p",null,"In above line, we set the value of Solidworks Sketch Segment variable ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchSegment")," by ",Object(o.b)("inlineCode",{parentName:"p"},"CreateFillet")," method from ",Object(o.b)("em",{parentName:"p"},"Solidworks Sketch Manager"),"."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateFillet")," method takes following parameters:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"Radius")," : ",Object(o.b)("em",{parentName:"p"},"Radius of the fillet in meters."))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"ConstrainedCorners")," : ",Object(o.b)("em",{parentName:"p"},"Action to take if the corner to be filleted is constrained or has a dimension.")))),Object(o.b)("p",null,"Below Image described ",Object(o.b)("strong",{parentName:"p"},"the Parameters for a Fillet"),"."),Object(o.b)("p",null,Object(o.b)("img",{alt:"fillet_parameters",src:n(167).default})),Object(o.b)("p",null,"In our code, I have used following values:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"Radius")," : I have used 0.1 (This value is in meter) as the radius of fillet.")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"ConstrainedCorners")," : I have used ",Object(o.b)("inlineCode",{parentName:"p"},"swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry")," enumerator as value for constraining corners."))),Object(o.b)("p",null,"In ",Object(o.b)("strong",{parentName:"p"},"swConstrainedCornerAction_e")," we have 4 constant values."),Object(o.b)("p",null,"These values are as follows:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"swConstrainedCornerDeleteGeometry")," : 2 = Delete the constraint or dimension and add the fillet")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"swConstrainedCornerInteract")," : 0 = Ask the user whether to delete the geometry or stop processing")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"swConstrainedCornerKeepGeometry")," : 1 = Keep the constraint or dimension by creating a virtual intersection point before adding the fillet")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"swConstrainedCornerStopProcessing")," : 3 = Do not delete the constrain or dimension and do not create the fillet."))),Object(o.b)("hr",null),Object(o.b)("h3",{id:"note"},"NOTE"),Object(o.b)("p",null,"It is ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"very important"))," to remember that, when you give distance or any other numeric value in ",Object(o.b)("strong",{parentName:"p"},"Solidworks API"),", Solidworks takes that numeric value in ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Meter only")),"."),Object(o.b)("p",null,"Solidworks API does not care about your application's Unit systems."),Object(o.b)("p",null,"For example, I works in ANSI system means inches for distance. But when I used Solidworks API through VBA macros or C#, I need to use converted numeric values."),Object(o.b)("p",null,"Because Solidworks API output the distance in ",Object(o.b)("strong",{parentName:"p"},"Meter")," which is not my requirement."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' De-select the Fillet after creation\nswDoc.ClearSelection2 True\n")),Object(o.b)("p",null,"In the above line of code, we deselect the ",Object(o.b)("strong",{parentName:"p"},"Fillet")," we have created."),Object(o.b)("p",null,"For de-selecting, we use ",Object(o.b)("inlineCode",{parentName:"p"},"ClearSelection2")," method from our Solidworks document name ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Show Front View after creating Fillet\nswDoc.ShowNamedView2 "", swStandardViews_e.swFrontView\n')),Object(o.b)("p",null,"In the above line of code, we update the ",Object(o.b)("em",{parentName:"p"},"view orientation")," to ",Object(o.b)("strong",{parentName:"p"},"Front View"),"."),Object(o.b)("p",null,"In my machine, after inserting a sketch view orientation does not changed."),Object(o.b)("p",null,"Because of this I have to update the view to ",Object(o.b)("strong",{parentName:"p"},"Front view"),"."),Object(o.b)("p",null,"For showing ",Object(o.b)("strong",{parentName:"p"},"Front View")," we used ",Object(o.b)("inlineCode",{parentName:"p"},"ShowNamedView2")," method from our Solidworks document name ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),"."),Object(o.b)("p",null,"This method takes 2 parameter described as follows:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"VName")," : Name of the view to display or an empty string to use ViewId instead")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("strong",{parentName:"p"},"ViewId")," : ID of the view to display as defined by ",Object(o.b)("inlineCode",{parentName:"p"},"swStandardViews_e")," or -1 to use the ",Object(o.b)("strong",{parentName:"p"},"VName")," argument instead."))),Object(o.b)("p",null,Object(o.b)("em",{parentName:"p"},"NOTE:")," If you specify both ",Object(o.b)("strong",{parentName:"p"},"VName")," and ",Object(o.b)("strong",{parentName:"p"},"ViewId"),", then ",Object(o.b)("strong",{parentName:"p"},"ViewId")," takes precedence if the two arguments do not resolve to the same view."),Object(o.b)("p",null,Object(o.b)("inlineCode",{parentName:"p"},"swStandardViews_e")," has following Standard View Types:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swBackView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swBottomView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swDimetricView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swFrontView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swIsometricView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swLeftView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swRightView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swTopView"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"swTrimetricView")))),Object(o.b)("p",null,"In our code, we did not use ",Object(o.b)("strong",{parentName:"p"},"VName")," instead I used empty string in form of ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},'""'))," symbol."),Object(o.b)("p",null,"I used ViewId value to specify view and used ",Object(o.b)("inlineCode",{parentName:"p"},"swStandardViews_e.swFrontView")," value to use ",Object(o.b)("em",{parentName:"p"},"Standard Front View"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Zoom to fit screen in Solidworks Window\nswDoc.ViewZoomtofit\n")),Object(o.b)("p",null,"In this last line we use ",Object(o.b)("em",{parentName:"p"},"zoom to fit")," command."),Object(o.b)("p",null,"For Zoom to fit, we use ",Object(o.b)("inlineCode",{parentName:"p"},"ViewZoomtofit")," method from our Solidworks document variable ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),"."),Object(o.b)("p",null,"This is it !!!"),Object(o.b)("p",null,"If you found anything to add or update, please let me know on my e-mail."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"vba-language-feature-used-in-this-post"},"VBA Language feature used in this post"),Object(o.b)("p",null,"In this post used some features of ",Object(o.b)("strong",{parentName:"p"},"VBA programming language"),"."),Object(o.b)("p",null,"This section of post, has some brief information about the VBA programming language specific features."),Object(o.b)("ol",null,Object(o.b)("li",{parentName:"ol"},"We use ",Object(o.b)("strong",{parentName:"li"},"Option Explicit")," for capturing un-declared variables.")),Object(o.b)("p",null,"If you want to read more about ",Object(o.b)("strong",{parentName:"p"},"Option Explicit")," then please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-variables-decl"}),"Declaring and Scoping of Variables")),"."),Object(o.b)("ol",{start:2},Object(o.b)("li",{parentName:"ol"},"Then we create ",Object(o.b)("strong",{parentName:"li"},"variable")," for different data types.")),Object(o.b)("p",null,"If you don't know about them, then please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-variables"}),"Variables"))," and ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-prog-concept#data-types-in-vba"}),"Data-types"))," posts of this blog."),Object(o.b)("p",null,"These posts will help you to understand what ",Object(o.b)("strong",{parentName:"p"},"Variables")," are and how to use them."),Object(o.b)("ol",{start:3},Object(o.b)("li",{parentName:"ol"},"Then we create ",Object(o.b)("strong",{parentName:"li"},"main Sub procedure")," for our macro.")),Object(o.b)("p",null,"If you don't know about the ",Object(o.b)("strong",{parentName:"p"},"Sub procedure"),", then I suggest you to visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-procedures"}),"VBA Sub and Function Procedures"))," and ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-procedures-exec"}),"Executing Sub and Function Procedures"))," posts of this blog."),Object(o.b)("p",null,"These posts will help you to understand what ",Object(o.b)("strong",{parentName:"p"},"Procedures")," are and how to use them."),Object(o.b)("ol",{start:4},Object(o.b)("li",{parentName:"ol"},"In most part we create some variables and set their values. We set those values by using some ",Object(o.b)("strong",{parentName:"li"},"functions")," provided from objects.")),Object(o.b)("p",null,"If you don't know about the ",Object(o.b)("strong",{parentName:"p"},"functions"),", then you should visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-functions"}),"VBA Functions"))," and ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"../vba/vba-functions-adv"}),"VBA Functions that do more"))," posts of this blog."),Object(o.b)("p",null,"These posts will help you to understand what ",Object(o.b)("strong",{parentName:"p"},"functions")," are and how to use them."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"solidworks-api-objects"},"Solidworks API Objects"),Object(o.b)("p",null,"In this post, for creating a ",Object(o.b)("strong",{parentName:"p"},"Fillet"),", we use ",Object(o.b)("em",{parentName:"p"},"Solidworks API objects and their methods"),"."),Object(o.b)("p",null,"This section contains the list of all ",Object(o.b)("strong",{parentName:"p"},"Solidworks Objects")," used in this post."),Object(o.b)("p",null,"I have also attached links of these ",Object(o.b)("strong",{parentName:"p"},"Solidworks API Objects")," in ",Object(o.b)("strong",{parentName:"p"},"API Help website"),"."),Object(o.b)("p",null,"If you want to explore those objects, you can use these links."),Object(o.b)("p",null,"These Solidworks API Objects are listed below:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("strong",{parentName:"li"},"Solidworks Application Object"))),Object(o.b)("p",null,"If you want explore ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Properties and Methods/Functions"))," of ",Object(o.b)("strong",{parentName:"p"},"Solidworks Application Object")," object you can visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html"}),"this link")),"."),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("strong",{parentName:"li"},"Solidworks Document Object"))),Object(o.b)("p",null,"If you want explore ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Properties and Methods/Functions"))," of ",Object(o.b)("strong",{parentName:"p"},"Solidworks Document Object")," object you can visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html"}),"this link")),"."),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("strong",{parentName:"li"},"Solidworks Sketch Manager Object"))),Object(o.b)("p",null,"If you want explore ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Properties and Methods/Functions"))," of ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Manager Object")," you can visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"}),"this link")),"."),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("strong",{parentName:"li"},"Solidworks Sketch Segment Object"))),Object(o.b)("p",null,"If you want explore ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Properties and Methods/Functions"))," of ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Segment Object")," you can visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSegment_members.html"}),"this link")),"."),Object(o.b)("hr",null),Object(o.b)("p",null,"Hope this post helps you to ",Object(o.b)("em",{parentName:"p"},"create a Fillet")," in Sketches with Solidworks VB Macros."),Object(o.b)("p",null,"For more such tutorials on ",Object(o.b)("strong",{parentName:"p"},"Solidworks VBA Macros"),", do come to this blog after sometime."),Object(o.b)("p",null,"Till then, Happy learning!!!"))}c.isMDXComponent=!0}}]);