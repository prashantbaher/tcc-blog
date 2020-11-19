(window.webpackJsonp=window.webpackJsonp||[]).push([[59],{116:function(e,t,n){"use strict";n.r(t),n.d(t,"frontMatter",(function(){return c})),n.d(t,"metadata",(function(){return l})),n.d(t,"rightToc",(function(){return i})),n.d(t,"default",(function(){return s}));var a=n(2),r=n(6),o=(n(0),n(152)),c={id:"sw-sketch-macro-create-corner-rec",title:"Create Corner Rectangle"},l={unversionedId:"solidworks-macros/sw-sketch-macro-create-corner-rec",id:"solidworks-macros/sw-sketch-macro-create-corner-rec",isDocsHomePage:!1,title:"Create Corner Rectangle",description:"In this post, I tell you about how to create Corner Rectangle through Solidworks VBA Macros in a sketch.",source:"@site/docs\\solidworks-macros\\2019-04-16-create-corner-rectangle.md",slug:"/solidworks-macros/sw-sketch-macro-create-corner-rec",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-corner-rec",version:"current",sidebar:"swvba",previous:{title:"Create CenterLine",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-centerline"},next:{title:"Create Center Rectangle",permalink:"/docs/solidworks-macros/sw-sketch-macro-create-center-rec"}},i=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",children:[]},{value:"Code Sample",id:"code-sample",children:[{value:"Understanding the Code",id:"understanding-the-code",children:[]},{value:"NOTE",id:"note",children:[]}]}],b={rightToc:i};function s(e){var t=e.components,c=Object(r.a)(e,["components"]);return Object(o.b)("wrapper",Object(a.a)({},b,c,{components:t,mdxType:"MDXLayout"}),Object(o.b)("p",null,"In this post, I tell you about ",Object(o.b)("em",{parentName:"p"},"how to create Corner Rectangle through Solidworks VBA Macros")," in a sketch."),Object(o.b)("p",null,"For this, I take the example from previous ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-sketch-macro-create-line"}),"Sketch - Create Lines"))," post."),Object(o.b)("p",null,"In this post, I tell you about ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method from ",Object(o.b)("strong",{parentName:"p"},"Solidworks")," ",Object(o.b)("inlineCode",{parentName:"p"},"SketchManager")," object."),Object(o.b)("p",null,"This method is ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"most updated"))," method, I found in ",Object(o.b)("em",{parentName:"p"},"Solidworks API Help"),". "),Object(o.b)("p",null,"So ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"use this method"))," if you want to create a new ",Object(o.b)("strong",{parentName:"p"},"Corner Rectangle"),"."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),Object(o.b)("p",null,"Please see below video on ",Object(o.b)("strong",{parentName:"p"},"how to create Corner Rectangle")," from Solidworks VBA Macros."),Object(o.b)("div",{class:"youtube-responsive-container"},Object(o.b)("iframe",{src:"https://www.youtube.com/embed/03s3pWNIC08",frameborder:"0",allowfullscreen:!0})),Object(o.b)("p",null,"Please note that there are ",Object(o.b)("strong",{parentName:"p"},"no explaination")," in the video. "),Object(o.b)("p",null,Object(o.b)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),Object(o.b)("hr",null),Object(o.b)("h2",{id:"code-sample"},"Code Sample"),Object(o.b)("p",null,"Below is the ",Object(o.b)("inlineCode",{parentName:"p"},"code")," sample for creating Corner Rectangle."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n  \n  ' Creating string type variable for storing default part location\n  Dim defaultTemplate As String\n  ' Setting value of this string type variable to \"Default part template\"\n  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n  ' Setting Solidworks document to new part document\n  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\n  ' Selecting Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  \n  ' Setting Sketch manager for our sketch\n  Set swSketchManager = swDoc.SketchManager\n  \n  ' Creating a \"Variant\" Variable which holds the values return by \"CreateCornerRectangle\" method\n  Dim vSketchLines As Variant\n  \n  ' Inserting a sketch into selected plane\n  swSketchManager.InsertSketch True\n  \n  ' Creating a Corner Rectangle\n  vSketchLines = swSketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0)\n  \n  ' De-select the lines after creation\n  swDoc.ClearSelection2 True\n  \n  ' Zoom to fit screen in Solidworks Window\n  swDoc.ViewZoomtofit2\n\nEnd Sub\n")),Object(o.b)("hr",null),Object(o.b)("h3",{id:"understanding-the-code"},"Understanding the Code"),Object(o.b)("p",null,"Now let us walk through ",Object(o.b)("em",{parentName:"p"},"each line")," in the above code, and ",Object(o.b)("strong",{parentName:"p"},"understand")," the meaning of every line."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n")),Object(o.b)("p",null,"This line forces us to define every variable we are going to use. "),Object(o.b)("p",null,"For more information please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swApp")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),Object(o.b)("p",null,"In this line, we are creating a variable which we named as ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," variable is ",Object(o.b)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),Object(o.b)("p",null,"Next is our ",Object(o.b)("inlineCode",{parentName:"p"},"Sub")," procedure named ",Object(o.b)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",Object(o.b)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),Object(o.b)("p",null,"In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),Object(o.b)("p",null,"In 1st statement of above example, we are defining a variable of ",Object(o.b)("inlineCode",{parentName:"p"},"string")," type and named it as ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),Object(o.b)("p",null,"This variable ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate"),", hold the location the location of ",Object(o.b)("strong",{parentName:"p"},"Default Part Template"),"."),Object(o.b)("p",null,"In 2nd line of above example. we assign value to our newly define ",Object(o.b)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),Object(o.b)("p",null,"We assign the value by using a ",Object(o.b)("em",{parentName:"p"},"Method")," named ",Object(o.b)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". This method is a part of our main Solidworks variable ",Object(o.b)("inlineCode",{parentName:"p"},"swApp"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),Object(o.b)("p",null,"In this line, we set the value of our ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc")," variable to new document."),Object(o.b)("p",null,"For ",Object(o.b)("strong",{parentName:"p"},"detailed information")," about these lines please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-open-part"}),"Solidworks Macros - Open new Part document"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Boolean Variable\nDim BoolStatus As Boolean\n\n\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),Object(o.b)("p",null,"In 1st line, we create a variable named ",Object(o.b)("inlineCode",{parentName:"p"},"BoolStatus")," as ",Object(o.b)("inlineCode",{parentName:"p"},"Boolean")," object."),Object(o.b)("p",null,"In next line, we select the ",Object(o.b)("em",{parentName:"p"},"front plane")," by using ",Object(o.b)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",Object(o.b)("inlineCode",{parentName:"p"},"Extension")," object."),Object(o.b)("p",null,"For more information about selection method please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"sw-macro-selection-methods"}),"Solidworks Macros - Selection Methods"))," post."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n")),Object(o.b)("p",null,"In above line, we create variable ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," for ",Object(o.b)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),Object(o.b)("p",null,"As the name suggested, a Sketch Manager holds variours methods and properties to manage Sketches."),Object(o.b)("p",null,"To see methods and properties related to SketchManager object, please visit ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(a.a)({parentName:"strong"},{href:"https://help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"}),"this page")),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Setting Sketch manager for our sketch\nSet swSketchManager = swDoc.SketchManager\n")),Object(o.b)("p",null,"In above line, we set the sketch manager variable to current document's sketch manager."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Inserting a sketch into selected plane\nswSketchManager.InsertSketch True\n")),Object(o.b)("p",null,"In above line, we use ",Object(o.b)("inlineCode",{parentName:"p"},"InsertSketch")," method of ",Object(o.b)("em",{parentName:"p"},"SketchManager")," and give ",Object(o.b)("inlineCode",{parentName:"p"},"True")," value."),Object(o.b)("p",null,"This method allows us to insert a sketch in selected plane."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),'\' Creating a "Variant" Variable which holds the values return by "CreateCornerRectangle" method\nDim vSketchLines As Variant\n    \n\' Creating a Corner Rectangle\nvSketchLines = swSketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0)\n')),Object(o.b)("p",null,"In above sample code, we 1st create a variable named ",Object(o.b)("inlineCode",{parentName:"p"},"vSketchLines")," of type ",Object(o.b)("inlineCode",{parentName:"p"},"Variant"),"."),Object(o.b)("p",null,"A ",Object(o.b)("inlineCode",{parentName:"p"},"Variant")," type variable can hold ",Object(o.b)("strong",{parentName:"p"},"any")," type of value depends upon the use of variable."),Object(o.b)("p",null,"In 2nd line, we set the value of variable ",Object(o.b)("inlineCode",{parentName:"p"},"vSketchLines"),"."),Object(o.b)("p",null,"Value of ",Object(o.b)("inlineCode",{parentName:"p"},"vSketchLinesis")," an array of lines. This array is send as return value when we use ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method is part of ",Object(o.b)("inlineCode",{parentName:"p"},"swSketchManager")," and it is the latest method to create a corner rectangle."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method takes following parameters as explained:"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"X1")," : X coordinate of the Upper-left point")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Y1")," : Y coordinate of the Upper-left point")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Z1")," : Z coordinate of the Upper-left point")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"X2")," : X coordinate of the Lower-right point")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Y2")," : Y coordinate of the Lower-right point")),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Z2")," : Z coordinate of the Lower-right point"))),Object(o.b)("p",null,"Below image shows more clearly about these parameters."),Object(o.b)("p",null,Object(o.b)("img",{alt:"corner-rectangle-parameter",src:n(270).default})),Object(o.b)("p",null,"In the above code sample I have used (0, 1, 0) Upper-left point in ",Object(o.b)("em",{parentName:"p"},"Y-direction"),"."),Object(o.b)("p",null,"For Lower-right point I used (1, 0, 0) which is 1 point distance in ",Object(o.b)("em",{parentName:"p"},"X-direction"),"."),Object(o.b)("p",null,"This ",Object(o.b)("inlineCode",{parentName:"p"},"CreateCornerRectangle")," method returns ",Object(o.b)("strong",{parentName:"p"},"an array")," of ",Object(o.b)("em",{parentName:"p"},"sketch segments")," that represent the edges created for this corner rectangle."),Object(o.b)("p",null,"A ",Object(o.b)("em",{parentName:"p"},"Sketch Segment")," can represent a sketch arc, line, ellipse, parabola or spline."),Object(o.b)("p",null,"Sketch Segment has ",Object(o.b)("inlineCode",{parentName:"p"},"ISketchSegment")," Interface, which provides functions that are generic to every type of sketch segment."),Object(o.b)("p",null,"For example, every sketch segment has an ID and can be programmatically selected."),Object(o.b)("p",null,"Therefore, the ",Object(o.b)("inlineCode",{parentName:"p"},"ISketchSegment")," interface provides functions to obtain the ID and to select the item."),Object(o.b)("hr",null),Object(o.b)("h3",{id:"note"},"NOTE"),Object(o.b)("p",null,"It is ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"very important"))," to remember that, when you give distance or any other numeric value in ",Object(o.b)("strong",{parentName:"p"},"Solidworks API"),", Solidworks takes that numeric value in ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Meter only")),"."),Object(o.b)("p",null,"Solidworks API does not care about your application's Unit systems."),Object(o.b)("p",null,"For example, I works in ANSI system means inches for distance. But when I used Solidworks API through VBA macros or C#, I need to use converted numeric values."),Object(o.b)("p",null,"Because Solidworks API output the distance in ",Object(o.b)("strong",{parentName:"p"},"Meter")," which is not my requirement."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' De-select the lines after creation\nswDoc.ClearSelection2 True\n")),Object(o.b)("p",null,"In the this line of code, we deselect the Corner rectangle we have created."),Object(o.b)("p",null,"For de-selecting, we use ",Object(o.b)("inlineCode",{parentName:"p"},"ClearSelection2")," method from our Solidworks document name ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),"."),Object(o.b)("pre",null,Object(o.b)("code",Object(a.a)({parentName:"pre"},{className:"language-vb"}),"' Zoom to fit screen in Solidworks Window\nswDoc.ViewZoomtofit2\n")),Object(o.b)("p",null,"In this last line we use ",Object(o.b)("em",{parentName:"p"},"zoom to fit")," command."),Object(o.b)("p",null,"For Zoom to fit, we use ",Object(o.b)("inlineCode",{parentName:"p"},"ViewZoomtofit")," method from our Solidworks document variable ",Object(o.b)("inlineCode",{parentName:"p"},"swDoc"),"."),Object(o.b)("p",null,"Hope this post helps you to ",Object(o.b)("em",{parentName:"p"},"create Corner rectangle")," in Sketches with Solidworks VB Macros."),Object(o.b)("p",null,"For more such tutorials on ",Object(o.b)("strong",{parentName:"p"},"Solidworks VBA Macros"),", do come to this blog after sometime."),Object(o.b)("p",null,"Till then, Happy learning!!!"))}s.isMDXComponent=!0},152:function(e,t,n){"use strict";n.d(t,"a",(function(){return p})),n.d(t,"b",(function(){return u}));var a=n(0),r=n.n(a);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function c(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?c(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):c(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function i(e,t){if(null==e)return{};var n,a,r=function(e,t){if(null==e)return{};var n,a,r={},o=Object.keys(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var b=r.a.createContext({}),s=function(e){var t=r.a.useContext(b),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},p=function(e){var t=s(e.components);return r.a.createElement(b.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return r.a.createElement(r.a.Fragment,{},t)}},d=r.a.forwardRef((function(e,t){var n=e.components,a=e.mdxType,o=e.originalType,c=e.parentName,b=i(e,["components","mdxType","originalType","parentName"]),p=s(n),d=a,u=p["".concat(c,".").concat(d)]||p[d]||m[d]||o;return n?r.a.createElement(u,l(l({ref:t},b),{},{components:n})):r.a.createElement(u,l({ref:t},b))}));function u(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var o=n.length,c=new Array(o);c[0]=d;var l={};for(var i in t)hasOwnProperty.call(t,i)&&(l[i]=t[i]);l.originalType=e,l.mdxType="string"==typeof e?e:a,c[1]=l;for(var b=2;b<o;b++)c[b]=n[b];return r.a.createElement.apply(null,c)}return r.a.createElement.apply(null,n)}d.displayName="MDXCreateElement"},270:function(e,t,n){"use strict";n.r(t),t.default=n.p+"assets/images/corner-rectangle-parameter-473cb83387b85759e0ca3bfa8da4a6e6.png"}}]);