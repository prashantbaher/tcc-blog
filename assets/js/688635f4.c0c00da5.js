"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[8069],{3905:(e,t,n)=>{n.d(t,{Zo:()=>m,kt:()=>u});var a=n(67294);function r(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function o(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?o(Object(n),!0).forEach((function(t){r(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):o(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,a,r=function(e,t){if(null==e)return{};var n,a,r={},o=Object.keys(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var s=a.createContext({}),p=function(e){var t=a.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},m=function(e){var t=p(e.components);return a.createElement(s.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},d=a.forwardRef((function(e,t){var n=e.components,r=e.mdxType,o=e.originalType,s=e.parentName,m=l(e,["components","mdxType","originalType","parentName"]),d=p(n),u=r,k=d["".concat(s,".").concat(u)]||d[u]||c[u]||o;return n?a.createElement(k,i(i({ref:t},m),{},{components:n})):a.createElement(k,i({ref:t},m))}));function u(e,t){var n=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var o=n.length,i=new Array(o);i[0]=d;var l={};for(var s in t)hasOwnProperty.call(t,s)&&(l[s]=t[s]);l.originalType=e,l.mdxType="string"==typeof e?e:r,i[1]=l;for(var p=2;p<o;p++)i[p]=n[p];return a.createElement.apply(null,i)}return a.createElement.apply(null,n)}d.displayName="MDXCreateElement"},93043:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>p,contentTitle:()=>l,default:()=>d,frontMatter:()=>i,metadata:()=>s,toc:()=>m});var a=n(87462),r=(n(67294),n(3905)),o=n(74753);const i={categories:"Solidworks-macro",title:"Solidworks Macro - Create Center Rectangle",permalink:"/solidworks-macros/create-center-rectangle/",tags:["Solidworks Macro"],id:"create-center-rectangle"},l=void 0,s={unversionedId:"create-center-rectangle",id:"create-center-rectangle",title:"Solidworks Macro - Create Center Rectangle",description:"",source:"@site/docs/solidworks-macros/003.2-create-center-rectangle.md",sourceDirName:".",slug:"/create-center-rectangle",permalink:"/solidworks-macros/create-center-rectangle",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks Macro - Create Center Rectangle",permalink:"/solidworks-macros/create-center-rectangle/",tags:["Solidworks Macro"],id:"create-center-rectangle"},sidebar:"tutorialSidebar",previous:{title:"Solidworks Macro - Create Corner Rectangle",permalink:"/solidworks-macros/create-corner-rectangle"},next:{title:"Solidworks Macro - Create 3-Point Corner Rectangle",permalink:"/solidworks-macros/create-3point-corner-rectangle"}},p={},m=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"Code Sample",id:"code-sample",level:2},{value:"Understanding the Code",id:"understanding-the-code",level:2}],c={toc:m};function d(e){let{components:t,...i}=e;return(0,r.kt)("wrapper",(0,a.Z)({},c,i,{components:t,mdxType:"MDXLayout"}),(0,r.kt)(o.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,"In this post, I tell you about ",(0,r.kt)("em",{parentName:"p"},"how to create Center Rectangle through Solidworks VBA Macros")," in a sketch."),(0,r.kt)("p",null,"The process is almost identical with previous \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/create-corner-rectangle"},"Solidworks Macros - Create Corner Rectangle From VBA Macro"))," post."),(0,r.kt)("p",null,"In this post, I tell you about ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateCenterRectangle")," method from ",(0,r.kt)("strong",{parentName:"p"},"Solidworks")," ",(0,r.kt)("inlineCode",{parentName:"p"},"SketchManager")," object."),(0,r.kt)("p",null,"This method is ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"most updated"))," method, I found in ",(0,r.kt)("em",{parentName:"p"},"Solidworks API Help"),". "),(0,r.kt)("p",null,"So ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"use this method"))," if you want to create a ",(0,r.kt)("em",{parentName:"p"},"Center Rectangle"),"."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,r.kt)("p",null,"Please see below \ud83c\udfac video on ",(0,r.kt)("strong",{parentName:"p"},"how to create Center Rectangle")," from Solidworks VBA Macros."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/J_CjUAN4JOc",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,"Please note that there are ",(0,r.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"code-sample"},"Code Sample"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("inlineCode",{parentName:"p"},"code")," sample for creating a ",(0,r.kt)("em",{parentName:"p"},"Center Rectangle"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n  \n  ' Creating string type variable for storing default part location\n  Dim defaultTemplate As String\n  ' Setting value of this string type variable to \"Default part template\"\n  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n  ' Setting Solidworks document to new part document\n  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\n  ' Selecting Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  \n  ' Setting Sketch manager for our sketch\n  Set swSketchManager = swDoc.SketchManager\n  \n  ' Creating a \"Variant\" Variable which holds the values return by \"CreateCenterRectangle\" method\n  Dim vSketchLines As Variant\n  \n  ' Inserting a sketch into selected plane\n  swSketchManager.InsertSketch True\n  \n  ' Creating a Center Rectangle\n  vSketchLines = swSketchManager.CreateCenterRectangle(0, 0, 0, 1, 1, 0)\n  \n  ' De-select the corner rectangle after creation\n  swDoc.ClearSelection2 True\n  \n  ' Zoom to fit screen in Solidworks Window\n  swDoc.ViewZoomtofit\n\nEnd Sub\n")),(0,r.kt)("hr",null),(0,r.kt)(o.Z,{mdxType:"AdComponent"}),(0,r.kt)("h2",{id:"understanding-the-code"},"Understanding the Code"),(0,r.kt)("p",null,"Now let us walk through ",(0,r.kt)("em",{parentName:"p"},"each line")," in the above code, and ",(0,r.kt)("strong",{parentName:"p"},"understand")," the meaning of every line."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,r.kt)("p",null,"This line forces us to define every variable we are going to use. "),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"For more information please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,r.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp")," variable is ",(0,r.kt)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,r.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc")," variable is ",(0,r.kt)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),(0,r.kt)("p",null,"Next is our ",(0,r.kt)("inlineCode",{parentName:"p"},"Sub")," procedure named ",(0,r.kt)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",(0,r.kt)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),(0,r.kt)("p",null,"In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),(0,r.kt)("p",null,"In 1st statement of above example, we are defining a variable of ",(0,r.kt)("inlineCode",{parentName:"p"},"string")," type and named it as ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),(0,r.kt)("p",null,"This variable ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),", hold the location the location of ",(0,r.kt)("strong",{parentName:"p"},"Default Part Template"),"."),(0,r.kt)("p",null,"In 2nd line of above example. we assign value to our newly define ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),(0,r.kt)("p",null,"We assign the value by using a ",(0,r.kt)("em",{parentName:"p"},"Method")," named ",(0,r.kt)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". This method is a part of our main Solidworks variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),(0,r.kt)("p",null,"In this line, we set the value of our ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc")," variable to new document."),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"For ",(0,r.kt)("strong",{parentName:"p"},"detailed information")," about these lines please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Boolean Variable\nDim BoolStatus As Boolean\n\n\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),(0,r.kt)("p",null,"In 1st line, we create a variable named ",(0,r.kt)("inlineCode",{parentName:"p"},"BoolStatus")," as ",(0,r.kt)("inlineCode",{parentName:"p"},"Boolean")," object."),(0,r.kt)("p",null,"In next line, we select the ",(0,r.kt)("em",{parentName:"p"},"front plane")," by using ",(0,r.kt)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",(0,r.kt)("inlineCode",{parentName:"p"},"Extension")," object."),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"For more information about selection method please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/select-plane-from-tree"},"Solidworks Macros - Selection Methods"))," post.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n")),(0,r.kt)("p",null,"In above line, we create variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchManager")," for ",(0,r.kt)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),(0,r.kt)("p",null,"As the name suggested, a Sketch Manager holds variours methods and properties to manage Sketches."),(0,r.kt)("p",null,"To see methods and properties related to SketchManager object, please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"},"this page of Solidworks API"))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Sketch manager for our sketch\nSet swSketchManager = swDoc.SketchManager\n")),(0,r.kt)("p",null,"In above line, we set the sketch manager variable to current document's sketch manager."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Inserting a sketch into selected plane\nswSketchManager.InsertSketch True\n")),(0,r.kt)("p",null,"In above line, we use ",(0,r.kt)("inlineCode",{parentName:"p"},"InsertSketch")," method of ",(0,r.kt)("em",{parentName:"p"},"SketchManager")," and give ",(0,r.kt)("inlineCode",{parentName:"p"},"True")," value."),(0,r.kt)("p",null,"This method allows us to insert a sketch in selected plane."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Creating a "Variant" Variable which holds the values return by "CreateCenterRectangle" method\nDim vSketchLines As Variant\n    \n\' Creating a Center rectangle\nvSketchLines = swSketchManager.CreateCenterRectangle (0, 0, 0, 1, 1, 0)\n')),(0,r.kt)("p",null,"In above sample code, we 1st create a variable named ",(0,r.kt)("inlineCode",{parentName:"p"},"vSketchLines")," of type ",(0,r.kt)("inlineCode",{parentName:"p"},"Variant"),"."),(0,r.kt)("p",null,"A ",(0,r.kt)("inlineCode",{parentName:"p"},"Variant")," type variable can hold ",(0,r.kt)("strong",{parentName:"p"},"any")," type of value depends upon the use of variable."),(0,r.kt)("p",null,"In 2nd line, we set the value of variable ",(0,r.kt)("inlineCode",{parentName:"p"},"vSketchLines"),"."),(0,r.kt)("p",null,"Value of ",(0,r.kt)("inlineCode",{parentName:"p"},"vSketchLinesis")," an array of lines. This array is send as return value when we use ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateCenterRectangle")," method."),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateCenterRectangle")," method is part of ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchManager")," and it is the latest method to create a Center rectangle."),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateCenterRectangle")," method takes following parameters as explained:"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"X1")," : X coordinate of the center point of rectangle"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Y1")," : Y coordinate of the center point of rectangle"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Z1")," : Z coordinate of the center point of rectangle"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"X2")," : X coordinate of the point 2"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Y2")," : Y coordinate of the point 2"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Z2")," : Z coordinate of the point 2"),(0,r.kt)("p",null,"Point 2 is one of the corners of rectangle we want to create."),(0,r.kt)("p",null,"Below image shows more clearly about these parameters."),(0,r.kt)("p",null,(0,r.kt)("img",{alt:"center-rectangle-parameter",src:n(51314).Z,width:"950",height:"720"})),(0,r.kt)("p",null,"In the above code sample I have used (0, 0, 0) point which is ",(0,r.kt)("em",{parentName:"p"},"origin")," of sketch."),(0,r.kt)("p",null,"For point 2, I used (1, 1, 0) which is 1 point distance in ",(0,r.kt)("em",{parentName:"p"},"X-direction")," and 1 point distance in ",(0,r.kt)("em",{parentName:"p"},"Y-direction"),"."),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateCenterRectangle")," method returns ",(0,r.kt)("strong",{parentName:"p"},"an array")," of ",(0,r.kt)("em",{parentName:"p"},"sketch segments")," that represent the edges created for this Center rectangle."),(0,r.kt)("p",null,"A ",(0,r.kt)("em",{parentName:"p"},"Sketch Segment")," can represent a sketch arc, line, ellipse, parabola or spline."),(0,r.kt)("p",null,"Sketch Segment has ",(0,r.kt)("inlineCode",{parentName:"p"},"ISketchSegment")," Interface, which provides functions that are generic to every type of sketch segment."),(0,r.kt)("p",null,"For example, every sketch segment has an ID and can be programmatically selected."),(0,r.kt)("p",null,"Therefore, the ",(0,r.kt)("inlineCode",{parentName:"p"},"ISketchSegment")," interface provides functions to obtain the ID and to select the item."),(0,r.kt)("admonition",{title:"NOTE",type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"It is ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"very important"))," to remember that, when you give distance or any other numeric value in ",(0,r.kt)("strong",{parentName:"p"},"Solidworks API"),", Solidworks takes that numeric value in ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Meter only")),"."),(0,r.kt)("p",{parentName:"admonition"},"Solidworks API does not care about your application's Unit systems."),(0,r.kt)("p",{parentName:"admonition"},"For example, I works in ANSI system means inches for distance. But when I used Solidworks API through VBA macros or C#, I need to use converted numeric values."),(0,r.kt)("p",{parentName:"admonition"},"Because Solidworks API output the distance in ",(0,r.kt)("strong",{parentName:"p"},"Meter")," which is not my requirement.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' De-select the lines after creation\nswDoc.ClearSelection2 True\n")),(0,r.kt)("p",null,"In the this line of code, we deselect the Center rectangle we have created."),(0,r.kt)("p",null,"For de-selecting, we use ",(0,r.kt)("inlineCode",{parentName:"p"},"ClearSelection2")," method from our Solidworks document name ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Zoom to fit screen in Solidworks Window\nswDoc.ViewZoomtofit\n")),(0,r.kt)("p",null,"In this last line we use ",(0,r.kt)("em",{parentName:"p"},"zoom to fit")," command."),(0,r.kt)("p",null,"For Zoom to fit, we use ",(0,r.kt)("inlineCode",{parentName:"p"},"ViewZoomtofit")," method from our Solidworks document variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc"),"."),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("em",{parentName:"p"},"create Center rectangle")," in Sketches with Solidworks VB Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"Solidworks VBA Macros"),", do come to this blog after sometime."),(0,r.kt)("p",null,"Till then, Happy learning!!!"))}d.isMDXComponent=!0},74753:(e,t,n)=>{n.d(t,{Z:()=>r});var a=n(67294);class r extends a.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return a.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}},51314:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/images/center-rectangle-parameter-72e04c4bcd5511e610be9677ba9e99fa.png"}}]);