"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[9320],{3905:function(e,t,n){n.d(t,{Zo:function(){return m},kt:function(){return k}});var a=n(67294);function r(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function o(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?o(Object(n),!0).forEach((function(t){r(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):o(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,a,r=function(e,t){if(null==e)return{};var n,a,r={},o=Object.keys(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(a=0;a<o.length;a++)n=o[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var l=a.createContext({}),p=function(e){var t=a.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},m=function(e){var t=p(e.components);return a.createElement(l.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},d=a.forwardRef((function(e,t){var n=e.components,r=e.mdxType,o=e.originalType,l=e.parentName,m=s(e,["components","mdxType","originalType","parentName"]),d=p(n),k=r,u=d["".concat(l,".").concat(k)]||d[k]||c[k]||o;return n?a.createElement(u,i(i({ref:t},m),{},{components:n})):a.createElement(u,i({ref:t},m))}));function k(e,t){var n=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var o=n.length,i=new Array(o);i[0]=d;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:r,i[1]=s;for(var p=2;p<o;p++)i[p]=n[p];return a.createElement.apply(null,i)}return a.createElement.apply(null,n)}d.displayName="MDXCreateElement"},82956:function(e,t,n){n.r(t),n.d(t,{assets:function(){return c},contentTitle:function(){return p},default:function(){return u},frontMatter:function(){return l},metadata:function(){return m},toc:function(){return d}});var a=n(87462),r=n(63366),o=(n(67294),n(3905)),i=n(74753),s=["components"],l={categories:"Solidworks-macro",title:"Solidworks Sketch Macro - Create Line",permalink:"/solidworks-macros/sketch-create-line/",tags:["Solidworks Macro"],id:"sketch-create-line"},p=void 0,m={unversionedId:"sketch-create-line",id:"sketch-create-line",title:"Solidworks Sketch Macro - Create Line",description:"",source:"@site/docs/solidworks-macros/002.1-sketch-create-line.md",sourceDirName:".",slug:"/sketch-create-line",permalink:"/solidworks-macros/sketch-create-line",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks Sketch Macro - Create Line",permalink:"/solidworks-macros/sketch-create-line/",tags:["Solidworks Macro"],id:"sketch-create-line"},sidebar:"tutorialSidebar",previous:{title:"Solidworks Macro - Fix Unit Issue",permalink:"/solidworks-macros/unit-correction"},next:{title:"Solidworks Sketch Macro - Create CenterLine",permalink:"/solidworks-macros/sketch-create-centerline"}},c={},d=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"Code Sample",id:"code-sample",level:2},{value:"Understanding Code",id:"understanding-code",level:3}],k={toc:d};function u(e){var t=e.components,n=(0,r.Z)(e,s);return(0,o.kt)("wrapper",(0,a.Z)({},k,n,{components:t,mdxType:"MDXLayout"}),(0,o.kt)(i.Z,{mdxType:"AdComponent"}),(0,o.kt)("p",null,"In this post, I tell you about ",(0,o.kt)("em",{parentName:"p"},"how to create 2D Line through Solidworks VBA Macros")," in a sketch."),(0,o.kt)("p",null,"For this, I take the example from previous \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"/solidworks-macros/select-plane-from-tree"},"Solidworks Macros - Open Assembly and Drawing document"))," post."),(0,o.kt)("p",null,"In this post, I tell you about ",(0,o.kt)("inlineCode",{parentName:"p"},"CreateLine")," method from ",(0,o.kt)("strong",{parentName:"p"},"Solidworks")," ",(0,o.kt)("inlineCode",{parentName:"p"},"SketchManager")," object."),(0,o.kt)("p",null,"This method is ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"most updated"))," method, I found in ",(0,o.kt)("em",{parentName:"p"},"Solidworks API Help"),". "),(0,o.kt)("p",null,"So ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"use this method"))," if you want to create a new line."),(0,o.kt)("hr",null),(0,o.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,o.kt)("p",null,"Please see below video \ud83c\udfac on ",(0,o.kt)("strong",{parentName:"p"},"how to create 2D Line")," from Solidworks VBA Macros."),(0,o.kt)("iframe",{src:"https://www.youtube.com/embed/qDvQF8xBCSk",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,o.kt)("p",null,"Please note that there are ",(0,o.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,o.kt)("hr",null),(0,o.kt)("h2",{id:"code-sample"},"Code Sample"),(0,o.kt)("p",null,"Below is the ",(0,o.kt)("inlineCode",{parentName:"p"},"code")," sample for creating lines."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Main function of our VBA program\nSub main()\n\n    ' Setting Solidworks variable to Solidworks application\n    Set swApp = Application.SldWorks\n    \n    ' Creating string type variable for storing default part location\n    Dim defaultTemplate As String\n    ' Setting value of this string type variable to \"Default part template\"\n    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n    ' Setting Solidworks document to new part document\n    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\n    ' Selecting Front Plane\n    BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n    \n    ' Setting Sketch manager for our sketch\n    Set swSketchManager = swDoc.SketchManager\n    \n    ' Creating Variable for Solidworks Sketch segment\n    Dim mySketchSegment As SketchSegment\n    \n    ' Inserting a sketch into selected plane\n    swSketchManager.InsertSketch True\n    \n    ' Creating an horizontal line\n    Set mySketchSegment = swSketchManager.CreateLine(0, 0, 0, 2, 0, 0)\n    \n    ' De-select the line after creation\n    swDoc.ClearSelection2 True\n\nEnd Sub\n")),(0,o.kt)(i.Z,{mdxType:"AdComponent"}),(0,o.kt)("h3",{id:"understanding-code"},"Understanding Code"),(0,o.kt)("p",null,"Now let us walk through ",(0,o.kt)("em",{parentName:"p"},"each line")," in the above code, and ",(0,o.kt)("strong",{parentName:"p"},"understand")," the meaning of every line."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,o.kt)("p",null,"This line forces us to define every variable we are going to use. "),(0,o.kt)("p",null,"For more information please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,o.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp")," variable is ",(0,o.kt)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,o.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc")," variable is ",(0,o.kt)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),(0,o.kt)("p",null,"Next is our ",(0,o.kt)("inlineCode",{parentName:"p"},"Sub")," procedure named ",(0,o.kt)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",(0,o.kt)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),(0,o.kt)("p",null,"In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),(0,o.kt)("p",null,"In 1st statement of above example, we are defining a variable of ",(0,o.kt)("inlineCode",{parentName:"p"},"string")," type and named it as ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),(0,o.kt)("p",null,"This variable ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),", hold the location the location of ",(0,o.kt)("strong",{parentName:"p"},"Default Part Template"),"."),(0,o.kt)("p",null,"In 2nd line of above example. we assign value to our newly define ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),(0,o.kt)("p",null,"We assign the value by using a ",(0,o.kt)("em",{parentName:"p"},"Method")," named ",(0,o.kt)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". This method is a part of our main Solidworks variable ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp"),"."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),(0,o.kt)("p",null,"In this line, we set the value of our ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc")," variable to new document."),(0,o.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,o.kt)("div",{parentName:"div",className:"admonition-heading"},(0,o.kt)("h5",{parentName:"div"},(0,o.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,o.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,o.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,o.kt)("div",{parentName:"div",className:"admonition-content"},(0,o.kt)("p",{parentName:"div"},"For ",(0,o.kt)("strong",{parentName:"p"},"detailed information")," about these lines please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post."))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Boolean Variable\nDim BoolStatus As Boolean\n\n\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),(0,o.kt)("p",null,"In 1st line, we create a variable named ",(0,o.kt)("inlineCode",{parentName:"p"},"BoolStatus")," as ",(0,o.kt)("inlineCode",{parentName:"p"},"Boolean")," object."),(0,o.kt)("p",null,"In next line, we select the ",(0,o.kt)("em",{parentName:"p"},"front plane")," by using ",(0,o.kt)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",(0,o.kt)("inlineCode",{parentName:"p"},"Extension")," object."),(0,o.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,o.kt)("div",{parentName:"div",className:"admonition-heading"},(0,o.kt)("h5",{parentName:"div"},(0,o.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,o.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,o.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,o.kt)("div",{parentName:"div",className:"admonition-content"},(0,o.kt)("p",{parentName:"div"},"For more information about selection method please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"/solidworks-macros/select-plane-from-tree"},"Solidworks Macros - Selection Methods"))," post."))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n")),(0,o.kt)("p",null,"In above line, we create variable ",(0,o.kt)("inlineCode",{parentName:"p"},"swSketchManager")," for ",(0,o.kt)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),(0,o.kt)("p",null,"As the name suggested, a Sketch Manager holds variours methods and properties to manage Sketches."),(0,o.kt)("p",null,"To see methods and properties related to SketchManager object, please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"},"this page"))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Sketch manager for our sketch\nSet swSketchManager = swDoc.SketchManager\n")),(0,o.kt)("p",null,"In above line, we set the sketch manager variable to current document's sketch manager."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Inserting a sketch into selected plane\nswSketchManager.InsertSketch True\n")),(0,o.kt)("p",null,"In above line, we use ",(0,o.kt)("inlineCode",{parentName:"p"},"InsertSketch")," method of ",(0,o.kt)("em",{parentName:"p"},"SketchManager")," and give ",(0,o.kt)("inlineCode",{parentName:"p"},"True")," value."),(0,o.kt)("p",null,"This method allows us to insert a sketch in selected plane."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating Variable for Solidworks Sketch segment\nDim mySketchSegment As SketchSegment\n\n' Creating an horizontal line\nSet mySketchSegment = swSketchManager.CreateLine(0, 0, 0, 2, 0, 0)\n")),(0,o.kt)("p",null,"In above sample code, we 1st create a variable named ",(0,o.kt)("inlineCode",{parentName:"p"},"mySketchSegment")," of type ",(0,o.kt)("inlineCode",{parentName:"p"},"SketchSegment"),"."),(0,o.kt)("p",null,"A ",(0,o.kt)("inlineCode",{parentName:"p"},"SketchSegment")," represent ",(0,o.kt)("em",{parentName:"p"},"a line, ellipse, parabola or spline.")),(0,o.kt)("p",null,"A ",(0,o.kt)("inlineCode",{parentName:"p"},"SketchSegment")," provides functions that are ",(0,o.kt)("strong",{parentName:"p"},"generic")," to every type of sketch segment."),(0,o.kt)("p",null,"For example, every sketch segment has ",(0,o.kt)("strong",{parentName:"p"},"an ID")," and can be selected programmatically."),(0,o.kt)("p",null,"Therefore, the ",(0,o.kt)("inlineCode",{parentName:"p"},"SketchSegment")," interface provides functions to obtain the ID and to select the item."),(0,o.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,o.kt)("div",{parentName:"div",className:"admonition-heading"},(0,o.kt)("h5",{parentName:"div"},(0,o.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,o.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,o.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,o.kt)("div",{parentName:"div",className:"admonition-content"},(0,o.kt)("p",{parentName:"div"},"For detailed information about the ",(0,o.kt)("inlineCode",{parentName:"p"},"SketchSegment")," please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2017/english/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.ISketchSegment.html"},"this page of Solidworks API Help"))))),(0,o.kt)("p",null,"In 2nd line, we set the value of sketch segment variable ",(0,o.kt)("inlineCode",{parentName:"p"},"mySketchSegment"),"."),(0,o.kt)("p",null,"We get this value from ",(0,o.kt)("inlineCode",{parentName:"p"},"CreateLine")," method which is inside the ",(0,o.kt)("inlineCode",{parentName:"p"},"swSketchManager")," variable."),(0,o.kt)("p",null,(0,o.kt)("inlineCode",{parentName:"p"},"swSketchManager")," variable is a type of SketchManager, hence we used ",(0,o.kt)("inlineCode",{parentName:"p"},"CreateLine")," method from SketchManager."),(0,o.kt)("p",null,"This ",(0,o.kt)("inlineCode",{parentName:"p"},"CreateLine")," method takes following parameters as explained:"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},(0,o.kt)("em",{parentName:"p"},"X1")," : X coordinate of the line start point")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},(0,o.kt)("em",{parentName:"p"},"Y1")," : Y coordinate of the line start point")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},(0,o.kt)("em",{parentName:"p"},"Z1")," : Z coordinate of the line start point")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},(0,o.kt)("em",{parentName:"p"},"X2")," : X coordinate of the line end point")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},(0,o.kt)("em",{parentName:"p"},"Y2")," : Y coordinate of the line end point")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("p",{parentName:"li"},(0,o.kt)("em",{parentName:"p"},"Z2")," : Z coordinate of the line end point"))),(0,o.kt)("p",null,"In the above code sample I have used (0, 0, 0) for start point."),(0,o.kt)("p",null,"This is origin of sketch hence I start line from origin."),(0,o.kt)("p",null,"For End point I used (2, 0, 0) which is 2 point distance in X-direction or horizontal direction."),(0,o.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,o.kt)("div",{parentName:"div",className:"admonition-heading"},(0,o.kt)("h5",{parentName:"div"},(0,o.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,o.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,o.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"NOTE")),(0,o.kt)("div",{parentName:"div",className:"admonition-content"},(0,o.kt)("p",{parentName:"div"},"It is ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"very important"))," to remember that, when you give distance or any other numeric value in ",(0,o.kt)("strong",{parentName:"p"},"Solidworks API"),", Solidworks takes that numeric value in ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"Meter only")),"."),(0,o.kt)("p",{parentName:"div"},"Solidworks API does not care about your application's Unit systems."),(0,o.kt)("p",{parentName:"div"},"For example, I works in ANSI system means inches for distance. But when I used Solidworks API through VBA macros or C#, I need to use converted numeric values."),(0,o.kt)("p",{parentName:"div"},"Because Solidworks API output the distance in ",(0,o.kt)("strong",{parentName:"p"},"Meter")," which is not my requirement."))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' De-select the line after creation\nswDoc.ClearSelection2 True\n")),(0,o.kt)("p",null,"In the last line of code, we de-select the created line."),(0,o.kt)("p",null,"For de-selecting, we use ",(0,o.kt)("inlineCode",{parentName:"p"},"ClearSelection2")," method from our Solidworks document name ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc"),"."),(0,o.kt)("p",null,"Hope this post helps you to ",(0,o.kt)("em",{parentName:"p"},"create lines")," in Sketches with Solidworks VB Macros."),(0,o.kt)("p",null,"For more such tutorials on ",(0,o.kt)("strong",{parentName:"p"},"Solidworks VBA Macros"),", do come to this blog after sometime."),(0,o.kt)("p",null,"Till then, Happy learning!!!"))}u.isMDXComponent=!0},74753:function(e,t,n){n.d(t,{Z:function(){return o}});var a=n(94578),r=n(67294),o=function(e){function t(){return e.apply(this,arguments)||this}(0,a.Z)(t,e);var n=t.prototype;return n.componentDidMount=function(){var e;(e=document.createElement("script")).src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild),(window.adsbygoogle=window.adsbygoogle||[]).push({})},n.render=function(){return r.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})},t}(r.Component)}}]);