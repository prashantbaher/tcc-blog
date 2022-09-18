"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[6671],{3905:function(e,t,n){n.d(t,{Zo:function(){return m},kt:function(){return d}});var a=n(67294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,a,o=function(e,t){if(null==e)return{};var n,a,o={},r=Object.keys(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var s=a.createContext({}),p=function(e){var t=a.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},m=function(e){var t=p(e.components);return a.createElement(s.Provider,{value:t},e.children)},k={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},c=a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,r=e.originalType,s=e.parentName,m=l(e,["components","mdxType","originalType","parentName"]),c=p(n),d=o,h=c["".concat(s,".").concat(d)]||c[d]||k[d]||r;return n?a.createElement(h,i(i({ref:t},m),{},{components:n})):a.createElement(h,i({ref:t},m))}));function d(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var r=n.length,i=new Array(r);i[0]=c;var l={};for(var s in t)hasOwnProperty.call(t,s)&&(l[s]=t[s]);l.originalType=e,l.mdxType="string"==typeof e?e:o,i[1]=l;for(var p=2;p<r;p++)i[p]=n[p];return a.createElement.apply(null,i)}return a.createElement.apply(null,n)}c.displayName="MDXCreateElement"},39726:function(e,t,n){n.r(t),n.d(t,{assets:function(){return k},contentTitle:function(){return p},default:function(){return h},frontMatter:function(){return s},metadata:function(){return m},toc:function(){return c}});var a=n(87462),o=n(63366),r=(n(67294),n(3905)),i=n(74753),l=["components"],s={categories:"Solidworks-macro",title:"Solidworks Macro - Create a Centerpoint Straight Slot",permalink:"/solidworks-macros/create-centerpoint-straight-slot/",tags:["Solidworks Macro"],id:"create-centerpoint-straight-slot"},p=void 0,m={unversionedId:"create-centerpoint-straight-slot",id:"create-centerpoint-straight-slot",title:"Solidworks Macro - Create a Centerpoint Straight Slot",description:"",source:"@site/docs/solidworks-macros/006.2-create-centerpoint-straight-slot.md",sourceDirName:".",slug:"/create-centerpoint-straight-slot",permalink:"/solidworks-macros/create-centerpoint-straight-slot",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks Macro - Create a Centerpoint Straight Slot",permalink:"/solidworks-macros/create-centerpoint-straight-slot/",tags:["Solidworks Macro"],id:"create-centerpoint-straight-slot"},sidebar:"tutorialSidebar",previous:{title:"Solidworks Macro - Create a Straight Slot",permalink:"/solidworks-macros/create-straight-slot"},next:{title:"Solidworks Macro - Create a 3-Point Arc Slot",permalink:"/solidworks-macros/create-3point-arc-slot"}},k={},c=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"Code Sample",id:"code-sample",level:2},{value:"Understanding the Code",id:"understanding-the-code",level:2}],d={toc:c};function h(e){var t=e.components,s=(0,o.Z)(e,l);return(0,r.kt)("wrapper",(0,a.Z)({},d,s,{components:t,mdxType:"MDXLayout"}),(0,r.kt)(i.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,"In this post, I tell you about ",(0,r.kt)("em",{parentName:"p"},"how to create a Centerpoint Straight Slot through Solidworks VBA Macros")," in a sketch."),(0,r.kt)("p",null,"The process is almost identical with previous \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/create-straight-slot"},"Sketch - Create Straight Slot"))," post."),(0,r.kt)("p",null,"In this post, I tell you about ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method from ",(0,r.kt)("strong",{parentName:"p"},"Solidworks")," ",(0,r.kt)("inlineCode",{parentName:"p"},"SketchManager")," object."),(0,r.kt)("p",null,"This method is ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"most updated"))," method, I found in ",(0,r.kt)("em",{parentName:"p"},"Solidworks API Help"),". "),(0,r.kt)("p",null,"So ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"use this method"))," if you want to create a new ",(0,r.kt)("strong",{parentName:"p"},"Centerpoint Straight Slot"),"."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,r.kt)("p",null,"Please see below video \ud83c\udfac on ",(0,r.kt)("strong",{parentName:"p"},"how to create a Centerpoint Straight Slot")," from Solidworks VBA Macros."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/ZNNh8mzpc-w",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,"Please note that there are ",(0,r.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"code-sample"},"Code Sample"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("inlineCode",{parentName:"p"},"code")," sample for creating ",(0,r.kt)("em",{parentName:"p"},"a Centerpoint Straight Slot"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n  \n  ' Creating string type variable for storing default part location\n  Dim defaultTemplate As String\n  ' Setting value of this string type variable to \"Default part template\"\n  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n  ' Setting Solidworks document to new part document\n  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\n  ' Selecting Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  \n  ' Setting Sketch manager for our sketch\n  Set swSketchManager = swDoc.SketchManager\n  \n  ' Inserting a sketch into selected plane\n  swSketchManager.InsertSketch True\n  \n  ' Creating Variable for Solidworks Slot\n  Dim mySketchSlot As SketchSlot\n      \n  ' Creating a Centerpoint Straight slot\n  Set mySketchSlot = swSketchManager.CreateSketchSlot(swSketchSlotCreationType_e.swSketchSlotCreationType_center_line, swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter, 1, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, False)\n  \n  ' De-select the Slot after creation\n  swDoc.ClearSelection2 True\n  \n  ' Zoom to fit screen in Solidworks Window\n  swDoc.ViewZoomtofit\n\nEnd Sub\n")),(0,r.kt)("hr",null),(0,r.kt)(i.Z,{mdxType:"AdComponent"}),(0,r.kt)("h2",{id:"understanding-the-code"},"Understanding the Code"),(0,r.kt)("p",null,"Now let us walk through ",(0,r.kt)("em",{parentName:"p"},"each line")," in the above code, and ",(0,r.kt)("strong",{parentName:"p"},"understand")," the meaning of every line."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,r.kt)("p",null,"This line forces us to define every variable we are going to use. "),(0,r.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"For more information please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post."))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,r.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp")," variable is ",(0,r.kt)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,r.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc")," variable is ",(0,r.kt)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),(0,r.kt)("p",null,"Next is our ",(0,r.kt)("inlineCode",{parentName:"p"},"Sub")," procedure named as ",(0,r.kt)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",(0,r.kt)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),(0,r.kt)("p",null,"In this line, we are setting the value of our Solidworks variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp")," which we defined earlier to Solidworks application."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),(0,r.kt)("p",null,"In 1st statement of above example, we are defining a variable of ",(0,r.kt)("inlineCode",{parentName:"p"},"string")," type and named it as ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),(0,r.kt)("p",null,"This variable ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),", holds the location the location of ",(0,r.kt)("strong",{parentName:"p"},"Default Part Template"),"."),(0,r.kt)("p",null,"In 2nd line of above example. we assign value to our newly define ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),(0,r.kt)("p",null,"We assign the value by using a ",(0,r.kt)("em",{parentName:"p"},"Method")," named ",(0,r.kt)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". "),(0,r.kt)("p",null,"This method is a part of our main Solidworks variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swApp"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),(0,r.kt)("p",null,"In this line, we set the value of our ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc")," variable to new document."),(0,r.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"For ",(0,r.kt)("strong",{parentName:"p"},"more detailed information")," about above lines please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post. "),(0,r.kt)("p",{parentName:"div"},"I have discussed them ",(0,r.kt)("strong",{parentName:"p"},"thoroghly")," in \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document"))," post, so do checkout this post if you don't understand above code."))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Boolean Variable\nDim BoolStatus As Boolean\n\n\' Selecting Front Plane\nBoolStatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n')),(0,r.kt)("p",null,"In 1st line, we create a variable named ",(0,r.kt)("inlineCode",{parentName:"p"},"BoolStatus")," as ",(0,r.kt)("inlineCode",{parentName:"p"},"Boolean")," object."),(0,r.kt)("p",null,"In next line, we select the ",(0,r.kt)("em",{parentName:"p"},"front plane")," by using ",(0,r.kt)("inlineCode",{parentName:"p"},"SelectByID2")," method from ",(0,r.kt)("inlineCode",{parentName:"p"},"Extension")," object."),(0,r.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"For more information about selection method please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/select-plane-from-tree"},"Solidworks Macros - Selection Methods"))," post."),(0,r.kt)("p",{parentName:"div"},"I have discussed about different ",(0,r.kt)("em",{parentName:"p"},"Selection methods")," in details in \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/select-plane-from-tree"},"Soldworks Macros - Selection Methods"))," post, so do visit this post for more ",(0,r.kt)("em",{parentName:"p"},"Selection methods"),"."))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n")),(0,r.kt)("p",null,"In above line, we create variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchManager")," for ",(0,r.kt)("strong",{parentName:"p"},"Solidworks Sketch Manager"),"."),(0,r.kt)("p",null,"As the name suggested, a ",(0,r.kt)("strong",{parentName:"p"},"Sketch Manager")," holds variours methods and properties to manage ",(0,r.kt)("em",{parentName:"p"},"Sketches"),"."),(0,r.kt)("p",null,"To see methods and properties related to SketchManager object, please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2017/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchManager_members.html"},"this page of Solidworks API Help"))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Sketch manager for our sketch\nSet swSketchManager = swDoc.SketchManager\n")),(0,r.kt)("p",null,"In above line, we set the ",(0,r.kt)("strong",{parentName:"p"},"Sketch manager")," variable to current document's sketch manager."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Inserting a sketch into selected plane\nswSketchManager.InsertSketch True\n")),(0,r.kt)("p",null,"In above line, we use ",(0,r.kt)("inlineCode",{parentName:"p"},"InsertSketch")," method of ",(0,r.kt)("em",{parentName:"p"},"SketchManager")," and give ",(0,r.kt)("inlineCode",{parentName:"p"},"True")," value."),(0,r.kt)("p",null,"This method allows us to insert a sketch in selected plane."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating Variable for Solidworks Slot\nDim mySketchSlot As SketchSlot\n      \n' Creating a Centerpoint Straight slot\nSet mySketchSlot = swSketchManager.CreateSketchSlot(swSketchSlotCreationType_e.swSketchSlotCreationType_center_line, swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter, 1, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, False)\n")),(0,r.kt)("p",null,"In above sample code, we 1st create a variable named ",(0,r.kt)("inlineCode",{parentName:"p"},"mySketchSlot")," of type ",(0,r.kt)("inlineCode",{parentName:"p"},"SketchSlot"),"."),(0,r.kt)("p",null,"In 2nd line, we ",(0,r.kt)("strong",{parentName:"p"},"set")," the value of ",(0,r.kt)("em",{parentName:"p"},"SketchSlot")," variable ",(0,r.kt)("inlineCode",{parentName:"p"},"mySketchSlot"),"."),(0,r.kt)("p",null,"We get this value from ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method which is inside the ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchManager")," variable."),(0,r.kt)("p",null,(0,r.kt)("inlineCode",{parentName:"p"},"swSketchManager")," variable is a type of ",(0,r.kt)("strong",{parentName:"p"},"SketchManager"),", hence we used ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method from ",(0,r.kt)("strong",{parentName:"p"},"SketchManager"),"."),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method takes following parameters as explained:"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"SlotCreationType")," : ",(0,r.kt)("em",{parentName:"p"},"Type of sketch slot")," as defined in ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e"),"."),(0,r.kt)("p",null,"  There are 4 Different types of Slots we can create."),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Straight Slot"))," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_line")," or ",(0,r.kt)("strong",{parentName:"p"},"0"))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Centerpoint straight Slot"))," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_center_line")," or ",(0,r.kt)("strong",{parentName:"p"},"1"))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Centerpoint arc Slot"))," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_arc")," or ",(0,r.kt)("strong",{parentName:"p"},"2"))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"3-point arc Slot"))," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_3pointarc")," or ",(0,r.kt)("strong",{parentName:"p"},"4")))),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"SlotLengthType")," : ",(0,r.kt)("em",{parentName:"p"},"Type of length of sketch slot")," as defined in ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e"),"."),(0,r.kt)("p",null,"  There are 2 different types of Sketch slot length we can create."),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Center to Center"))," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter")," or ",(0,r.kt)("strong",{parentName:"p"},"0"))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Full Length"))," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e.swSketchSlotLengthType_FullLength")," or ",(0,r.kt)("strong",{parentName:"p"},"1")))),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Width")," : Width of Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"X1")," : X coordinate of the point 1, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Y1")," : Y coordinate of the point 1, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Z1")," : Z coordinate of the point 1, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"X2")," : X coordinate of the point 2, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Y2")," : Y coordinate of the point 2, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Z2")," : Z coordinate of the point 2, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"X3")," : X coordinate of the point 3, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Y3")," : Y coordinate of the point 3, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Z3")," : Z coordinate of the point 3, of the Slot"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"CenterArcDirection")," : We need to set the direction eiter Clockwise or Anti-Clockwise/Counterclockwise as follows:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Clockwise (CW)"))," : -1")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Anti-Clockwise/Counterclockwise (CCW)"))," : 1"))),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"AddDimension")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"True")," to automatically add dimensions, ",(0,r.kt)("inlineCode",{parentName:"p"},"False")," to not."),(0,r.kt)("p",null,"For ",(0,r.kt)("strong",{parentName:"p"},"more details")," about ",(0,r.kt)("em",{parentName:"p"},"Slot Parameter")," you can visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.isketchmanager~createsketchslot.html"},"this page of Solidworks API Help"))),(0,r.kt)("p",null,"For creating a ",(0,r.kt)("em",{parentName:"p"},"Centerpoint Straight Slot"),", I used following parameter Values:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"SlotCreationType")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotCreationType_e.swSketchSlotCreationType_center_line")),(0,r.kt)("p",{parentName:"li"},"Since we want to create a ",(0,r.kt)("em",{parentName:"p"},"Centerpoint Straight Slot")," hence I select above value.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"SlotLengthType")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter")),(0,r.kt)("p",{parentName:"li"},"I want length of this Slot from ",(0,r.kt)("em",{parentName:"p"},"Center to Center")," hence I select above value.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"Width")," : ",(0,r.kt)("strong",{parentName:"p"},"1"))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"X1, Y1, Z1")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"0, 0, 0")),(0,r.kt)("p",{parentName:"li"},"For Point 1, I use (0, 0, 0) values, which is ",(0,r.kt)("em",{parentName:"p"},"origin")," of Sketch.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"X2, Y2, Z2")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"1, 0, 0")),(0,r.kt)("p",{parentName:"li"},"For Point 2, I use (1, 0, 0) values, which is which is 1 point distance in X-direction.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"X3, Y3, Z3")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"1, 1, 0")),(0,r.kt)("p",{parentName:"li"},"For Point 2, I use (1, 1, 0) values, which is which is 1 point distance in X-direction and 1 point distance in Y-direction.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"CenterArcDirection")," : ",(0,r.kt)("strong",{parentName:"p"},"1")),(0,r.kt)("p",{parentName:"li"},"I want to create Anti-Clockwise/Counterclockwise Slot.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("em",{parentName:"p"},"AddDimension")," : ",(0,r.kt)("inlineCode",{parentName:"p"},"False")))),(0,r.kt)("p",null,"Below Image described ",(0,r.kt)("strong",{parentName:"p"},"the Parameters for Centerpoint Straight Slot")," in more detail."),(0,r.kt)("p",null,(0,r.kt)("img",{alt:"centerpoint-straight-slot-parameters",src:n(37352).Z,width:"927",height:"484"})),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"CreateSketchSlot")," method returns ",(0,r.kt)("em",{parentName:"p"},"Sketch Slot")," interface i.e. ",(0,r.kt)("inlineCode",{parentName:"p"},"ISketchSlot")," interface. "),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"ISketchSlot")," interface has various ",(0,r.kt)("strong",{parentName:"p"},"methods and properties")," for ",(0,r.kt)("em",{parentName:"p"},"a Slot"),"."),(0,r.kt)("p",null,"For more detail about ",(0,r.kt)("strong",{parentName:"p"},"methods and properties")," of ",(0,r.kt)("inlineCode",{parentName:"p"},"ISketchSlot")," interface you can visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISketchSlot_members.html"},"this page of Solidworks API Help"))),(0,r.kt)("p",null,"::tip NOTE"),(0,r.kt)("p",null,"It is ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"very important"))," to remember that, when you give distance or any other numeric value in ",(0,r.kt)("strong",{parentName:"p"},"Solidworks API"),", Solidworks takes that numeric value in ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Meter only")),"."),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Solidworks API")," does not care about your application's Unit systems."),(0,r.kt)("p",null,'For example, I works in ANSI system means "inches" for distance. '),(0,r.kt)("p",null,"But when I used Solidworks API through ",(0,r.kt)("em",{parentName:"p"},"VBA macros")," or ",(0,r.kt)("em",{parentName:"p"},"C#"),", I have to use ",(0,r.kt)("strong",{parentName:"p"},"converted")," numeric values."),(0,r.kt)("p",null,"Because Solidworks API output the distance in ",(0,r.kt)("strong",{parentName:"p"},"Meter")," only; which is not my requirement.\n:::"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' De-select the Slot after creation\nswDoc.ClearSelection2 True\n")),(0,r.kt)("p",null,"In the this line of code, we de-select the created Centerpoint Straight Slot."),(0,r.kt)("p",null,"For de-selecting, we use ",(0,r.kt)("inlineCode",{parentName:"p"},"ClearSelection2")," method from our Solidworks document variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Zoom to fit screen in Solidworks Window\nswDoc.ViewZoomtofit\n")),(0,r.kt)("p",null,"In this last line we use ",(0,r.kt)("em",{parentName:"p"},"zoom to fit")," command."),(0,r.kt)("p",null,"For Zoom to fit, we use ",(0,r.kt)("inlineCode",{parentName:"p"},"ViewZoomtofit")," method from our Solidworks document variable ",(0,r.kt)("inlineCode",{parentName:"p"},"swDoc"),". "),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("em",{parentName:"p"},"create a Centerpoint Straight Slot")," in Sketches with Solidworks VB Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"Solidworks VBA Macros"),", do come to this blog after sometime."),(0,r.kt)("p",null,"Till then, Happy learning!!!"))}h.isMDXComponent=!0},74753:function(e,t,n){n.d(t,{Z:function(){return r}});var a=n(94578),o=n(67294),r=function(e){function t(){return e.apply(this,arguments)||this}(0,a.Z)(t,e);var n=t.prototype;return n.componentDidMount=function(){var e;(e=document.createElement("script")).src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild),(window.adsbygoogle=window.adsbygoogle||[]).push({})},n.render=function(){return o.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})},t}(o.Component)},37352:function(e,t,n){t.Z=n.p+"assets/images/centerpoint-straight-slot-parameters-57abb8c1d18bbfc60aaa79f1683c355b.png"}}]);