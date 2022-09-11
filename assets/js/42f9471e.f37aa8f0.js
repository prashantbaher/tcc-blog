"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[4995],{3905:function(e,t,n){n.d(t,{Zo:function(){return p},kt:function(){return k}});var o=n(67294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},r=Object.keys(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var l=o.createContext({}),c=function(e){var t=o.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},p=function(e){var t=c(e.components);return o.createElement(l.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},g=o.forwardRef((function(e,t){var n=e.components,a=e.mdxType,r=e.originalType,l=e.parentName,p=s(e,["components","mdxType","originalType","parentName"]),g=c(n),k=a,d=g["".concat(l,".").concat(k)]||g[k]||m[k]||r;return n?o.createElement(d,i(i({ref:t},p),{},{components:n})):o.createElement(d,i({ref:t},p))}));function k(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var r=n.length,i=new Array(r);i[0]=g;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:a,i[1]=s;for(var c=2;c<r;c++)i[c]=n[c];return o.createElement.apply(null,i)}return o.createElement.apply(null,n)}g.displayName="MDXCreateElement"},98931:function(e,t,n){n.r(t),n.d(t,{assets:function(){return p},contentTitle:function(){return l},default:function(){return k},frontMatter:function(){return s},metadata:function(){return c},toc:function(){return m}});var o=n(87462),a=n(63366),r=(n(67294),n(3905)),i=["components"],s={categories:"Solidworks-macro",title:"Solidworks Macro - Toggle (Hide/Show) Sketch Relations",permalink:"/solidworks-macros/toggle-display-sketch-relation/",tags:["Solidworks Macro"],id:"toggle-display-sketch-relation"},l=void 0,c={unversionedId:"toggle-display-sketch-relation",id:"toggle-display-sketch-relation",title:"Solidworks Macro - Toggle (Hide/Show) Sketch Relations",description:"In this post, I tell you about how to Toggle (Hide/Show) Sketch Relations using Solidworks VBA Macros in a Sketch.",source:"@site/docs/solidworks-macros/013.1-toggle-display-sketch-relation.md",sourceDirName:".",slug:"/toggle-display-sketch-relation",permalink:"/solidworks-macros/toggle-display-sketch-relation",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks Macro - Toggle (Hide/Show) Sketch Relations",permalink:"/solidworks-macros/toggle-display-sketch-relation/",tags:["Solidworks Macro"],id:"toggle-display-sketch-relation"},sidebar:"tutorialSidebar",previous:{title:"Solidworks Macro - Rotate/Copy Sketch Entities",permalink:"/solidworks-macros/rotate-copy-sketch-entities"},next:{title:"Solidworks Macro - Add Sketch Relations (Constraints)",permalink:"/solidworks-macros/add-sketch-relations"}},p={},m=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"Code Sample",id:"code-sample",level:2},{value:"Understanding the Code",id:"understanding-the-code",level:2}],g={toc:m};function k(e){var t=e.components,s=(0,a.Z)(e,i);return(0,r.kt)("wrapper",(0,o.Z)({},g,s,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("p",null,"In this post, I tell you about ",(0,r.kt)("strong",{parentName:"p"},"how to Toggle (Hide/Show) Sketch Relations using Solidworks VBA Macros")," in a Sketch."),(0,r.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"This post is extension to \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},(0,r.kt)("a",{parentName:"em",href:"/solidworks-macros/rotate-copy-sketch-entities"},"Sketch Transformation - Rotate/Copy Sketch Entities")))," post."),(0,r.kt)("p",{parentName:"div"},"Hence I will explained only ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Toggle (Hide/Show) Sketch Relations"))," related code."))),(0,r.kt)("p",null,"In this post, I explain about ",(0,r.kt)("inlineCode",{parentName:"p"},"SetUserPreferenceToggle")," method from ",(0,r.kt)("strong",{parentName:"p"},"Solidworks"),"'s ",(0,r.kt)("inlineCode",{parentName:"p"},"ModelDoc2")," object."),(0,r.kt)("p",null,"This method is ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"NOT updated"))," method, but it is easiest way to ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Toggle (Hide/Show) Sketch Relations")),"."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,r.kt)("p",null,"Please see below video \ud83c\udfac on ",(0,r.kt)("strong",{parentName:"p"},"how to Toggle (Hide/Show) Sketch Relations")," from Solidworks VBA Macros."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/9Ck6iPY_4gs",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,"Please note that there are ",(0,r.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"code-sample"},"Code Sample"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("inlineCode",{parentName:"p"},"code")," sample to ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Toggle (Hide/Show) Sketch Relations")),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Create variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n\n' Create variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Boolean Variable\nDim BoolStatus As Boolean\n\n' Create variable for Solidworks Sketch Manager\nDim swSketchManager As SldWorks.SketchManager\n\n' Create Variable for Solidworks Sketch Segment\nDim swSketchSegment As SldWorks.SketchSegment\n\n' Main function of our VBA program\nSub main()\n\n  ' Set Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n  \n  ' Create string type variable for storing default part location\n  Dim defaultTemplate As String\n\n  ' Set value of this string type variable to \"Default part template\"\n  defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n  ' Set Solidworks document to new part document\n  Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n  \n  '-----------------------UNIT CONVERSION----------------------------------------\n\n  ' Local variables used as Conversion Factors\n  Dim LengthConversionFactor As Double\n  Dim AngleConversionFactor As Double\n  \n  ' Use a Select Case, to get the length of active Unit and set the different factors\n  Select Case swDoc.GetUnits(0)       ' GetUnits function gives us, active unit\n    \n    Case swMETER    ' If length is in Meter\n      LengthConversionFactor = 1\n      AngleConversionFactor = 1\n    \n    Case swMM       ' If length is in MM\n      LengthConversionFactor = 1 / 1000\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swCM       ' If length is in CM\n      LengthConversionFactor = 1 / 100\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swINCHES   ' If length is in INCHES\n      LengthConversionFactor = 1 * 0.0254\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swFEET     ' If length is in FEET\n      LengthConversionFactor = 1 * (0.0254 * 12)\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swFEETINCHES     ' If length is in FEET & INCHES\n      LengthConversionFactor = 1 * 0.0254  ' For length we use sama as Inch\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swANGSTROM        ' If length is in ANGSTROM\n      LengthConversionFactor = 1 / 10000000000#\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swNANOMETER       ' If length is in NANOMETER\n      LengthConversionFactor = 1 / 1000000000\n      AngleConversionFactor = 1 * 0.01745329\n    \n    Case swMICRON       ' If length is in MICRON\n      LengthConversionFactor = 1 / 1000000\n      AngleConversionFactor = 1 * 0.01745329\n  End Select\n\n  '----------------------------------------------------------------\n\n  ' Select Front Plane\n  BoolStatus = swDoc.Extension.SelectByID2(\"Front Plane\", \"PLANE\", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n\n  ' Set Sketch manager for our sketch\n  Set swSketchManager = swDoc.SketchManager\n\n  ' Insert a sketch into selected plane\n  swSketchManager.InsertSketch True\n  \n  ' Create a local variable for CenterPoint ractangle\n  Dim vSketch As Variant\n  \n  ' Create CenterPoint ractangle\n  vSketch = swSketchManager.CreateCenterRectangle(0, 0, 0, 1 * LengthConversionFactor, 1 * LengthConversionFactor, 0)\n    \n  ' De-select the lines after creation\n  swDoc.ClearSelection2 True\n  \n  ' Select all lines of CenterPoint Ractangle\n  BoolStatus = swDoc.Extension.SelectByID2(\"Line1\", \"SKETCHSEGMENT\", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  BoolStatus = swDoc.Extension.SelectByID2(\"Line2\", \"SKETCHSEGMENT\", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  BoolStatus = swDoc.Extension.SelectByID2(\"Line3\", \"SKETCHSEGMENT\", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  BoolStatus = swDoc.Extension.SelectByID2(\"Line4\", \"SKETCHSEGMENT\", 0, 0, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)\n  \n  ' Rotate CenterPoint Ractangle by 45 degree only\n  swDoc.Extension.RotateOrCopy True, 2, True, 0, 0, 0, 0, 0, 1, 45 * AngleConversionFactor\n  \n  ' Toggle (Hide/Show) Sketch Relations\n  BoolStatus = swDoc.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewSketchRelations, True)\n  \n  ' De-select all after creation\n  swDoc.ClearSelection2 True\n  \n  ' Show Front View after Circular Sketch Pattern\n  swDoc.ShowNamedView2 \"\", swStandardViews_e.swFrontView\n  \n  ' Zoom to fit screen in Solidworks Window\n  swDoc.ViewZoomtofit2\n  \nEnd Sub\n")),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"understanding-the-code"},"Understanding the Code"),(0,r.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"I have already discuss above code in previous \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},(0,r.kt)("a",{parentName:"em",href:"/solidworks-macros/rotate-copy-sketch-entities"},"Sketch Transformation - Rotate/Copy Sketch Entities")))," post except below line of code."))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Toggle (Hide/Show) Sketch Relations\nBoolStatus = swDoc.SetUserPreferenceToggle(swUserPreferenceToggle_e.swViewSketchRelations, True)\n")),(0,r.kt)("p",null,'For "',(0,r.kt)("strong",{parentName:"p"},"Toggle (Hide/Show)"),'"  Sketch Relations, we need ',(0,r.kt)("inlineCode",{parentName:"p"},"SetUserPreferenceToggle")," method from ",(0,r.kt)("strong",{parentName:"p"},"Solidworks"),"'s ",(0,r.kt)("inlineCode",{parentName:"p"},"ModelDoc2")," object."),(0,r.kt)("p",null,"This ",(0,r.kt)("inlineCode",{parentName:"p"},"SetUserPreferenceToggle")," method takes following parameters as explained:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"UserPreferenceValue")," : ",(0,r.kt)("em",{parentName:"li"},"Use Preference Values to toggle as defined in ",(0,r.kt)("inlineCode",{parentName:"em"},"swUserPreferenceToggle_e"),"."))),(0,r.kt)("div",{className:"admonition admonition-info alert alert--info"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"14",height:"16",viewBox:"0 0 14 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M7 2.3c3.14 0 5.7 2.56 5.7 5.7s-2.56 5.7-5.7 5.7A5.71 5.71 0 0 1 1.3 8c0-3.14 2.56-5.7 5.7-5.7zM7 1C3.14 1 0 4.14 0 8s3.14 7 7 7 7-3.14 7-7-3.14-7-7-7zm1 3H6v5h2V4zm0 6H6v2h2v-2z"}))),"NOTE ")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},(0,r.kt)("inlineCode",{parentName:"p"},"swUserPreferenceToggle_e")," has many values!!!  Hence it is not possible to list all of them here. If you want to check full list, please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2020/English/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swUserPreferenceToggle_e.html"},"this page of Solidworks API Help")),"."))),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"OnFlag")," : ",(0,r.kt)("em",{parentName:"li"},"True to toggle the value on, false to toggle the value off."))),(0,r.kt)("p",null,"In our code, we used following values:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"UserPreferenceValue")," : ",(0,r.kt)("em",{parentName:"p"},(0,r.kt)("inlineCode",{parentName:"em"},"swUserPreferenceToggle_e.swViewSketchRelations")))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"OnFlag")," : ",(0,r.kt)("em",{parentName:"p"},(0,r.kt)("inlineCode",{parentName:"em"},"True"))))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"Return Value"),":"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"True"),": ",(0,r.kt)("em",{parentName:"p"},"If Toggle (Hide/Show) of Sketch Relations is ",(0,r.kt)("strong",{parentName:"em"},"Success"),"."))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"False"),": ",(0,r.kt)("em",{parentName:"p"},"If Toggle (Hide/Show) of Sketch Relations is ",(0,r.kt)("strong",{parentName:"em"},"Fail"),".")))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"Before Toggle (Hide/Show) of Sketch Relations")),(0,r.kt)("p",null,(0,r.kt)("img",{alt:"before-toggle-sketch-relation",src:n(66459).Z,width:"1121",height:"553"})),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"After Toggle (Hide/Show) of Sketch Relations")),(0,r.kt)("p",null,(0,r.kt)("img",{alt:"after-toggle-sketch-relation",src:n(80526).Z,width:"1122",height:"549"})),(0,r.kt)("hr",null),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"This is it !!!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")),(0,r.kt)("p",null,"If you found anything to ",(0,r.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,r.kt)("em",{parentName:"p"},"e-mail"),"."),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Toggle (Hide/Show) of Sketch Relations"))," with Solidworks VBA Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"Solidworks VBA Macro"),", do come to this blog after sometime."),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}k.isMDXComponent=!0},80526:function(e,t,n){t.Z=n.p+"assets/images/after-toggle-sketch-relation-48ecae9fb1999c23432f426296ceb1f4.png"},66459:function(e,t,n){t.Z=n.p+"assets/images/before-toggle-sketch-relation-5d4937bee3806c8ebe4cc343a8157914.png"}}]);