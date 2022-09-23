"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[7237],{3905:function(e,t,n){n.d(t,{Zo:function(){return d},kt:function(){return c}});var o=n(67294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},r=Object.keys(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var l=o.createContext({}),p=function(e){var t=o.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},d=function(e){var t=p(e.components);return o.createElement(l.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},u=o.forwardRef((function(e,t){var n=e.components,a=e.mdxType,r=e.originalType,l=e.parentName,d=s(e,["components","mdxType","originalType","parentName"]),u=p(n),c=a,w=u["".concat(l,".").concat(c)]||u[c]||m[c]||r;return n?o.createElement(w,i(i({ref:t},d),{},{components:n})):o.createElement(w,i({ref:t},d))}));function c(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var r=n.length,i=new Array(r);i[0]=u;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:a,i[1]=s;for(var p=2;p<r;p++)i[p]=n[p];return o.createElement.apply(null,i)}return o.createElement.apply(null,n)}u.displayName="MDXCreateElement"},11522:function(e,t,n){n.r(t),n.d(t,{assets:function(){return m},contentTitle:function(){return p},default:function(){return w},frontMatter:function(){return l},metadata:function(){return d},toc:function(){return u}});var o=n(87462),a=n(63366),r=(n(67294),n(3905)),i=n(74753),s=["components"],l={categories:"solidworks-macro",title:"Solidworks Macro - Open Assembly and Drawing document",permalink:"/solidworks-macros/open-assembly-and-drawing/",tags:["Solidworks Macro"],id:"open-assembly-and-drawing"},p=void 0,d={unversionedId:"open-assembly-and-drawing",id:"open-assembly-and-drawing",title:"Solidworks Macro - Open Assembly and Drawing document",description:"",source:"@site/docs/solidworks-macros/001.3-open-assembly-and-drawing.md",sourceDirName:".",slug:"/open-assembly-and-drawing",permalink:"/solidworks-macros/open-assembly-and-drawing",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"solidworks-macro",title:"Solidworks Macro - Open Assembly and Drawing document",permalink:"/solidworks-macros/open-assembly-and-drawing/",tags:["Solidworks Macro"],id:"open-assembly-and-drawing"},sidebar:"tutorialSidebar",previous:{title:"Solidworks Macro - Open new Part document",permalink:"/solidworks-macros/open-new-document"},next:{title:"Solidworks Macro - Selection Methods",permalink:"/solidworks-macros/select-plane-from-tree"}},m={},u=[{value:"Solidworks Assembly Document",id:"solidworks-assembly-document",level:2},{value:"Solidworks Drawing Document without Defining Paper size",id:"solidworks-drawing-document-without-defining-paper-size",level:2},{value:"Solidworks Drawing Document with Default Paper size",id:"solidworks-drawing-document-with-default-paper-size",level:3},{value:"Solidworks Drawing Document with Custom Paper size",id:"solidworks-drawing-document-with-custom-paper-size",level:3}],c={toc:u};function w(e){var t=e.components,n=(0,a.Z)(e,s);return(0,r.kt)("wrapper",(0,o.Z)({},c,n,{components:t,mdxType:"MDXLayout"}),(0,r.kt)(i.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,"In this post, we see how to open following documents with ",(0,r.kt)("em",{parentName:"p"},"Solidworks VBA macro"),":"),(0,r.kt)("ol",null,(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Solidworks Assembly document")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Solidworks Drawing document"),(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Without")," Pre-defined Sheet size"),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"With")," Pre-defined Sheet size"),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("em",{parentName:"li"},"With Custom Sheet size"))))),(0,r.kt)("h2",{id:"solidworks-assembly-document"},"Solidworks Assembly Document"),(0,r.kt)("p",null,"The code for opening ",(0,r.kt)("em",{parentName:"p"},"default Assembly document")," is identical to the ",(0,r.kt)("em",{parentName:"p"},"default Part template")," with only one change."),(0,r.kt)("p",null,"First, let us see the code to open default Assembly document."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Main function of our VBA program\nSub main()\n\n    ' Setting Solidworks variable to Solidworks application\n    Set swApp = Application.SldWorks\n    \n    ' Creating string type variable for storing default Assembly location\n    Dim defaultTemplate As String\n    ' Setting value of this string type variable to \"Default Assembly template\"\n    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)\n\n    ' Setting Solidworks document to new Assembly document\n    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\nEnd Sub\n")),(0,r.kt)("p",null,"As you can see in the above code and the code is given in the \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"previous post"))," is almost identically."),(0,r.kt)("p",null,"In case you have not read my previous post (\ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"Solidworks Macros - Open new Part document")),"), I recommend you to read that post first. "),(0,r.kt)("p",null,"I have already explained each and every line in this code there. So I will not explain them in this post."),(0,r.kt)("p",null,"To open default assembly template, you just need to change ",(0,r.kt)("inlineCode",{parentName:"p"},"defaultTemplate")," variable and use ",(0,r.kt)("inlineCode",{parentName:"p"},"swDefaultTemplateAssembly")," in place of ",(0,r.kt)("inlineCode",{parentName:"p"},"swDefaultTemplatePart"),"."),(0,r.kt)("p",null,"With this you can open a new assembly document."),(0,r.kt)("hr",null),(0,r.kt)("h2",{id:"solidworks-drawing-document-without-defining-paper-size"},"Solidworks Drawing Document without Defining Paper size"),(0,r.kt)("p",null,"To open new ",(0,r.kt)("em",{parentName:"p"},"Default drawing document")," we use same code as used above but with slight modification."),(0,r.kt)("p",null,"If we make similar change as we have done for ",(0,r.kt)("em",{parentName:"p"},"Assembly document")," then we open ",(0,r.kt)("em",{parentName:"p"},"Default drawing document")," ",(0,r.kt)("strong",{parentName:"p"},"without")," specifying sheet size."),(0,r.kt)("p",null,"The code sample shows how to do it."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Main function of our VBA program\nSub main()\n\n    ' Setting Solidworks variable to Solidworks application\n    Set swApp = Application.SldWorks\n    \n    ' Creating string type variable for storing default drawing location\n    Dim defaultTemplate As String\n    ' Setting value of this string type variable to \"Default drawing template\" without define paper size\n    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)\n\n    ' Setting Solidworks document to new drawing document\n    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\nEnd Sub\n")),(0,r.kt)("hr",null),(0,r.kt)("h3",{id:"solidworks-drawing-document-with-default-paper-size"},"Solidworks Drawing Document with Default Paper size"),(0,r.kt)("p",null,"To open a ",(0,r.kt)("em",{parentName:"p"},"new Drawing")," with ",(0,r.kt)("strong",{parentName:"p"},"pre-define")," sheet size we use following code sample:"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Main function of our VBA program\nSub main()\n\n    ' Setting Solidworks variable to Solidworks application\n    Set swApp = Application.SldWorks\n    \n    ' Creating string type variable for storing default drawing location\n    Dim defaultTemplate As String\n    ' Setting value of this string type variable to \"Default drawing template\" with pre-define paper size\n    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)\n\n    ' Setting Solidworks document to new drawing document\n    Set swDoc = swApp.NewDocument(defaultTemplate, swDwgPaperSizes_e.swDwgPaperA4sizeVertical, 0, 0)\n\nEnd Sub\n")),(0,r.kt)("p",null,"This code is ",(0,r.kt)("em",{parentName:"p"},"similar")," to what we have used in the ",(0,r.kt)("em",{parentName:"p"},"previous section")," but has a ",(0,r.kt)("em",{parentName:"p"},"one modification"),"."),(0,r.kt)("p",null,"That is while setting the document (at ",(0,r.kt)("inlineCode",{parentName:"p"},"Set Doc"),") we ",(0,r.kt)("em",{parentName:"p"},"define Paper size or Sheet size"),"."),(0,r.kt)("p",null,"I used A4 Sheet with vertical orientation by using ",(0,r.kt)("inlineCode",{parentName:"p"},"swDwgPaperSizes_e.swDwgPaperA4sizeVertical")," enumarator."),(0,r.kt)("p",null,"You can use other values from ",(0,r.kt)("a",{parentName:"p",href:"http://help.solidworks.com/2013/English/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swDwgPaperSizes_e.html"},"this list"),"."),(0,r.kt)("hr",null),(0,r.kt)("h3",{id:"solidworks-drawing-document-with-custom-paper-size"},"Solidworks Drawing Document with Custom Paper size"),(0,r.kt)("p",null,"To open a new Drawing with ",(0,r.kt)("em",{parentName:"p"},"Custom sheet size")," we use following code sample:"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Main function of our VBA program\nSub main()\n\n    ' Setting Solidworks variable to Solidworks application\n    Set swApp = Application.SldWorks\n    \n    ' Creating string type variable for storing default drawing location\n    Dim defaultTemplate As String\n    ' Setting value of this string type variable to \"Default drawing template\" with custom paper size\n    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)\n\n    ' Setting Solidworks document to new drawing document\n    Set swDoc = swApp.NewDocument(defaultTemplate, swDwgPaperSizes_e.swDwgPapersUserDefined, 2, 3)\n\nEnd Sub\n")),(0,r.kt)("p",null,"For ",(0,r.kt)("em",{parentName:"p"},"custom paper size"),", we need to use ",(0,r.kt)("inlineCode",{parentName:"p"},"swDwgPaperSizes_e.swDwgPapersUserDefined")," value of paper size."),(0,r.kt)("p",null,"Since we are using custom value, we need to define ",(0,r.kt)("strong",{parentName:"p"},"paper width")," and ",(0,r.kt)("strong",{parentName:"p"},"paper height")," also."),(0,r.kt)("div",{className:"admonition admonition-info alert alert--info"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"14",height:"16",viewBox:"0 0 14 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M7 2.3c3.14 0 5.7 2.56 5.7 5.7s-2.56 5.7-5.7 5.7A5.71 5.71 0 0 1 1.3 8c0-3.14 2.56-5.7 5.7-5.7zM7 1C3.14 1 0 4.14 0 8s3.14 7 7 7 7-3.14 7-7-3.14-7-7-7zm1 3H6v5h2V4zm0 6H6v2h2v-2z"}))),"info")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"It important to remember that API use ",(0,r.kt)("strong",{parentName:"p"},"Metric system")," only. So you need to use the converted value in defining paper width and paper height."))),(0,r.kt)("p",null,"This is all for now. In the next post I will tell you how select a plane in a part document and if possible how to create a skecth segment."))}w.isMDXComponent=!0},74753:function(e,t,n){n.d(t,{Z:function(){return r}});var o=n(94578),a=n(67294),r=function(e){function t(){return e.apply(this,arguments)||this}(0,o.Z)(t,e);var n=t.prototype;return n.componentDidMount=function(){var e;(e=document.createElement("script")).src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild),(window.adsbygoogle=window.adsbygoogle||[]).push({})},n.render=function(){return a.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})},t}(a.Component)}}]);