(window.webpackJsonp=window.webpackJsonp||[]).push([[66],{123:function(e,t,n){"use strict";n.r(t),n.d(t,"frontMatter",(function(){return p})),n.d(t,"metadata",(function(){return i})),n.d(t,"rightToc",(function(){return c})),n.d(t,"default",(function(){return l}));var o=n(2),a=n(6),r=(n(0),n(152)),p={id:"sw-macro-open-saved-documents",title:"Open Saved Documents"},i={unversionedId:"solidworks-macros/sw-macro-open-saved-documents",id:"solidworks-macros/sw-macro-open-saved-documents",isDocsHomePage:!1,title:"Open Saved Documents",description:"In this post, I tell you how to open a saved document in Solidworks using VBA Macro.",source:"@site/docs\\solidworks-macros\\2019-03-02-open-saved-document.md",slug:"/solidworks-macros/sw-macro-open-saved-documents",permalink:"/docs/solidworks-macros/sw-macro-open-saved-documents",version:"current",sidebar:"swvba",previous:{title:"Selection Methods",permalink:"/docs/solidworks-macros/sw-macro-selection-methods"},next:{title:"Fix Unit Issue",permalink:"/docs/solidworks-macros/sw-sketch-macro-fix-unit-issue"}},c=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",children:[]},{value:"By OpenDoc method",id:"by-opendoc-method",children:[]},{value:"By OpenDoc6 method",id:"by-opendoc6-method",children:[]}],b={rightToc:c};function l(e){var t=e.components,n=Object(a.a)(e,["components"]);return Object(r.b)("wrapper",Object(o.a)({},b,n,{components:t,mdxType:"MDXLayout"}),Object(r.b)("p",null,"In this post, I tell you how to ",Object(r.b)("em",{parentName:"p"},"open a saved document")," in Solidworks using ",Object(r.b)("em",{parentName:"p"},"VBA Macro"),"."),Object(r.b)("p",null,"We open document with 2 different methods."),Object(r.b)("ol",null,Object(r.b)("li",{parentName:"ol"},Object(r.b)("p",{parentName:"li"},"By ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method")),Object(r.b)("li",{parentName:"ol"},Object(r.b)("p",{parentName:"li"},"By ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method"))),Object(r.b)("hr",null),Object(r.b)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),Object(r.b)("p",null,"Please see below video on ",Object(r.b)("strong",{parentName:"p"},"how to Open Saved Documents")," from Visual Studio."),Object(r.b)("div",{class:"youtube-responsive-container"},Object(r.b)("iframe",{src:"https://www.youtube.com/embed/DeltLKXAIjY",frameborder:"0",allowfullscreen:!0})),Object(r.b)("p",null,"Please note that there are ",Object(r.b)("strong",{parentName:"p"},"no explaination")," in the video. "),Object(r.b)("p",null,Object(r.b)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),Object(r.b)("hr",null),Object(r.b)("h2",{id:"by-opendoc-method"},"By OpenDoc method"),Object(r.b)("p",null,"This is ",Object(r.b)("strong",{parentName:"p"},"the simplest")," method to open a saved part from your computer."),Object(r.b)("p",null,"In this method we just need two information."),Object(r.b)("ul",null,Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},"Location of document to open")),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},"Type of document which we want to open"))),Object(r.b)("p",null,"Below is the example code for opening a saved document using ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method."),Object(r.b)("pre",null,Object(r.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n      \n  ' Open a saved document\n  Set swDoc = swApp.OpenDoc(\"H:\\Solidworks studies\\API Studies\\Chapter 1 - The Basics\\1st example part.SLDPRT\", swDocumentTypes_e.swDocPART)\n      \n  ' Selecting Front Plane\n  BoolStatus = swDoc.SelectByID(\"Front Plane\", \"PLANE\", 0, 0, 0)\n\nEnd Sub\n")),Object(r.b)("p",null,"I have already explained every line except ",Object(r.b)("em",{parentName:"p"},"middle line")," in above code sample in previous ",Object(r.b)("a",Object(o.a)({parentName:"p"},{href:"sw-macro-open-part"}),"posts"),"."),Object(r.b)("p",null,"To open a saved document we used following lline."),Object(r.b)("pre",null,Object(r.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),'\' Open a saved document\nSet swDoc = swApp.OpenDoc("H:\\Solidworks studies\\API Studies\\Chapter 1 - The Basics\\1st example part.SLDPRT", swDocumentTypes_e.swDocPART)\n')),Object(r.b)("p",null,"Here, we set the ",Object(r.b)("inlineCode",{parentName:"p"},"ModelDoc2")," variable ",Object(r.b)("inlineCode",{parentName:"p"},"swDoc")," to a value."),Object(r.b)("p",null,"This value is ",Object(r.b)("em",{parentName:"p"},"return")," or ",Object(r.b)("em",{parentName:"p"},"provided")," by ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method."),Object(r.b)("p",null,"This method is part of ",Object(r.b)("em",{parentName:"p"},"Solidworks document"),". "),Object(r.b)("p",null,"Since we define ",Object(r.b)("inlineCode",{parentName:"p"},"swApp")," variable as Solidworks document hence we 1st call ",Object(r.b)("inlineCode",{parentName:"p"},"swApp")," and then using ",Object(r.b)("inlineCode",{parentName:"p"},"Dot operator")," we access the ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method."),Object(r.b)("p",null,Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method takes 2 ",Object(r.b)("em",{parentName:"p"},"arguments")," or ",Object(r.b)("em",{parentName:"p"},"parameter"),"."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"FileName")," : Document name or full path if not in current directory, including extension."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"Type")," : Document type as define in ",Object(r.b)("inlineCode",{parentName:"p"},"swDocumentTypes_e")," as follows."),Object(r.b)("ul",null,Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocASSEMBLY"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocDRAWING"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocLAYOUT"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocNONE"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocPART"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocSDM")))),Object(r.b)("div",{className:"admonition admonition-note alert alert--secondary"},Object(r.b)("div",Object(o.a)({parentName:"div"},{className:"admonition-heading"}),Object(r.b)("h5",{parentName:"div"},Object(r.b)("span",Object(o.a)({parentName:"h5"},{className:"admonition-icon"}),Object(r.b)("svg",Object(o.a)({parentName:"span"},{xmlns:"http://www.w3.org/2000/svg",width:"14",height:"16",viewBox:"0 0 14 16"}),Object(r.b)("path",Object(o.a)({parentName:"svg"},{fillRule:"evenodd",d:"M6.3 5.69a.942.942 0 0 1-.28-.7c0-.28.09-.52.28-.7.19-.18.42-.28.7-.28.28 0 .52.09.7.28.18.19.28.42.28.7 0 .28-.09.52-.28.7a1 1 0 0 1-.7.3c-.28 0-.52-.11-.7-.3zM8 7.99c-.02-.25-.11-.48-.31-.69-.2-.19-.42-.3-.69-.31H6c-.27.02-.48.13-.69.31-.2.2-.3.44-.31.69h1v3c.02.27.11.5.31.69.2.2.42.31.69.31h1c.27 0 .48-.11.69-.31.2-.19.3-.42.31-.69H8V7.98v.01zM7 2.3c-3.14 0-5.7 2.54-5.7 5.68 0 3.14 2.56 5.7 5.7 5.7s5.7-2.55 5.7-5.7c0-3.15-2.56-5.69-5.7-5.69v.01zM7 .98c3.86 0 7 3.14 7 7s-3.14 7-7 7-7-3.12-7-7 3.14-7 7-7z"})))),"note")),Object(r.b)("div",Object(o.a)({parentName:"div"},{className:"admonition-content"}),Object(r.b)("p",{parentName:"div"},"If you want to open a Library feature part then we use ",Object(r.b)("inlineCode",{parentName:"p"},"swDocPART")," as document type."))),Object(r.b)("p",null,Object(r.b)("strong",{parentName:"p"},"Return Value")," - If the document opens then this method returns ",Object(r.b)("inlineCode",{parentName:"p"},"True")," and otherwise ",Object(r.b)("inlineCode",{parentName:"p"},"False"),"."),Object(r.b)("p",null,"If you just want to open a saved document then this method is what you are looking for."),Object(r.b)("p",null,"For most of the part, ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method works well."),Object(r.b)("p",null,"If you want more option while opening a document, then next method is for you."),Object(r.b)("hr",null),Object(r.b)("h2",{id:"by-opendoc6-method"},"By OpenDoc6 method"),Object(r.b)("p",null,Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method is extension to ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," with some additional parameters."),Object(r.b)("p",null,"How ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," works is shown in below code sample:"),Object(r.b)("pre",null,Object(r.b)("code",Object(o.a)({parentName:"pre"},{className:"language-vb"}),'Option Explicit\n\n\' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n\' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\' Boolean Variable\nDim BoolStatus As Boolean\n\n\n\' Main function of our VBA program\nSub main()\n\n  \' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n      \n  \' Open an saved document\n  Set swDoc = swApp.OpenDoc6("H:\\Solidworks studies\\API Studies\\Chapter 1 - The Basics\\1st example part.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)\n      \n  \' Selecting Front Plane\n  BoolStatus = swDoc.SelectByID("Front Plane", "PLANE", 0, 0, 0)\n\nEnd Sub\n')),Object(r.b)("p",null,"This code sample is similar is to previous example code except for ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method takes extra 4 parameters."),Object(r.b)("p",null,Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method takes 6 ",Object(r.b)("em",{parentName:"p"},"arguments")," or ",Object(r.b)("em",{parentName:"p"},"parameter"),"."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"FileName")," : Document name or full path if not in current directory, including extension."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"Type")," : Document type as define in ",Object(r.b)("inlineCode",{parentName:"p"},"swDocumentTypes_e")," as follows."),Object(r.b)("ul",null,Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocASSEMBLY"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocDRAWING"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocLAYOUT"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocNONE"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocPART"))),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},Object(r.b)("inlineCode",{parentName:"p"},"swDocSDM")))),Object(r.b)("div",{className:"admonition admonition-note alert alert--secondary"},Object(r.b)("div",Object(o.a)({parentName:"div"},{className:"admonition-heading"}),Object(r.b)("h5",{parentName:"div"},Object(r.b)("span",Object(o.a)({parentName:"h5"},{className:"admonition-icon"}),Object(r.b)("svg",Object(o.a)({parentName:"span"},{xmlns:"http://www.w3.org/2000/svg",width:"14",height:"16",viewBox:"0 0 14 16"}),Object(r.b)("path",Object(o.a)({parentName:"svg"},{fillRule:"evenodd",d:"M6.3 5.69a.942.942 0 0 1-.28-.7c0-.28.09-.52.28-.7.19-.18.42-.28.7-.28.28 0 .52.09.7.28.18.19.28.42.28.7 0 .28-.09.52-.28.7a1 1 0 0 1-.7.3c-.28 0-.52-.11-.7-.3zM8 7.99c-.02-.25-.11-.48-.31-.69-.2-.19-.42-.3-.69-.31H6c-.27.02-.48.13-.69.31-.2.2-.3.44-.31.69h1v3c.02.27.11.5.31.69.2.2.42.31.69.31h1c.27 0 .48-.11.69-.31.2-.19.3-.42.31-.69H8V7.98v.01zM7 2.3c-3.14 0-5.7 2.54-5.7 5.68 0 3.14 2.56 5.7 5.7 5.7s5.7-2.55 5.7-5.7c0-3.15-2.56-5.69-5.7-5.69v.01zM7 .98c3.86 0 7 3.14 7 7s-3.14 7-7 7-7-3.12-7-7 3.14-7 7-7z"})))),"note")),Object(r.b)("div",Object(o.a)({parentName:"div"},{className:"admonition-content"}),Object(r.b)("p",{parentName:"div"},"If you want to open a Library feature part then we use ",Object(r.b)("inlineCode",{parentName:"p"},"swDocPART")," as document type."))),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"Options")," : Mode in which to open the document as defined in ",Object(r.b)("inlineCode",{parentName:"p"},"swOpenDocOptions_e"),"."),Object(r.b)("p",null,"For more details about ",Object(r.b)("em",{parentName:"p"},"Options")," parameters, please visit ",Object(r.b)("a",Object(o.a)({parentName:"p"},{href:"help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swOpenDocOptions_e.html"}),"this page"),"."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"Configuration")," : Configuration in which you want to open this document."),Object(r.b)("ul",null,Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},"Applies to ",Object(r.b)("em",{parentName:"p"},"Part")," and ",Object(r.b)("em",{parentName:"p"},"Assemblies"),", not ",Object(r.b)("em",{parentName:"p"},"drawings"),".")),Object(r.b)("li",{parentName:"ul"},Object(r.b)("p",{parentName:"li"},"If this argument is ",Object(r.b)("em",{parentName:"p"},"empty")," or the specified configuration is ",Object(r.b)("em",{parentName:"p"},"not present")," in the model, the model is opened in the last-used configuration."))),Object(r.b)("p",null,"I used an ",Object(r.b)("inlineCode",{parentName:"p"},'""')," in the above code sample, because I want to open part file in last saved configuration."),Object(r.b)("p",null,"If you don't know about ",Object(r.b)("inlineCode",{parentName:"p"},'""'),", then this symbol represent an ",Object(r.b)("strong",{parentName:"p"},"empty string"),"."),Object(r.b)("p",null,"When we don't want to pass any value as ",Object(r.b)("inlineCode",{parentName:"p"},"string"),", at that time I use ",Object(r.b)("inlineCode",{parentName:"p"},'""'),"."),Object(r.b)("p",null,"You can also use ",Object(r.b)("inlineCode",{parentName:"p"},'""')," when you want to pass an empty string in VBA."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"Errors")," : Load errors as defined in ",Object(r.b)("inlineCode",{parentName:"p"},"swFileLoadError_e"),"."),Object(r.b)("p",null,"For more details about ",Object(r.b)("em",{parentName:"p"},"Errors")," parameters, please visit ",Object(r.b)("a",Object(o.a)({parentName:"p"},{href:"http://help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swFileLoadError_e.html"}),"this page"),"."),Object(r.b)("p",null,"Since this parameter is ",Object(r.b)("inlineCode",{parentName:"p"},"long")," type, and I don't want to enter any value from the provided list; I used ",Object(r.b)("strong",{parentName:"p"},"0")," as value."),Object(r.b)("p",null,"If you want to use options from option link then you can use values from there."),Object(r.b)("p",null,"It is just I don't want to load any error information about the part."),Object(r.b)("p",null,Object(r.b)("em",{parentName:"p"},"Warnings")," : Warnings or extra information generated during the open operation as defined in ",Object(r.b)("inlineCode",{parentName:"p"},"swFileLoadWarning_e"),"."),Object(r.b)("p",null,"For more details about ",Object(r.b)("em",{parentName:"p"},"Warnings")," parameters, please visit ",Object(r.b)("a",Object(o.a)({parentName:"p"},{href:"http://help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swFileLoadWarning_e.html"}),"this page"),"."),Object(r.b)("p",null,"As in the previous parameter, I use ",Object(r.b)("strong",{parentName:"p"},"0")," as value."),Object(r.b)("p",null,Object(r.b)("strong",{parentName:"p"},"Return Value")," - If the document opens then this method returns ",Object(r.b)("inlineCode",{parentName:"p"},"True")," and otherwise ",Object(r.b)("inlineCode",{parentName:"p"},"False"),"."),Object(r.b)("p",null,"As you can see, in ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method, we need to defined the extra parameters compared to ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc")," method."),Object(r.b)("p",null,"It is worth noted that, ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method is the most updated method for opening a saved document."),Object(r.b)("p",null,"Hence if did not use any of the above method, I would recommend you to use ",Object(r.b)("inlineCode",{parentName:"p"},"OpenDoc6")," method."),Object(r.b)("p",null,"Hope this post helps you to understand opening methods with Solidworks VB Macros."),Object(r.b)("p",null,"For more such tutorials on Solidworks VBA Macros, do come to this blog after sometime."),Object(r.b)("p",null,"Till then, Happy learning!!!"))}l.isMDXComponent=!0},152:function(e,t,n){"use strict";n.d(t,"a",(function(){return s})),n.d(t,"b",(function(){return u}));var o=n(0),a=n.n(o);function r(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function p(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?p(Object(n),!0).forEach((function(t){r(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):p(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function c(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},r=Object.keys(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var b=a.a.createContext({}),l=function(e){var t=a.a.useContext(b),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},s=function(e){var t=l(e.components);return a.a.createElement(b.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return a.a.createElement(a.a.Fragment,{},t)}},d=a.a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,r=e.originalType,p=e.parentName,b=c(e,["components","mdxType","originalType","parentName"]),s=l(n),d=o,u=s["".concat(p,".").concat(d)]||s[d]||m[d]||r;return n?a.a.createElement(u,i(i({ref:t},b),{},{components:n})):a.a.createElement(u,i({ref:t},b))}));function u(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var r=n.length,p=new Array(r);p[0]=d;var i={};for(var c in t)hasOwnProperty.call(t,c)&&(i[c]=t[c]);i.originalType=e,i.mdxType="string"==typeof e?e:o,p[1]=i;for(var b=2;b<r;b++)p[b]=n[b];return a.a.createElement.apply(null,p)}return a.a.createElement.apply(null,n)}d.displayName="MDXCreateElement"}}]);