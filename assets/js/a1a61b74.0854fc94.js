"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[956],{3905:(e,t,n)=>{n.d(t,{Zo:()=>m,kt:()=>c});var o=n(67294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function p(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function i(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},r=Object.keys(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(o=0;o<r.length;o++)n=r[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var l=o.createContext({}),s=function(e){var t=o.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):p(p({},t),e)),n},m=function(e){var t=s(e.components);return o.createElement(l.Provider,{value:t},e.children)},d={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},u=o.forwardRef((function(e,t){var n=e.components,a=e.mdxType,r=e.originalType,l=e.parentName,m=i(e,["components","mdxType","originalType","parentName"]),u=s(n),c=a,k=u["".concat(l,".").concat(c)]||u[c]||d[c]||r;return n?o.createElement(k,p(p({ref:t},m),{},{components:n})):o.createElement(k,p({ref:t},m))}));function c(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var r=n.length,p=new Array(r);p[0]=u;var i={};for(var l in t)hasOwnProperty.call(t,l)&&(i[l]=t[l]);i.originalType=e,i.mdxType="string"==typeof e?e:a,p[1]=i;for(var s=2;s<r;s++)p[s]=n[s];return o.createElement.apply(null,p)}return o.createElement.apply(null,n)}u.displayName="MDXCreateElement"},30344:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>s,contentTitle:()=>i,default:()=>u,frontMatter:()=>p,metadata:()=>l,toc:()=>m});var o=n(87462),a=(n(67294),n(3905)),r=n(74753);const p={categories:"Solidworks-macro",title:"Solidworks Macro - Open Saved Documents",permalink:"/solidworks-macros/open-saved-document/",tags:["Solidworks Macro"],id:"open-saved-document"},i=void 0,l={unversionedId:"open-saved-document",id:"open-saved-document",title:"Solidworks Macro - Open Saved Documents",description:"",source:"@site/docs/solidworks-macros/001.5-open-saved-document.md",sourceDirName:".",slug:"/open-saved-document",permalink:"/solidworks-macros/open-saved-document",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks Macro - Open Saved Documents",permalink:"/solidworks-macros/open-saved-document/",tags:["Solidworks Macro"],id:"open-saved-document"},sidebar:"tutorialSidebar",previous:{title:"Solidworks Macro - Selection Methods",permalink:"/solidworks-macros/select-plane-from-tree"},next:{title:"Solidworks Macro - Fix Unit Issue",permalink:"/solidworks-macros/unit-correction"}},s={},m=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"By OpenDoc method",id:"by-opendoc-method",level:2},{value:"By OpenDoc6 method",id:"by-opendoc6-method",level:2}],d={toc:m};function u(e){let{components:t,...n}=e;return(0,a.kt)("wrapper",(0,o.Z)({},d,n,{components:t,mdxType:"MDXLayout"}),(0,a.kt)(r.Z,{mdxType:"AdComponent"}),(0,a.kt)("p",null,"In this post, I tell you how to ",(0,a.kt)("em",{parentName:"p"},"open a saved document")," in Solidworks using ",(0,a.kt)("em",{parentName:"p"},"VBA Macro"),"."),(0,a.kt)("p",null,"We open document with 2 different methods."),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"By ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"By ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method"))),(0,a.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,a.kt)("p",null,"Please see below video \ud83c\udfac on ",(0,a.kt)("strong",{parentName:"p"},"how to Open Saved Documents")," from Visual Studio."),(0,a.kt)("iframe",{src:"https://www.youtube.com/embed/DeltLKXAIjY",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,a.kt)("p",null,"Please note that there are ",(0,a.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,a.kt)("h2",{id:"by-opendoc-method"},"By OpenDoc method"),(0,a.kt)("p",null,"This is ",(0,a.kt)("strong",{parentName:"p"},"the simplest")," method to open a saved part from your computer."),(0,a.kt)("p",null,"In this method we just need two information."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"Location of document to open")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"Type of document which we want to open"))),(0,a.kt)("p",null,"Below is the example code for opening a saved document using ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method."),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n' Boolean Variable\nDim BoolStatus As Boolean\n\n\n' Main function of our VBA program\nSub main()\n\n  ' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n      \n  ' Open a saved document\n  Set swDoc = swApp.OpenDoc(\"H:\\Solidworks studies\\API Studies\\Chapter 1 - The Basics\\1st example part.SLDPRT\", swDocumentTypes_e.swDocPART)\n      \n  ' Selecting Front Plane\n  BoolStatus = swDoc.SelectByID(\"Front Plane\", \"PLANE\", 0, 0, 0)\n\nEnd Sub\n")),(0,a.kt)("p",null,"I have already explained every line except ",(0,a.kt)("em",{parentName:"p"},"middle line")," in above code sample in previous posts."),(0,a.kt)("p",null,"To open a saved document we used following lline."),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Open a saved document\nSet swDoc = swApp.OpenDoc("H:\\Solidworks studies\\API Studies\\Chapter 1 - The Basics\\1st example part.SLDPRT", swDocumentTypes_e.swDocPART)\n')),(0,a.kt)("p",null,"Here, we set the ",(0,a.kt)("inlineCode",{parentName:"p"},"ModelDoc2")," variable ",(0,a.kt)("inlineCode",{parentName:"p"},"swDoc")," to a value."),(0,a.kt)("p",null,"This value is ",(0,a.kt)("em",{parentName:"p"},"return")," or ",(0,a.kt)("em",{parentName:"p"},"provided")," by ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method."),(0,a.kt)("p",null,"This method is part of ",(0,a.kt)("em",{parentName:"p"},"Solidworks document"),". "),(0,a.kt)("p",null,"Since we define ",(0,a.kt)("inlineCode",{parentName:"p"},"swApp")," variable as Solidworks document hence we 1st call ",(0,a.kt)("inlineCode",{parentName:"p"},"swApp")," and then using ",(0,a.kt)("inlineCode",{parentName:"p"},"Dot operator")," we access the ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method."),(0,a.kt)("p",null,(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method takes 2 ",(0,a.kt)("em",{parentName:"p"},"arguments")," or ",(0,a.kt)("em",{parentName:"p"},"parameter"),"."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"FileName")," : Document name or full path if not in current directory, including extension."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Type")," : Document type as define in ",(0,a.kt)("inlineCode",{parentName:"p"},"swDocumentTypes_e")," as follows."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocASSEMBLY"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocDRAWING"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocLAYOUT"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocNONE"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocPART"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocSDM")))),(0,a.kt)("admonition",{type:"info"},(0,a.kt)("p",{parentName:"admonition"},"If you want to open a Library feature part then we use ",(0,a.kt)("inlineCode",{parentName:"p"},"swDocPART")," as document type.")),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Return Value")," - If the document opens then this method returns ",(0,a.kt)("inlineCode",{parentName:"p"},"True")," and otherwise ",(0,a.kt)("inlineCode",{parentName:"p"},"False"),"."),(0,a.kt)("p",null,"If you just want to open a saved document then this method is what you are looking for."),(0,a.kt)("p",null,"For most of the part, ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method works well."),(0,a.kt)("p",null,"If you want more option while opening a document, then next method is for you."),(0,a.kt)(r.Z,{mdxType:"AdComponent"}),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"by-opendoc6-method"},"By OpenDoc6 method"),(0,a.kt)("p",null,(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method is extension to ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," with some additional parameters."),(0,a.kt)("p",null,"How ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," works is shown in below code sample:"),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'Option Explicit\n\n\' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n\' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\' Boolean Variable\nDim BoolStatus As Boolean\n\n\n\' Main function of our VBA program\nSub main()\n\n  \' Setting Solidworks variable to Solidworks application\n  Set swApp = Application.SldWorks\n      \n  \' Open an saved document\n  Set swDoc = swApp.OpenDoc6("H:\\Solidworks studies\\API Studies\\Chapter 1 - The Basics\\1st example part.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)\n      \n  \' Selecting Front Plane\n  BoolStatus = swDoc.SelectByID("Front Plane", "PLANE", 0, 0, 0)\n\nEnd Sub\n')),(0,a.kt)("p",null,"This code sample is similar is to previous example code except for ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method takes extra 4 parameters."),(0,a.kt)("p",null,(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method takes 6 ",(0,a.kt)("em",{parentName:"p"},"arguments")," or ",(0,a.kt)("em",{parentName:"p"},"parameter"),"."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"FileName")," : Document name or full path if not in current directory, including extension."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Type")," : Document type as define in ",(0,a.kt)("inlineCode",{parentName:"p"},"swDocumentTypes_e")," as follows."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocASSEMBLY"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocDRAWING"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocLAYOUT"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocNONE"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocPART"))),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("inlineCode",{parentName:"p"},"swDocSDM")))),(0,a.kt)("admonition",{type:"info"},(0,a.kt)("p",{parentName:"admonition"},"f you want to open a Library feature part then we use ",(0,a.kt)("inlineCode",{parentName:"p"},"swDocPART")," as document type.")),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Options")," : Mode in which to open the document as defined in ",(0,a.kt)("inlineCode",{parentName:"p"},"swOpenDocOptions_e"),"."),(0,a.kt)("p",null,"For more details about ",(0,a.kt)("em",{parentName:"p"},"Options")," parameters, please visit \ud83d\ude80 ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swOpenDocOptions_e.html"},"this page of Solidworks API Help")),"."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Configuration")," : Configuration in which you want to open this document."),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"Applies to ",(0,a.kt)("em",{parentName:"p"},"Part")," and ",(0,a.kt)("em",{parentName:"p"},"Assemblies"),", not ",(0,a.kt)("em",{parentName:"p"},"drawings"),".")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"If this argument is ",(0,a.kt)("em",{parentName:"p"},"empty")," or the specified configuration is ",(0,a.kt)("em",{parentName:"p"},"not present")," in the model, the model is opened in the last-used configuration."))),(0,a.kt)("p",null,"I used an ",(0,a.kt)("inlineCode",{parentName:"p"},'""')," in the above code sample, because I want to open part file in last saved configuration."),(0,a.kt)("p",null,"If you don't know about ",(0,a.kt)("inlineCode",{parentName:"p"},'""'),", then this symbol represent an ",(0,a.kt)("strong",{parentName:"p"},"empty string"),"."),(0,a.kt)("p",null,"When we don't want to pass any value as ",(0,a.kt)("inlineCode",{parentName:"p"},"string"),", at that time I use ",(0,a.kt)("inlineCode",{parentName:"p"},'""'),"."),(0,a.kt)("p",null,"You can also use ",(0,a.kt)("inlineCode",{parentName:"p"},'""')," when you want to pass an empty string in VBA."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Errors")," : Load errors as defined in ",(0,a.kt)("inlineCode",{parentName:"p"},"swFileLoadError_e"),"."),(0,a.kt)("p",null,"For more details about ",(0,a.kt)("em",{parentName:"p"},"Errors")," parameters, please visit \ud83d\ude80 ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swFileLoadError_e.html"},"this page of Solidworks API Help")),"."),(0,a.kt)("p",null,"Since this parameter is ",(0,a.kt)("inlineCode",{parentName:"p"},"long")," type, and I don't want to enter any value from the provided list; I used ",(0,a.kt)("strong",{parentName:"p"},"0")," as value."),(0,a.kt)("p",null,"If you want to use options from option link then you can use values from there."),(0,a.kt)("p",null,"It is just I don't want to load any error information about the part."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Warnings")," : Warnings or extra information generated during the open operation as defined in ",(0,a.kt)("inlineCode",{parentName:"p"},"swFileLoadWarning_e"),"."),(0,a.kt)("p",null,"For more details about ",(0,a.kt)("em",{parentName:"p"},"Warnings")," parameters, please visit \ud83d\ude80 ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2017/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swFileLoadWarning_e.html"},"this page of Solidworks API Help"),".")),(0,a.kt)("p",null,"As in the previous parameter, I use ",(0,a.kt)("strong",{parentName:"p"},"0")," as value."),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Return Value")," - If the document opens then this method returns ",(0,a.kt)("inlineCode",{parentName:"p"},"True")," and otherwise ",(0,a.kt)("inlineCode",{parentName:"p"},"False"),"."),(0,a.kt)("p",null,"As you can see, in ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method, we need to defined the extra parameters compared to ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc")," method."),(0,a.kt)("p",null,"It is worth noted that, ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method is the most updated method for opening a saved document."),(0,a.kt)("p",null,"Hence if did not use any of the above method, I would recommend you to use ",(0,a.kt)("inlineCode",{parentName:"p"},"OpenDoc6")," method."),(0,a.kt)("p",null,"Hope this post helps you to understand opening methods with Solidworks VB Macros."),(0,a.kt)("p",null,"For more such tutorials on Solidworks VBA Macros, do come to this blog after sometime."),(0,a.kt)("p",null,"Till then, Happy learning!!!"))}u.isMDXComponent=!0},74753:(e,t,n)=>{n.d(t,{Z:()=>a});var o=n(67294);class a extends o.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return o.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}}}]);