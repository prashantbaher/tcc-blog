"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[7605],{3905:(e,t,n)=>{n.d(t,{Zo:()=>p,kt:()=>d});var a=n(67294);function r(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function s(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?s(Object(n),!0).forEach((function(t){r(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):s(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function o(e,t){if(null==e)return{};var n,a,r=function(e,t){if(null==e)return{};var n,a,r={},s=Object.keys(e);for(a=0;a<s.length;a++)n=s[a],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var s=Object.getOwnPropertySymbols(e);for(a=0;a<s.length;a++)n=s[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var i=a.createContext({}),m=function(e){var t=a.useContext(i),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},p=function(e){var t=m(e.components);return a.createElement(i.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},k=a.forwardRef((function(e,t){var n=e.components,r=e.mdxType,s=e.originalType,i=e.parentName,p=o(e,["components","mdxType","originalType","parentName"]),k=m(n),d=r,c=k["".concat(i,".").concat(d)]||k[d]||u[d]||s;return n?a.createElement(c,l(l({ref:t},p),{},{components:n})):a.createElement(c,l({ref:t},p))}));function d(e,t){var n=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var s=n.length,l=new Array(s);l[0]=k;var o={};for(var i in t)hasOwnProperty.call(t,i)&&(o[i]=t[i]);o.originalType=e,o.mdxType="string"==typeof e?e:r,l[1]=o;for(var m=2;m<s;m++)l[m]=n[m];return a.createElement.apply(null,l)}return a.createElement.apply(null,n)}k.displayName="MDXCreateElement"},66514:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>m,contentTitle:()=>o,default:()=>k,frontMatter:()=>l,metadata:()=>i,toc:()=>p});var a=n(87462),r=(n(67294),n(3905)),s=n(74753);const l={categories:"Solidworks-macro",title:"Solidworks VBA Macro - Insert Virtual Assembly",permalink:"/solidworks-vba-macros/assembly-insert-virtual-assembly/",tags:["Solidworks Macro"],id:"assembly-insert-virtual-assembly"},o=void 0,i={unversionedId:"assembly-insert-virtual-assembly",id:"assembly-insert-virtual-assembly",title:"Solidworks VBA Macro - Insert Virtual Assembly",description:"",source:"@site/docs/solidworks-macros/025.2-assembly-insert-virtual-assembly.md",sourceDirName:".",slug:"/assembly-insert-virtual-assembly",permalink:"/solidworks-macros/assembly-insert-virtual-assembly",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks VBA Macro - Insert Virtual Assembly",permalink:"/solidworks-vba-macros/assembly-insert-virtual-assembly/",tags:["Solidworks Macro"],id:"assembly-insert-virtual-assembly"},sidebar:"tutorialSidebar",previous:{title:"Solidworks VBA Macro - Insert Virtual Part",permalink:"/solidworks-macros/assembly-insert-virtual-part"},next:{title:"Solidworks VBA Macro - Copy With Mate",permalink:"/solidworks-macros/assembly-copy-with-mate"}},m={},p=[{value:"Objective",id:"objective",level:2},{value:"Results We Can Get",id:"results-we-can-get",level:2},{value:"Macro Video",id:"macro-video",level:2},{value:"VBA Macro",id:"vba-macro",level:2},{value:"Prerequisite",id:"prerequisite",level:2},{value:"Steps To Follow",id:"steps-to-follow",level:2},{value:"Create Global Variables",id:"create-global-variables",level:3},{value:"Initialize Global Variables",id:"initialize-global-variables",level:3},{value:"Insert Virtual Assembly",id:"insert-virtual-assembly",level:3}],u={toc:p};function k(e){let{components:t,...l}=e;return(0,r.kt)("wrapper",(0,a.Z)({},u,l,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("h2",{id:"objective"},"Objective"),(0,r.kt)(s.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,'In this article, we understand "how to" ',(0,r.kt)("strong",{parentName:"p"},"Insert Virtual Assembly")," in ",(0,r.kt)("strong",{parentName:"p"},"Assembly document")," from VBA macro."),(0,r.kt)("p",null,"This is most updated method of ",(0,r.kt)("strong",{parentName:"p"},"inserting Virtual Assembly")," in an assembly document."),(0,r.kt)("h2",{id:"results-we-can-get"},"Results We Can Get"),(0,r.kt)("p",null,"Below image shows the result we get."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:n(65112).Z},(0,r.kt)("img",{alt:"assembly-insert-virtual-assembly",src:n(52175).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,"We ",(0,r.kt)("strong",{parentName:"p"},"Insert Virtual Assembly")," in simple manners."),(0,r.kt)("p",null,"There are no extra steps required."),(0,r.kt)("admonition",{type:"caution"},(0,r.kt)("p",{parentName:"admonition"},"To get the correct result, please follow the steps correctly.")),(0,r.kt)("h2",{id:"macro-video"},"Macro Video"),(0,r.kt)("p",null,"Below \ud83c\udfac video shows how to ",(0,r.kt)("strong",{parentName:"p"},"Insert Virtual Assembly")," from ",(0,r.kt)("em",{parentName:"p"},"SOLIDWORKS VBA Macros"),"."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/j0N1NvzW_Pc",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Above video is just for visualization and there is no explanation."))," "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"I have explained every line in this article."))),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"It is advisable to watch video, since it helps you to better understand the process.")),(0,r.kt)("h2",{id:"vba-macro"},"VBA Macro"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"VBA macro"))," for ",(0,r.kt)("em",{parentName:"p"},"Insert Virtual Assembly"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Variable for Solidworks Application\nDim swApp As SldWorks.SldWorks\n\n' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n\n' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n\n' Program to Insert virtual Assembly\nSub main()\n  \n  ' Set Solidworks Application variable to current application\n  Set swApp = Application.SldWorks\n  \n  ' Set Solidworks document variable to currently opened document\n  Set swDoc = swApp.ActiveDoc\n  \n  ' Check if Solidworks document is opened or not\n  If swDoc Is Nothing Then\n    MsgBox \"Solidworks document is not opened.\"\n    Exit Sub\n  End If\n  \n  ' Set Solidworks Assembly document\n  Set swAssembly = swDoc\n  \n  ' Insert Virtual Assembly\n  swAssembly.InsertNewVirtualAssembly swComponent\n  \n  ' If there are error\n  If swComponent Is Nothing Then\n    ' Inform user and exit function.\n    MsgBox \"Failed to add Virtual Assembly.\"\n    Exit Sub\n  End If\n  \nEnd Sub\n")),(0,r.kt)(s.Z,{mdxType:"AdComponent"}),(0,r.kt)("h2",{id:"prerequisite"},"Prerequisite"),(0,r.kt)("p",null,"There are some ",(0,r.kt)("em",{parentName:"p"},"prerequisites")," for this article."),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"Knowledge of ",(0,r.kt)("strong",{parentName:"li"},"VBA programming language")," is \u2757",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("em",{parentName:"strong"},"required")),"."),(0,r.kt)("li",{parentName:"ul"},"We use existing parts in Assembly document."),(0,r.kt)("li",{parentName:"ul"},"Both components are fully constraint as shown in below image.")),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:n(90135).Z},(0,r.kt)("img",{alt:"prerequisite",src:n(46222).Z,width:"1366",height:"728"}))),(0,r.kt)("admonition",{type:"note"},(0,r.kt)("p",{parentName:"admonition"},"We will apply checks in this article, so the code we write, should be ",(0,r.kt)("strong",{parentName:"p"},"error free")," mostly.")),(0,r.kt)("h2",{id:"steps-to-follow"},"Steps To Follow"),(0,r.kt)("p",null,"This ",(0,r.kt)("strong",{parentName:"p"},"VBA macro")," can be divided into following sections:"),(0,r.kt)("ol",null,(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Create Global Variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Initialize Global Variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Insert Virtual Assembly"))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Every section with each line is explained below."))),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"I also give some ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"links (see icon \ud83d\ude80)"))," so that you can go through them if there are anything I explained in previous articles.")),(0,r.kt)("h3",{id:"create-global-variables"},"Create Global Variables"),(0,r.kt)("p",null,"In this section, we create global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": Above line forces us to define every variable we are going to use. "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"SOLIDWORKS Macros - Open new Part document"))," article.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks application"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swApp")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.SldWorks")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks document"),". "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swDoc")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.ModelDoc2")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Assembly"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swAssembly")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.AssemblyDoc")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),":  In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Component"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swComponent")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.Component2"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"These all are our global variables."))),(0,r.kt)("p",null,"They are ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS API Objects"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Program to Insert virtual Assembly\nSub main()\n\nEnd Sub\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create ",(0,r.kt)("em",{parentName:"li"},"main Program to Insert virtual Assembly"),"."),(0,r.kt)("li",{parentName:"ul"},"This is a ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Sub"))," procedure which has name of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"main")),". "),(0,r.kt)("li",{parentName:"ul"},"This procedure hold all the ",(0,r.kt)("em",{parentName:"li"},"statements (instructions)")," we give to computer."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Detailed information \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-sub-and-function-procedure/"},"VBA Sub and Function Procedures"))," article of this website.")),(0,r.kt)(s.Z,{mdxType:"AdComponent"}),(0,r.kt)("h3",{id:"initialize-global-variables"},"Initialize Global Variables"),(0,r.kt)("p",null,"In this section, we initialize global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Application variable to current application\nSet swApp = Application.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swApp"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently opened Solidworks application.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks document variable to currently opened document\nSet swDoc = swApp.ActiveDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently ",(0,r.kt)("em",{parentName:"li"},"opened part document"),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Check if Solidworks document is opened or not\nIf swDoc Is Nothing Then\n  MsgBox ("Solidworks document is not opened.")\n  Exit Sub\nEnd If\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully set the value of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc Is Nothing"))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": ",(0,r.kt)("em",{parentName:"li"},"SOLIDWORKS document is not opened.")),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Assembly document\nSet swAssembly = swDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable.")),(0,r.kt)("h3",{id:"insert-virtual-assembly"},"Insert Virtual Assembly"),(0,r.kt)("p",null,"In this section, we ",(0,r.kt)("em",{parentName:"p"},"Insert Virtual Assembly"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Insert Virtual Assembly\nswAssembly.InsertNewVirtualAssembly swComponent\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"In above code, we ",(0,r.kt)("strong",{parentName:"p"},"Insert Virtual Assembly")," into assemly.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"For this, we use ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("inlineCode",{parentName:"strong"},"InsertNewVirtualAssembly"))," method.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"This ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("inlineCode",{parentName:"strong"},"InsertNewVirtualAssembly"))," method is part of ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"This method takes 1 parameter."),(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"InsertedComponent"),": ",(0,r.kt)("em",{parentName:"li"},"New assembly inserted as virtual component.")))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"Return Value")," : This ",(0,r.kt)("inlineCode",{parentName:"p"},"InsertNewVirtualAssembly")," method return \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swInsertNewPartErrorCode_e.html"},"Error as defined by swInsertNewPartErrorCode_e")),".")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"In our code, I have used following values:"),(0,r.kt)("table",{parentName:"li"},(0,r.kt)("thead",{parentName:"table"},(0,r.kt)("tr",{parentName:"thead"},(0,r.kt)("th",{parentName:"tr",align:null},"Parameter Name"),(0,r.kt)("th",{parentName:"tr",align:null},"Value Used"))),(0,r.kt)("tbody",{parentName:"table"},(0,r.kt)("tr",{parentName:"tbody"},(0,r.kt)("td",{parentName:"tr",align:null},(0,r.kt)("strong",{parentName:"td"},"InsertedComponent")),(0,r.kt)("td",{parentName:"tr",align:null},(0,r.kt)("inlineCode",{parentName:"td"},"swComponent")))))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"Reference"),": For more details please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.iassemblydoc~insertnewvirtualassembly.html"},"online SOLIDWORKS API Help")),"."))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' If there are error\nIf swComponent Is Nothing Then\n  ' Inform user and exit function.\n  MsgBox \"Failed to add Virtual Assembly.\"\n  Exit Sub\nEnd If\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully added ",(0,r.kt)("strong",{parentName:"li"},"Virtual Assembly")," or not."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent Is Nothing"))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": *Failed to add Virtual Assembly."),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("p",null,"Now we run the macro and after running macro we get ",(0,r.kt)("strong",{parentName:"p"},"a New Virtual Assembly")," as shown in below image."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:n(65112).Z},(0,r.kt)("img",{alt:"assembly-insert-virtual-assembly",src:n(52175).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"This is it !!!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")," \ud83d\ude0a"),(0,r.kt)("p",null,"If you found anything to ",(0,r.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,r.kt)("em",{parentName:"p"},"e-mail")," \ud83d\udce7."),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("strong",{parentName:"p"},"Insert Virtual Assembly")," with SOLIDWORKS VBA Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS VBA Macro"),", do come to this website after sometime."),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")," \ud83d\ude4f\ud83c\udffb"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}k.isMDXComponent=!0},74753:(e,t,n)=>{n.d(t,{Z:()=>r});var a=n(67294);class r extends a.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return a.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}},65112:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/files/final-result-gif-6fd85996c2632c192cddd792a3cda7f8.gif"},90135:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/files/prerequisite-ac6d80f81a4c3adbb62d3681210938ed.png"},52175:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/images/final-result-gif-6fd85996c2632c192cddd792a3cda7f8.gif"},46222:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/images/prerequisite-ac6d80f81a4c3adbb62d3681210938ed.png"}}]);