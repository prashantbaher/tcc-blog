"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3843],{3905:(e,t,n)=>{n.d(t,{Zo:()=>p,kt:()=>c});var a=n(67294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,a,o=function(e,t){if(null==e)return{};var n,a,o={},r=Object.keys(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var i=a.createContext({}),m=function(e){var t=a.useContext(i),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},p=function(e){var t=m(e.components);return a.createElement(i.Provider,{value:t},e.children)},k={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},u=a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,r=e.originalType,i=e.parentName,p=s(e,["components","mdxType","originalType","parentName"]),u=m(n),c=o,d=u["".concat(i,".").concat(c)]||u[c]||k[c]||r;return n?a.createElement(d,l(l({ref:t},p),{},{components:n})):a.createElement(d,l({ref:t},p))}));function c(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var r=n.length,l=new Array(r);l[0]=u;var s={};for(var i in t)hasOwnProperty.call(t,i)&&(s[i]=t[i]);s.originalType=e,s.mdxType="string"==typeof e?e:o,l[1]=s;for(var m=2;m<r;m++)l[m]=n[m];return a.createElement.apply(null,l)}return a.createElement.apply(null,n)}u.displayName="MDXCreateElement"},88499:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>m,contentTitle:()=>s,default:()=>u,frontMatter:()=>l,metadata:()=>i,toc:()=>p});var a=n(87462),o=(n(67294),n(3905)),r=n(74753);const l={categories:"Solidworks-macro",title:"Solidworks VBA Macro - Show Component",permalink:"/solidworks-vba-macros/assembly-show-component/",tags:["Solidworks Macro"],id:"assembly-show-component"},s=void 0,i={unversionedId:"assembly-show-component",id:"assembly-show-component",title:"Solidworks VBA Macro - Show Component",description:"",source:"@site/docs/solidworks-macros/025.5-assembly-show-component.md",sourceDirName:".",slug:"/assembly-show-component",permalink:"/solidworks-macros/assembly-show-component",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks VBA Macro - Show Component",permalink:"/solidworks-vba-macros/assembly-show-component/",tags:["Solidworks Macro"],id:"assembly-show-component"},sidebar:"tutorialSidebar",previous:{title:"Solidworks VBA Macro - Hide Component",permalink:"/solidworks-macros/assembly-hide-component"},next:{title:"Solidworks VBA Macro - Suppress Component",permalink:"/solidworks-macros/assembly-suppress-component"}},m={},p=[{value:"Objective",id:"objective",level:2},{value:"Results We Can Get",id:"results-we-can-get",level:2},{value:"Macro Video",id:"macro-video",level:2},{value:"VBA Macro",id:"vba-macro",level:2},{value:"Prerequisite",id:"prerequisite",level:2},{value:"Steps To Follow",id:"steps-to-follow",level:2},{value:"Create Global Variables",id:"create-global-variables",level:3},{value:"Initialize Global Variables",id:"initialize-global-variables",level:3},{value:"Show Component",id:"show-component",level:3}],k={toc:p};function u(e){let{components:t,...l}=e;return(0,o.kt)("wrapper",(0,a.Z)({},k,l,{components:t,mdxType:"MDXLayout"}),(0,o.kt)("h2",{id:"objective"},"Objective"),(0,o.kt)(r.Z,{mdxType:"AdComponent"}),(0,o.kt)("p",null,'In this article, we understand "how to" ',(0,o.kt)("strong",{parentName:"p"},"Show Component")," in ",(0,o.kt)("strong",{parentName:"p"},"Assembly document")," from VBA macro."),(0,o.kt)("p",null,"This is most updated method of ",(0,o.kt)("strong",{parentName:"p"},"Show Component")," in an assembly document."),(0,o.kt)("h2",{id:"results-we-can-get"},"Results We Can Get"),(0,o.kt)("p",null,"Below image shows the result we get."),(0,o.kt)("p",null,(0,o.kt)("a",{target:"_blank",href:n(25801).Z},(0,o.kt)("img",{alt:"assembly-show-component",src:n(86565).Z,width:"1366",height:"728"}))),(0,o.kt)("p",null,"We ",(0,o.kt)("strong",{parentName:"p"},"Show Component")," in simple manners."),(0,o.kt)("p",null,"There are no extra steps required."),(0,o.kt)("admonition",{type:"caution"},(0,o.kt)("p",{parentName:"admonition"},"To get the correct result, please follow the steps correctly.")),(0,o.kt)("h2",{id:"macro-video"},"Macro Video"),(0,o.kt)("p",null,"Below \ud83c\udfac video shows how to ",(0,o.kt)("strong",{parentName:"p"},"Show Component")," from ",(0,o.kt)("em",{parentName:"p"},"SOLIDWORKS VBA Macros"),"."),(0,o.kt)("iframe",{src:"https://www.youtube.com/embed/3YCXV0gpN3U",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"Above video is just for visualization and there is no explanation."))," "),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"I have explained every line in this article."))),(0,o.kt)("admonition",{type:"caution"},(0,o.kt)("p",{parentName:"admonition"},"It is advisable to watch video, since it helps you to better understand the process.")),(0,o.kt)("h2",{id:"vba-macro"},"VBA Macro"),(0,o.kt)("p",null,"Below is the ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"VBA macro"))," for ",(0,o.kt)("em",{parentName:"p"},"Show Component"),"."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Variable for Solidworks Application\nDim swApp As SldWorks.SldWorks\n\n' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n\n' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n\n' Program to Show Selected Component\nSub main()\n  \n  ' Set Solidworks Application variable to current application\n  Set swApp = Application.SldWorks\n  \n  ' Set Solidworks document variable to currently opened document\n  Set swDoc = swApp.ActiveDoc\n  \n  ' Check if Solidworks document is opened or not\n  If swDoc Is Nothing Then\n    MsgBox \"Solidworks document is not opened.\"\n    Exit Sub\n  End If\n  \n  ' Set Solidworks Assembly document\n  Set swAssembly = swDoc\n  \n  ' Variable for List of elements\n  Dim vArray As Variant\n  \n  ' Get Components list in opened assembly\n  vArray = swAssembly.GetComponents(True)\n  \n  ' Variable for component\n  Dim component As Variant\n  \n  ' Loop Components List\n  For Each component In vArray\n  \n    ' Set Solidworks Component variable\n    Set swComponent = component\n    \n    ' If current component is hidden\n    If swComponent.IsHidden(False) Then\n      \n      ' Select the component\n      swComponent.Select False\n      \n      ' Show selected component\n      swDoc.ShowComponent2\n    End If\n    \n  Next component\n  \nEnd Sub\n")),(0,o.kt)(r.Z,{mdxType:"AdComponent"}),(0,o.kt)("h2",{id:"prerequisite"},"Prerequisite"),(0,o.kt)("p",null,"There are some ",(0,o.kt)("em",{parentName:"p"},"prerequisites")," for this article."),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"Knowledge of ",(0,o.kt)("strong",{parentName:"li"},"VBA programming language")," is \u2757",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("em",{parentName:"strong"},"required")),"."),(0,o.kt)("li",{parentName:"ul"},"We use existing parts in Assembly document."),(0,o.kt)("li",{parentName:"ul"},"Both components are fully constraint as shown in below image."),(0,o.kt)("li",{parentName:"ul"},"We select the part which we want to Show.")),(0,o.kt)("p",null,(0,o.kt)("a",{target:"_blank",href:n(90306).Z},(0,o.kt)("img",{alt:"prerequisite",src:n(91013).Z,width:"1366",height:"709"}))),(0,o.kt)("admonition",{type:"info"},(0,o.kt)("p",{parentName:"admonition"},"We will apply checks in this article, so the code we write, should be ",(0,o.kt)("strong",{parentName:"p"},"error free")," mostly.")),(0,o.kt)("h2",{id:"steps-to-follow"},"Steps To Follow"),(0,o.kt)(r.Z,{mdxType:"AdComponent"}),(0,o.kt)("p",null,"This ",(0,o.kt)("strong",{parentName:"p"},"VBA macro")," can be divided into following sections:"),(0,o.kt)("ol",null,(0,o.kt)("li",{parentName:"ol"},(0,o.kt)("em",{parentName:"li"},"Create Global Variables")),(0,o.kt)("li",{parentName:"ol"},(0,o.kt)("em",{parentName:"li"},"Initialize Global Variables")),(0,o.kt)("li",{parentName:"ol"},(0,o.kt)("em",{parentName:"li"},"Show Component"))),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"Every section with each line is explained below."))),(0,o.kt)("admonition",{type:"tip"},(0,o.kt)("p",{parentName:"admonition"},"I also give some ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"links (see icon \ud83d\ude80)"))," so that you can go through them if there are anything I explained in previous articles.")),(0,o.kt)("h3",{id:"create-global-variables"},"Create Global Variables"),(0,o.kt)("p",null,"In this section, we create global variables."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Purpose"),": Above line forces us to define every variable we are going to use. "),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Reference"),": \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"SOLIDWORKS Macros - Open new Part document"))," article.")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,o.kt)("em",{parentName:"li"},"Solidworks application"),"."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"swApp")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Type"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"SldWorks.SldWorks")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html"},"online SOLIDWORKS API Help")),".")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,o.kt)("em",{parentName:"li"},"Solidworks document"),". "),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"swDoc")," "),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Type"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"SldWorks.ModelDoc2")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html"},"online SOLIDWORKS API Help")),".")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,o.kt)("em",{parentName:"li"},"Solidworks Assembly"),"."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"swAssembly")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Type"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"SldWorks.AssemblyDoc")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc_members.html"},"online SOLIDWORKS API Help")),".")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,o.kt)("em",{parentName:"li"},"Solidworks Component"),"."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"swComponent")," "),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Type"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"SldWorks.Component2"),"."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2_members.html"},"online SOLIDWORKS API Help")),".")),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"These all are our global variables."))),(0,o.kt)("p",null,"They are ",(0,o.kt)("strong",{parentName:"p"},"SOLIDWORKS API Objects"),"."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Program to Show Selected Component\nSub main()\n\nEnd Sub\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we create ",(0,o.kt)("em",{parentName:"li"},"Program to Show Selected Component"),"."),(0,o.kt)("li",{parentName:"ul"},"This is a ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"Sub"))," procedure which has name of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"main")),". "),(0,o.kt)("li",{parentName:"ul"},"This procedure hold all the ",(0,o.kt)("em",{parentName:"li"},"statements (instructions)")," we give to computer."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Reference"),": Detailed information \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"/vba/vba-sub-and-function-procedure/"},"VBA Sub and Function Procedures"))," article of this website.")),(0,o.kt)(r.Z,{mdxType:"AdComponent"}),(0,o.kt)("h3",{id:"initialize-global-variables"},"Initialize Global Variables"),(0,o.kt)("p",null,"In this section, we initialize global variables."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Application variable to current application\nSet swApp = Application.SldWorks\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we set ",(0,o.kt)("em",{parentName:"li"},"value")," of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swApp"))," variable."),(0,o.kt)("li",{parentName:"ul"},"This ",(0,o.kt)("em",{parentName:"li"},"value")," is currently opened Solidworks application.")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks document variable to currently opened document\nSet swDoc = swApp.ActiveDoc\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we set ",(0,o.kt)("em",{parentName:"li"},"value")," of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,o.kt)("li",{parentName:"ul"},"This ",(0,o.kt)("em",{parentName:"li"},"value")," is currently ",(0,o.kt)("em",{parentName:"li"},"opened part document"),".")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Check if Solidworks document is opened or not\nIf swDoc Is Nothing Then\n  MsgBox ("Solidworks document is not opened.")\n  Exit Sub\nEnd If\n')),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully set the value of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,o.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Condition"),": ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swDoc Is Nothing"))),(0,o.kt)("li",{parentName:"ul"},"When this condition is ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,o.kt)("ul",{parentName:"li"},(0,o.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Message"),": ",(0,o.kt)("em",{parentName:"li"},"SOLIDWORKS document is not opened.")),(0,o.kt)("li",{parentName:"ul"},"Then we ",(0,o.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Assembly document\nSet swAssembly = swDoc\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we set ",(0,o.kt)("em",{parentName:"li"},"value")," of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable."),(0,o.kt)("li",{parentName:"ul"},"This ",(0,o.kt)("em",{parentName:"li"},"value")," is ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable.")),(0,o.kt)("h3",{id:"show-component"},"Show Component"),(0,o.kt)("p",null,"In this section, we perform ",(0,o.kt)("em",{parentName:"p"},"Show Component")," action."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for List of elements\nDim vArray As Variant\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,o.kt)("em",{parentName:"li"},"List of elements"),"."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"vArray")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Type"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"Variant"))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Get Components list in opened assembly\nvArray = swAssembly.GetComponents(True)\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we set the value of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"vArray"))," variable. "),(0,o.kt)("li",{parentName:"ul"},"We set value by ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"GetComponents"))," method of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable.")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for component\nDim component As Variant\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we create ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"component"))," variable for looping."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"component")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Type"),": ",(0,o.kt)("inlineCode",{parentName:"li"},"Variant"))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Loop Components List\nFor Each component In vArray\n  \nNext\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we create a ",(0,o.kt)("inlineCode",{parentName:"li"},"For Each")," loop."),(0,o.kt)("li",{parentName:"ul"},"In this loop, ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"component"))," variable loops every item in ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"vArray")),".")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Component variable\nSet swComponent = component\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we set ",(0,o.kt)("em",{parentName:"li"},"value")," of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swComponent"))," variable."),(0,o.kt)("li",{parentName:"ul"},"This ",(0,o.kt)("em",{parentName:"li"},"value")," is current value of array ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"vArray")),"."),(0,o.kt)("li",{parentName:"ul"},"Current value is represented by ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"component"))," variable.")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' If current component is hidden\nIf swComponent.IsHidden(False) Then\n  \nEnd If\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above code block, we check ",(0,o.kt)("em",{parentName:"li"},"if current component is hidden"),"."),(0,o.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Condition"),": ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swComponent.IsHidden(False)")))),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Select the component\nswComponent.Select False\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we ",(0,o.kt)("em",{parentName:"li"},"select the component"),"."),(0,o.kt)("li",{parentName:"ul"},"We use ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"Select"))," method of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swComponent"))," variable.")),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Show selected component\nswDoc.ShowComponent2\n")),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},"In above line, we Show selected component."),(0,o.kt)("li",{parentName:"ul"},"We use ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"ShowComponent2"))," method of ",(0,o.kt)("strong",{parentName:"li"},(0,o.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,o.kt)("li",{parentName:"ul"},"This method return nothing.")),(0,o.kt)("p",null,"Now we run the macro and after running macro we show selected component as shown in below image."),(0,o.kt)("p",null,(0,o.kt)("a",{target:"_blank",href:n(25801).Z},(0,o.kt)("img",{alt:"assembly-show-component",src:n(86565).Z,width:"1366",height:"728"}))),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},"This is it !!!")),(0,o.kt)("p",null,(0,o.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")," \ud83d\ude0a"),(0,o.kt)("p",null,"If you found anything to ",(0,o.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,o.kt)("em",{parentName:"p"},"e-mail")," \ud83d\udce7."),(0,o.kt)("p",null,"Hope this post helps you to ",(0,o.kt)("strong",{parentName:"p"},"Show Component")," with SOLIDWORKS VBA Macros."),(0,o.kt)("p",null,"For more such tutorials on ",(0,o.kt)("strong",{parentName:"p"},"SOLIDWORKS VBA Macro"),", do come to this website after sometime."),(0,o.kt)("p",null,(0,o.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")," \ud83d\ude4f\ud83c\udffb"),(0,o.kt)("p",null,(0,o.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,o.kt)("p",null,(0,o.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}u.isMDXComponent=!0},74753:(e,t,n)=>{n.d(t,{Z:()=>o});var a=n(67294);class o extends a.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return a.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}},25801:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/files/final-result-gif-6d7d1ac9cd33746cf4134bda6ba87348.gif"},90306:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/files/prerequisite-e92c591b038f03980a00214a560f21fc.png"},86565:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/images/final-result-gif-6d7d1ac9cd33746cf4134bda6ba87348.gif"},91013:(e,t,n)=>{n.d(t,{Z:()=>a});const a=n.p+"assets/images/prerequisite-e92c591b038f03980a00214a560f21fc.png"}}]);