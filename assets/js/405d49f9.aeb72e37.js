"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[6221],{3905:(e,n,t)=>{t.d(n,{Zo:()=>m,kt:()=>c});var a=t(67294);function r(e,n,t){return n in e?Object.defineProperty(e,n,{value:t,enumerable:!0,configurable:!0,writable:!0}):e[n]=t,e}function o(e,n){var t=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);n&&(a=a.filter((function(n){return Object.getOwnPropertyDescriptor(e,n).enumerable}))),t.push.apply(t,a)}return t}function s(e){for(var n=1;n<arguments.length;n++){var t=null!=arguments[n]?arguments[n]:{};n%2?o(Object(t),!0).forEach((function(n){r(e,n,t[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(t)):o(Object(t)).forEach((function(n){Object.defineProperty(e,n,Object.getOwnPropertyDescriptor(t,n))}))}return e}function l(e,n){if(null==e)return{};var t,a,r=function(e,n){if(null==e)return{};var t,a,r={},o=Object.keys(e);for(a=0;a<o.length;a++)t=o[a],n.indexOf(t)>=0||(r[t]=e[t]);return r}(e,n);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(a=0;a<o.length;a++)t=o[a],n.indexOf(t)>=0||Object.prototype.propertyIsEnumerable.call(e,t)&&(r[t]=e[t])}return r}var i=a.createContext({}),p=function(e){var n=a.useContext(i),t=n;return e&&(t="function"==typeof e?e(n):s(s({},n),e)),t},m=function(e){var n=p(e.components);return a.createElement(i.Provider,{value:n},e.children)},u={inlineCode:"code",wrapper:function(e){var n=e.children;return a.createElement(a.Fragment,{},n)}},k=a.forwardRef((function(e,n){var t=e.components,r=e.mdxType,o=e.originalType,i=e.parentName,m=l(e,["components","mdxType","originalType","parentName"]),k=p(t),c=r,d=k["".concat(i,".").concat(c)]||k[c]||u[c]||o;return t?a.createElement(d,s(s({ref:n},m),{},{components:t})):a.createElement(d,s({ref:n},m))}));function c(e,n){var t=arguments,r=n&&n.mdxType;if("string"==typeof e||r){var o=t.length,s=new Array(o);s[0]=k;var l={};for(var i in n)hasOwnProperty.call(n,i)&&(l[i]=n[i]);l.originalType=e,l.mdxType="string"==typeof e?e:r,s[1]=l;for(var p=2;p<o;p++)s[p]=t[p];return a.createElement.apply(null,s)}return a.createElement.apply(null,t)}k.displayName="MDXCreateElement"},538:(e,n,t)=>{t.r(n),t.d(n,{assets:()=>p,contentTitle:()=>l,default:()=>k,frontMatter:()=>s,metadata:()=>i,toc:()=>m});var a=t(87462),r=(t(67294),t(3905)),o=t(74753);const s={categories:"Solidworks-macro",title:"Solidworks VBA Macro - UnSuppress Component",permalink:"/solidworks-vba-macros/assembly-unsuppress-component/",tags:["Solidworks Macro"],id:"assembly-unsuppress-component"},l=void 0,i={unversionedId:"assembly-unsuppress-component",id:"assembly-unsuppress-component",title:"Solidworks VBA Macro - UnSuppress Component",description:"",source:"@site/docs/solidworks-macros/025.7-assembly-unsuppress-component.md",sourceDirName:".",slug:"/assembly-unsuppress-component",permalink:"/solidworks-macros/assembly-unsuppress-component",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks VBA Macro - UnSuppress Component",permalink:"/solidworks-vba-macros/assembly-unsuppress-component/",tags:["Solidworks Macro"],id:"assembly-unsuppress-component"},sidebar:"tutorialSidebar",previous:{title:"Solidworks VBA Macro - Suppress Component",permalink:"/solidworks-macros/assembly-suppress-component"},next:{title:"Solidworks Macro - Convert to Construction Sketch",permalink:"/solidworks-macros/convert-to-construction-sketch-segment"}},p={},m=[{value:"Objective",id:"objective",level:2},{value:"Results We Can Get",id:"results-we-can-get",level:2},{value:"Macro Video",id:"macro-video",level:2},{value:"VBA Macro",id:"vba-macro",level:2},{value:"Prerequisite",id:"prerequisite",level:2},{value:"Steps To Follow",id:"steps-to-follow",level:2},{value:"Create Global Variables",id:"create-global-variables",level:3},{value:"Initialize Global Variables",id:"initialize-global-variables",level:3},{value:"UnSuppress Component",id:"unsuppress-component",level:3}],u={toc:m};function k(e){let{components:n,...s}=e;return(0,r.kt)("wrapper",(0,a.Z)({},u,s,{components:n,mdxType:"MDXLayout"}),(0,r.kt)("h2",{id:"objective"},"Objective"),(0,r.kt)(o.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,'In this article, we understand "how to" ',(0,r.kt)("strong",{parentName:"p"},"UnSuppress Component")," in ",(0,r.kt)("strong",{parentName:"p"},"Assembly document")," from VBA macro."),(0,r.kt)("p",null,"This is most updated method of ",(0,r.kt)("strong",{parentName:"p"},"UnSuppress Component")," in an assembly document."),(0,r.kt)("h2",{id:"results-we-can-get"},"Results We Can Get"),(0,r.kt)("p",null,"Below image shows the result we get."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:t(19370).Z},(0,r.kt)("img",{alt:"assembly-unsuppress-component",src:t(75433).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,"We ",(0,r.kt)("strong",{parentName:"p"},"UnSuppress Component")," in simple manners."),(0,r.kt)("p",null,"Macro will work automatically, so no extra steps required."),(0,r.kt)("admonition",{type:"caution"},(0,r.kt)("p",{parentName:"admonition"},"To get the correct result, please follow the steps correctly.")),(0,r.kt)("h2",{id:"macro-video"},"Macro Video"),(0,r.kt)("p",null,"Below \ud83c\udfac video shows how to ",(0,r.kt)("strong",{parentName:"p"},"UnSuppress Component")," from ",(0,r.kt)("em",{parentName:"p"},"SOLIDWORKS VBA Macros"),"."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/Gem2n32rwf4",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Above video is just for visualization and there is no explanation."))," "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"I have explained every line in this article."))),(0,r.kt)("admonition",{type:"caution"},(0,r.kt)("p",{parentName:"admonition"},"It is advisable to watch video, since it helps you to better understand the process.")),(0,r.kt)("h2",{id:"vba-macro"},"VBA Macro"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"VBA macro"))," for ",(0,r.kt)("em",{parentName:"p"},"Show Component"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers {48-57}",showlinenumbers:!0,showLineNumbers:!0,"{48-57}":!0},"Option Explicit\n\n' Variable for Solidworks Application\nDim swApp As SldWorks.SldWorks\n\n' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n\n' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n\n' Program to UnSuppress Selected Component\nSub main()\n  \n  ' Set Solidworks Application variable to current application\n  Set swApp = Application.SldWorks\n  \n  ' Set Solidworks document variable to currently opened document\n  Set swDoc = swApp.ActiveDoc\n  \n  ' Check if Solidworks document is opened or not\n  If swDoc Is Nothing Then\n    MsgBox \"Solidworks document is not opened.\"\n    Exit Sub\n  End If\n  \n  ' Set Solidworks Assembly document\n  Set swAssembly = swDoc\n  \n  ' Variable for List of elements\n  Dim vArray As Variant\n  \n  ' Get Components list in opened assembly\n  vArray = swAssembly.GetComponents(False)\n  \n  ' Variable for component\n  Dim component As Variant\n  \n  ' Loop Components List\n  For Each component In vArray\n  \n    ' Set Solidworks Component variable\n    Set swComponent = component\n    \n    ' If current component is Suppress\n    If swComponent.IsSuppressed Then\n\n      ' Select the component\n      swComponent.Select False\n      \n      ' UnSuppress selected component\n      swDoc.EditUnsuppress2\n      \n    End If\n    \n  Next component\n  \nEnd Sub\n")),(0,r.kt)(o.Z,{mdxType:"AdComponent"}),(0,r.kt)("h2",{id:"prerequisite"},"Prerequisite"),(0,r.kt)("p",null,"There are some ",(0,r.kt)("em",{parentName:"p"},"prerequisites")," for this article."),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"Knowledge of ",(0,r.kt)("strong",{parentName:"li"},"VBA programming language")," is \u2757",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("em",{parentName:"strong"},"required")),"."),(0,r.kt)("li",{parentName:"ul"},"We use existing parts in Assembly document."),(0,r.kt)("li",{parentName:"ul"},"All components are fully constraint as shown in below image."),(0,r.kt)("li",{parentName:"ul"},"We select the part which we want to Unsuppress.")),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:t(81725).Z},(0,r.kt)("img",{alt:"prerequisite",src:t(15156).Z,width:"1333",height:"578"}))),(0,r.kt)("admonition",{type:"info"},(0,r.kt)("p",{parentName:"admonition"},"We will apply checks in this article, so the code we write, should be ",(0,r.kt)("strong",{parentName:"p"},"error free")," mostly.")),(0,r.kt)("h2",{id:"steps-to-follow"},"Steps To Follow"),(0,r.kt)("p",null,"This ",(0,r.kt)("strong",{parentName:"p"},"VBA macro")," can be divided into following sections:"),(0,r.kt)("ol",null,(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Create Global Variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Initialize Global Variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"UnSuppress Component"))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Every section with each line is explained below."))),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"I also give some ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"links (see icon \ud83d\ude80)"))," so that you can go through them if there are anything I explained in previous articles.")),(0,r.kt)("h3",{id:"create-global-variables"},"Create Global Variables"),(0,r.kt)("p",null,"In this section, we create global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": Above line forces us to define every variable we are going to use. "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"SOLIDWORKS Macros - Open new Part document"))," article.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks application"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swApp")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.SldWorks")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks document"),". "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swDoc")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.ModelDoc2")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Assembly"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swAssembly")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.AssemblyDoc")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Component"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swComponent")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.Component2"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"These all are our global variables."))),(0,r.kt)("p",null,"They are ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS API Objects"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Program to UnSuppress Selected Component\nSub main()\n\nEnd Sub\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create ",(0,r.kt)("em",{parentName:"li"},"Program to UnSuppress Selected Component"),"."),(0,r.kt)("li",{parentName:"ul"},"This is a ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Sub"))," procedure which has name of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"main")),". "),(0,r.kt)("li",{parentName:"ul"},"This procedure hold all the ",(0,r.kt)("em",{parentName:"li"},"statements (instructions)")," we give to computer."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Detailed information \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-sub-and-function-procedure/"},"VBA Sub and Function Procedures"))," article of this website.")),(0,r.kt)(o.Z,{mdxType:"AdComponent"}),(0,r.kt)("h3",{id:"initialize-global-variables"},"Initialize Global Variables"),(0,r.kt)("p",null,"In this section, we initialize global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Application variable to current application\nSet swApp = Application.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swApp"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently opened Solidworks application.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks document variable to currently opened document\nSet swDoc = swApp.ActiveDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently ",(0,r.kt)("em",{parentName:"li"},"opened part document"),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Check if Solidworks document is opened or not\nIf swDoc Is Nothing Then\n  MsgBox ("Solidworks document is not opened.")\n  Exit Sub\nEnd If\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully set the value of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc Is Nothing"))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": ",(0,r.kt)("em",{parentName:"li"},"SOLIDWORKS document is not opened.")),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Assembly document\nSet swAssembly = swDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable.")),(0,r.kt)("h3",{id:"unsuppress-component"},"UnSuppress Component"),(0,r.kt)("p",null,"In this section, we perform ",(0,r.kt)("em",{parentName:"p"},"UnSuppress Component")," action."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for List of elements\nDim vArray As Variant\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"List of elements"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"vArray")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"Variant"))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Get Components list in opened assembly\nvArray = swAssembly.GetComponents(False)\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set the value of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"vArray"))," variable. "),(0,r.kt)("li",{parentName:"ul"},"We set value by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"GetComponents"))," method of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable."),(0,r.kt)("li",{parentName:"ul"},"By passing ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"False"))," to ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"GetComponents"))," method, we get all components from Feature Tree.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for component\nDim component As Variant\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"component"))," variable for looping."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"component")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"Variant"))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Loop Components List\nFor Each component In vArray\n  \nNext\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create a ",(0,r.kt)("inlineCode",{parentName:"li"},"For Each")," loop."),(0,r.kt)("li",{parentName:"ul"},"In this loop, ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"component"))," variable loops every item in ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"vArray")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Component variable\nSet swComponent = component\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is current value of array ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"vArray")),"."),(0,r.kt)("li",{parentName:"ul"},"Current value is represented by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"component"))," variable.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' If current component is Suppress\nIf swComponent.IsSuppressed Then\n  \nEnd If\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check ",(0,r.kt)("em",{parentName:"li"},"if current component is Suppress"),"."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent.IsSuppressed")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Select the component\nswComponent.Select False\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we ",(0,r.kt)("em",{parentName:"li"},"select the component"),"."),(0,r.kt)("li",{parentName:"ul"},"We use ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Select"))," method of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent"))," variable.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' UnSuppress selected component\nswDoc.EditUnsuppress2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we Show selected component."),(0,r.kt)("li",{parentName:"ul"},"We use ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"EditUnsuppress2"))," method of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This method return nothing.")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Return Value")," : This ",(0,r.kt)("inlineCode",{parentName:"li"},"EditUnsuppress2")," method return ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True"))," or ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"False")),".")),(0,r.kt)("p",null,"Now we run the macro and after running macro we show selected component as shown in below image."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:t(19370).Z},(0,r.kt)("img",{alt:"assembly-unsuppress-component",src:t(75433).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"This is it !!!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")," \ud83d\ude0a"),(0,r.kt)("p",null,"If you found anything to ",(0,r.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,r.kt)("em",{parentName:"p"},"e-mail")," \ud83d\udce7."),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("strong",{parentName:"p"},"UnSuppress Component")," with SOLIDWORKS VBA Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS VBA Macro"),", do come to this website after sometime."),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")," \ud83d\ude4f\ud83c\udffb"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}k.isMDXComponent=!0},74753:(e,n,t)=>{t.d(n,{Z:()=>r});var a=t(67294);class r extends a.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return a.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}},19370:(e,n,t)=>{t.d(n,{Z:()=>a});const a=t.p+"assets/files/final-result-gif-ab57a40e0f7ac6e1bec717bec3e7f107.gif"},81725:(e,n,t)=>{t.d(n,{Z:()=>a});const a=t.p+"assets/files/prerequisite-cf6f16b9a0ea0aba3440e0e4f3f13d72.png"},75433:(e,n,t)=>{t.d(n,{Z:()=>a});const a=t.p+"assets/images/final-result-gif-ab57a40e0f7ac6e1bec717bec3e7f107.gif"},15156:(e,n,t)=>{t.d(n,{Z:()=>a});const a=t.p+"assets/images/prerequisite-cf6f16b9a0ea0aba3440e0e4f3f13d72.png"}}]);