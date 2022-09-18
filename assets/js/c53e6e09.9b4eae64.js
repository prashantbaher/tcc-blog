"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3843],{3905:function(e,t,n){n.d(t,{Zo:function(){return p},kt:function(){return c}});var a=n(67294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,a,o=function(e,t){if(null==e)return{};var n,a,o={},r=Object.keys(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var i=a.createContext({}),m=function(e){var t=a.useContext(i),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},p=function(e){var t=m(e.components);return a.createElement(i.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},k=a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,r=e.originalType,i=e.parentName,p=s(e,["components","mdxType","originalType","parentName"]),k=m(n),c=o,d=k["".concat(i,".").concat(c)]||k[c]||u[c]||r;return n?a.createElement(d,l(l({ref:t},p),{},{components:n})):a.createElement(d,l({ref:t},p))}));function c(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var r=n.length,l=new Array(r);l[0]=k;var s={};for(var i in t)hasOwnProperty.call(t,i)&&(s[i]=t[i]);s.originalType=e,s.mdxType="string"==typeof e?e:o,l[1]=s;for(var m=2;m<r;m++)l[m]=n[m];return a.createElement.apply(null,l)}return a.createElement.apply(null,n)}k.displayName="MDXCreateElement"},88499:function(e,t,n){n.r(t),n.d(t,{assets:function(){return u},contentTitle:function(){return m},default:function(){return d},frontMatter:function(){return i},metadata:function(){return p},toc:function(){return k}});var a=n(87462),o=n(63366),r=(n(67294),n(3905)),l=n(74753),s=["components"],i={categories:"Solidworks-macro",title:"Solidworks VBA Macro - Show Component",permalink:"/solidworks-vba-macros/assembly-show-component/",tags:["Solidworks Macro"],id:"assembly-show-component"},m=void 0,p={unversionedId:"assembly-show-component",id:"assembly-show-component",title:"Solidworks VBA Macro - Show Component",description:"",source:"@site/docs/solidworks-macros/025.5-assembly-show-component.md",sourceDirName:".",slug:"/assembly-show-component",permalink:"/solidworks-macros/assembly-show-component",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks VBA Macro - Show Component",permalink:"/solidworks-vba-macros/assembly-show-component/",tags:["Solidworks Macro"],id:"assembly-show-component"},sidebar:"tutorialSidebar",previous:{title:"Solidworks VBA Macro - Hide Component",permalink:"/solidworks-macros/assembly-hide-component"},next:{title:"Solidworks VBA Macro - Suppress Component",permalink:"/solidworks-macros/assembly-suppress-component"}},u={},k=[{value:"Objective",id:"objective",level:2},{value:"Results We Can Get",id:"results-we-can-get",level:2},{value:"Macro Video",id:"macro-video",level:2},{value:"VBA Macro",id:"vba-macro",level:2},{value:"Prerequisite",id:"prerequisite",level:2},{value:"Steps To Follow",id:"steps-to-follow",level:2},{value:"Create Global Variables",id:"create-global-variables",level:3},{value:"Initialize Global Variables",id:"initialize-global-variables",level:3},{value:"Show Component",id:"show-component",level:3}],c={toc:k};function d(e){var t=e.components,i=(0,o.Z)(e,s);return(0,r.kt)("wrapper",(0,a.Z)({},c,i,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("h2",{id:"objective"},"Objective"),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,'In this article, we understand "how to" ',(0,r.kt)("strong",{parentName:"p"},"Show Component")," in ",(0,r.kt)("strong",{parentName:"p"},"Assembly document")," from VBA macro."),(0,r.kt)("p",null,"This is most updated method of ",(0,r.kt)("strong",{parentName:"p"},"Show Component")," in an assembly document."),(0,r.kt)("h2",{id:"results-we-can-get"},"Results We Can Get"),(0,r.kt)("p",null,"Below image shows the result we get."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:n(25801).Z},(0,r.kt)("img",{alt:"assembly-show-component",src:n(86565).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,"We ",(0,r.kt)("strong",{parentName:"p"},"Show Component")," in simple manners."),(0,r.kt)("p",null,"There are no extra steps required."),(0,r.kt)("div",{className:"admonition admonition-caution alert alert--warning"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"16",height:"16",viewBox:"0 0 16 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M8.893 1.5c-.183-.31-.52-.5-.887-.5s-.703.19-.886.5L.138 13.499a.98.98 0 0 0 0 1.001c.193.31.53.501.886.501h13.964c.367 0 .704-.19.877-.5a1.03 1.03 0 0 0 .01-1.002L8.893 1.5zm.133 11.497H6.987v-2.003h2.039v2.003zm0-3.004H6.987V5.987h2.039v4.006z"}))),"caution")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"To get the correct result, please follow the steps correctly."))),(0,r.kt)("h2",{id:"macro-video"},"Macro Video"),(0,r.kt)("p",null,"Below \ud83c\udfac video shows how to ",(0,r.kt)("strong",{parentName:"p"},"Show Component")," from ",(0,r.kt)("em",{parentName:"p"},"SOLIDWORKS VBA Macros"),"."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/3YCXV0gpN3U",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Above video is just for visualization and there is no explanation."))," "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"I have explained every line in this article."))),(0,r.kt)("div",{className:"admonition admonition-caution alert alert--warning"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"16",height:"16",viewBox:"0 0 16 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M8.893 1.5c-.183-.31-.52-.5-.887-.5s-.703.19-.886.5L.138 13.499a.98.98 0 0 0 0 1.001c.193.31.53.501.886.501h13.964c.367 0 .704-.19.877-.5a1.03 1.03 0 0 0 .01-1.002L8.893 1.5zm.133 11.497H6.987v-2.003h2.039v2.003zm0-3.004H6.987V5.987h2.039v4.006z"}))),"caution")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"It is advisable to watch video, since it helps you to better understand the process."))),(0,r.kt)("h2",{id:"vba-macro"},"VBA Macro"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"VBA macro"))," for ",(0,r.kt)("em",{parentName:"p"},"Show Component"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Variable for Solidworks Application\nDim swApp As SldWorks.SldWorks\n\n' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n\n' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n\n' Program to Show Selected Component\nSub main()\n  \n  ' Set Solidworks Application variable to current application\n  Set swApp = Application.SldWorks\n  \n  ' Set Solidworks document variable to currently opened document\n  Set swDoc = swApp.ActiveDoc\n  \n  ' Check if Solidworks document is opened or not\n  If swDoc Is Nothing Then\n    MsgBox \"Solidworks document is not opened.\"\n    Exit Sub\n  End If\n  \n  ' Set Solidworks Assembly document\n  Set swAssembly = swDoc\n  \n  ' Variable for List of elements\n  Dim vArray As Variant\n  \n  ' Get Components list in opened assembly\n  vArray = swAssembly.GetComponents(True)\n  \n  ' Variable for component\n  Dim component As Variant\n  \n  ' Loop Components List\n  For Each component In vArray\n  \n    ' Set Solidworks Component variable\n    Set swComponent = component\n    \n    ' If current component is hidden\n    If swComponent.IsHidden(False) Then\n      \n      ' Select the component\n      swComponent.Select False\n      \n      ' Show selected component\n      swDoc.ShowComponent2\n    End If\n    \n  Next component\n  \nEnd Sub\n")),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("h2",{id:"prerequisite"},"Prerequisite"),(0,r.kt)("p",null,"There are some ",(0,r.kt)("em",{parentName:"p"},"prerequisites")," for this article."),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"Knowledge of ",(0,r.kt)("strong",{parentName:"li"},"VBA programming language")," is \u2757",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("em",{parentName:"strong"},"required")),"."),(0,r.kt)("li",{parentName:"ul"},"We use existing parts in Assembly document."),(0,r.kt)("li",{parentName:"ul"},"Both components are fully constraint as shown in below image."),(0,r.kt)("li",{parentName:"ul"},"We select the part which we want to Show.")),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:n(90306).Z},(0,r.kt)("img",{alt:"prerequisite",src:n(91013).Z,width:"1366",height:"709"}))),(0,r.kt)("div",{className:"admonition admonition-info alert alert--info"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"14",height:"16",viewBox:"0 0 14 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M7 2.3c3.14 0 5.7 2.56 5.7 5.7s-2.56 5.7-5.7 5.7A5.71 5.71 0 0 1 1.3 8c0-3.14 2.56-5.7 5.7-5.7zM7 1C3.14 1 0 4.14 0 8s3.14 7 7 7 7-3.14 7-7-3.14-7-7-7zm1 3H6v5h2V4zm0 6H6v2h2v-2z"}))),"info")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"We will apply checks in this article, so the code we write, should be ",(0,r.kt)("strong",{parentName:"p"},"error free")," mostly."))),(0,r.kt)("h2",{id:"steps-to-follow"},"Steps To Follow"),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,"This ",(0,r.kt)("strong",{parentName:"p"},"VBA macro")," can be divided into following sections:"),(0,r.kt)("ol",null,(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Create Global Variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Initialize Global Variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Show Component"))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Every section with each line is explained below."))),(0,r.kt)("div",{className:"admonition admonition-tip alert alert--success"},(0,r.kt)("div",{parentName:"div",className:"admonition-heading"},(0,r.kt)("h5",{parentName:"div"},(0,r.kt)("span",{parentName:"h5",className:"admonition-icon"},(0,r.kt)("svg",{parentName:"span",xmlns:"http://www.w3.org/2000/svg",width:"12",height:"16",viewBox:"0 0 12 16"},(0,r.kt)("path",{parentName:"svg",fillRule:"evenodd",d:"M6.5 0C3.48 0 1 2.19 1 5c0 .92.55 2.25 1 3 1.34 2.25 1.78 2.78 2 4v1h5v-1c.22-1.22.66-1.75 2-4 .45-.75 1-2.08 1-3 0-2.81-2.48-5-5.5-5zm3.64 7.48c-.25.44-.47.8-.67 1.11-.86 1.41-1.25 2.06-1.45 3.23-.02.05-.02.11-.02.17H5c0-.06 0-.13-.02-.17-.2-1.17-.59-1.83-1.45-3.23-.2-.31-.42-.67-.67-1.11C2.44 6.78 2 5.65 2 5c0-2.2 2.02-4 4.5-4 1.22 0 2.36.42 3.22 1.19C10.55 2.94 11 3.94 11 5c0 .66-.44 1.78-.86 2.48zM4 14h5c-.23 1.14-1.3 2-2.5 2s-2.27-.86-2.5-2z"}))),"tip")),(0,r.kt)("div",{parentName:"div",className:"admonition-content"},(0,r.kt)("p",{parentName:"div"},"I also give some ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"links (see icon \ud83d\ude80)"))," so that you can go through them if there are anything I explained in previous articles."))),(0,r.kt)("h3",{id:"create-global-variables"},"Create Global Variables"),(0,r.kt)("p",null,"In this section, we create global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": Above line forces us to define every variable we are going to use. "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"SOLIDWORKS Macros - Open new Part document"))," article.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks application"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swApp")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.SldWorks")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks document"),". "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swDoc")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.ModelDoc2")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Assembly"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swAssembly")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.AssemblyDoc")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Component\nDim swComponent As SldWorks.Component2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Component"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swComponent")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.Component2"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IComponent2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"These all are our global variables."))),(0,r.kt)("p",null,"They are ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS API Objects"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Program to Show Selected Component\nSub main()\n\nEnd Sub\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create ",(0,r.kt)("em",{parentName:"li"},"Program to Show Selected Component"),"."),(0,r.kt)("li",{parentName:"ul"},"This is a ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Sub"))," procedure which has name of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"main")),". "),(0,r.kt)("li",{parentName:"ul"},"This procedure hold all the ",(0,r.kt)("em",{parentName:"li"},"statements (instructions)")," we give to computer."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Detailed information \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-sub-and-function-procedure/"},"VBA Sub and Function Procedures"))," article of this website.")),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("h3",{id:"initialize-global-variables"},"Initialize Global Variables"),(0,r.kt)("p",null,"In this section, we initialize global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Application variable to current application\nSet swApp = Application.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swApp"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently opened Solidworks application.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks document variable to currently opened document\nSet swDoc = swApp.ActiveDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently ",(0,r.kt)("em",{parentName:"li"},"opened part document"),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Check if Solidworks document is opened or not\nIf swDoc Is Nothing Then\n  MsgBox ("Solidworks document is not opened.")\n  Exit Sub\nEnd If\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully set the value of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc Is Nothing"))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": ",(0,r.kt)("em",{parentName:"li"},"SOLIDWORKS document is not opened.")),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Assembly document\nSet swAssembly = swDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable.")),(0,r.kt)("h3",{id:"show-component"},"Show Component"),(0,r.kt)("p",null,"In this section, we perform ",(0,r.kt)("em",{parentName:"p"},"Show Component")," action."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for List of elements\nDim vArray As Variant\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"List of elements"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"vArray")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"Variant"))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Get Components list in opened assembly\nvArray = swAssembly.GetComponents(True)\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set the value of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"vArray"))," variable. "),(0,r.kt)("li",{parentName:"ul"},"We set value by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"GetComponents"))," method of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for component\nDim component As Variant\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"component"))," variable for looping."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"component")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"Variant"))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Loop Components List\nFor Each component In vArray\n  \nNext\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create a ",(0,r.kt)("inlineCode",{parentName:"li"},"For Each")," loop."),(0,r.kt)("li",{parentName:"ul"},"In this loop, ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"component"))," variable loops every item in ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"vArray")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Component variable\nSet swComponent = component\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is current value of array ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"vArray")),"."),(0,r.kt)("li",{parentName:"ul"},"Current value is represented by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"component"))," variable.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' If current component is hidden\nIf swComponent.IsHidden(False) Then\n  \nEnd If\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check ",(0,r.kt)("em",{parentName:"li"},"if current component is hidden"),"."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent.IsHidden(False)")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Select the component\nswComponent.Select False\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we ",(0,r.kt)("em",{parentName:"li"},"select the component"),"."),(0,r.kt)("li",{parentName:"ul"},"We use ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Select"))," method of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swComponent"))," variable.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Show selected component\nswDoc.ShowComponent2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we Show selected component."),(0,r.kt)("li",{parentName:"ul"},"We use ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"ShowComponent2"))," method of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This method return nothing.")),(0,r.kt)("p",null,"Now we run the macro and after running macro we show selected component as shown in below image."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:n(25801).Z},(0,r.kt)("img",{alt:"assembly-show-component",src:n(86565).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"This is it !!!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")," \ud83d\ude0a"),(0,r.kt)("p",null,"If you found anything to ",(0,r.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,r.kt)("em",{parentName:"p"},"e-mail")," \ud83d\udce7."),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("strong",{parentName:"p"},"Show Component")," with SOLIDWORKS VBA Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS VBA Macro"),", do come to this website after sometime."),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")," \ud83d\ude4f\ud83c\udffb"),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}d.isMDXComponent=!0},74753:function(e,t,n){n.d(t,{Z:function(){return r}});var a=n(94578),o=n(67294),r=function(e){function t(){return e.apply(this,arguments)||this}(0,a.Z)(t,e);var n=t.prototype;return n.componentDidMount=function(){var e;(e=document.createElement("script")).src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild),(window.adsbygoogle=window.adsbygoogle||[]).push({})},n.render=function(){return o.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})},t}(o.Component)},25801:function(e,t,n){t.Z=n.p+"assets/files/final-result-gif-6d7d1ac9cd33746cf4134bda6ba87348.gif"},90306:function(e,t,n){t.Z=n.p+"assets/files/prerequisite-e92c591b038f03980a00214a560f21fc.png"},86565:function(e,t,n){t.Z=n.p+"assets/images/final-result-gif-6d7d1ac9cd33746cf4134bda6ba87348.gif"},91013:function(e,t,n){t.Z=n.p+"assets/images/prerequisite-e92c591b038f03980a00214a560f21fc.png"}}]);