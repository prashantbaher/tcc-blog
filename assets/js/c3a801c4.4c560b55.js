"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[7731],{3905:(e,t,a)=>{a.d(t,{Zo:()=>p,kt:()=>c});var n=a(67294);function r(e,t,a){return t in e?Object.defineProperty(e,t,{value:a,enumerable:!0,configurable:!0,writable:!0}):e[t]=a,e}function l(e,t){var a=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),a.push.apply(a,n)}return a}function s(e){for(var t=1;t<arguments.length;t++){var a=null!=arguments[t]?arguments[t]:{};t%2?l(Object(a),!0).forEach((function(t){r(e,t,a[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(a)):l(Object(a)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(a,t))}))}return e}function o(e,t){if(null==e)return{};var a,n,r=function(e,t){if(null==e)return{};var a,n,r={},l=Object.keys(e);for(n=0;n<l.length;n++)a=l[n],t.indexOf(a)>=0||(r[a]=e[a]);return r}(e,t);if(Object.getOwnPropertySymbols){var l=Object.getOwnPropertySymbols(e);for(n=0;n<l.length;n++)a=l[n],t.indexOf(a)>=0||Object.prototype.propertyIsEnumerable.call(e,a)&&(r[a]=e[a])}return r}var i=n.createContext({}),m=function(e){var t=n.useContext(i),a=t;return e&&(a="function"==typeof e?e(t):s(s({},t),e)),a},p=function(e){var t=m(e.components);return n.createElement(i.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},k=n.forwardRef((function(e,t){var a=e.components,r=e.mdxType,l=e.originalType,i=e.parentName,p=o(e,["components","mdxType","originalType","parentName"]),k=m(a),c=r,d=k["".concat(i,".").concat(c)]||k[c]||u[c]||l;return a?n.createElement(d,s(s({ref:t},p),{},{components:a})):n.createElement(d,s({ref:t},p))}));function c(e,t){var a=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var l=a.length,s=new Array(l);s[0]=k;var o={};for(var i in t)hasOwnProperty.call(t,i)&&(o[i]=t[i]);o.originalType=e,o.mdxType="string"==typeof e?e:r,s[1]=o;for(var m=2;m<l;m++)s[m]=a[m];return n.createElement.apply(null,s)}return n.createElement.apply(null,a)}k.displayName="MDXCreateElement"},82088:(e,t,a)=>{a.r(t),a.d(t,{assets:()=>m,contentTitle:()=>o,default:()=>k,frontMatter:()=>s,metadata:()=>i,toc:()=>p});var n=a(87462),r=(a(67294),a(3905)),l=a(74753);const s={categories:"Solidworks-macro",title:"Solidworks VBA Macro - Rename Mate",permalink:"/solidworks-vba-macros/assembly-rename-mate/",tags:["Solidworks Macro"],id:"assembly-rename-mate"},o=void 0,i={unversionedId:"assembly-rename-mate",id:"assembly-rename-mate",title:"Solidworks VBA Macro - Rename Mate",description:"",source:"@site/docs/solidworks-macros/024.1-assembly-rename-mate.md",sourceDirName:".",slug:"/assembly-rename-mate",permalink:"/solidworks-macros/assembly-rename-mate",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks VBA Macro - Rename Mate",permalink:"/solidworks-vba-macros/assembly-rename-mate/",tags:["Solidworks Macro"],id:"assembly-rename-mate"},sidebar:"tutorialSidebar",previous:{title:"Solidworks VBA Macro - Add Symmetric Mate",permalink:"/solidworks-macros/assembly-symmetric-mate"},next:{title:"Solidworks VBA Macro - Edit Distance Mate",permalink:"/solidworks-macros/assembly-edit-distance-mate"}},m={},p=[{value:"Objective",id:"objective",level:2},{value:"Results We Can Get",id:"results-we-can-get",level:2},{value:"Macro Video",id:"macro-video",level:2},{value:"VBA Macro",id:"vba-macro",level:2},{value:"Prerequisite",id:"prerequisite",level:2},{value:"Steps To Follow",id:"steps-to-follow",level:2},{value:"Create global variables",id:"create-global-variables",level:3},{value:"Initialize global variables",id:"initialize-global-variables",level:3},{value:"Add Rename Mate",id:"add-rename-mate",level:3}],u={toc:p};function k(e){let{components:t,...s}=e;return(0,r.kt)("wrapper",(0,n.Z)({},u,s,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("h2",{id:"objective"},"Objective"),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("p",null,'In this article, we understand "how to" ',(0,r.kt)("strong",{parentName:"p"},"Rename a Mate")," in ",(0,r.kt)("strong",{parentName:"p"},"Assembly document")," from VBA macro."),(0,r.kt)("p",null,"You can use this method to ",(0,r.kt)("strong",{parentName:"p"},"rename any feature"),"."),(0,r.kt)("p",null,"In my example, ",(0,r.kt)("em",{parentName:"p"},"I am renaming last added mate"),"."),(0,r.kt)("h2",{id:"results-we-can-get"},"Results We Can Get"),(0,r.kt)("p",null,"Below image shows the result we get."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:a(4907).Z},(0,r.kt)("img",{alt:"assembly-rename-mate",src:a(38930).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,"We add ",(0,r.kt)("strong",{parentName:"p"},"Rename Mate")," in following steps."),(0,r.kt)("ol",null,(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Ask a name from user.")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Rename Mate."))),(0,r.kt)("admonition",{type:"caution"},(0,r.kt)("p",{parentName:"admonition"},"To get the correct result, please follow the steps correctly.")),(0,r.kt)("h2",{id:"macro-video"},"Macro Video"),(0,r.kt)("p",null,"Below \ud83c\udfac video shows how to ",(0,r.kt)("strong",{parentName:"p"},"Rename Mate")," from ",(0,r.kt)("em",{parentName:"p"},"SOLIDWORKS VBA Macros"),"."),(0,r.kt)("iframe",{src:"https://www.youtube.com/embed/fWtj-YwXMFU",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Above video is just for visualization and there is no explanation."))," "),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"I have explained every line in this article."))),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"It is advisable to watch video, since it helps you to better understand the process.")),(0,r.kt)("h2",{id:"vba-macro"},"VBA Macro"),(0,r.kt)("p",null,"Below is the ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"VBA macro"))," for ",(0,r.kt)("em",{parentName:"p"},"Rename Mate"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Variable for Solidworks Application\nDim swApp As SldWorks.SldWorks\n\n' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n\n' Variable for Solidworks Mate Feature\nDim swMateFeature As SldWorks.Feature\n\n' Program to Rename Mate\nSub main()\n  \n  ' Set Solidworks Application variable to current application\n  Set swApp = Application.SldWorks\n  \n  ' Set Solidworks document variable to currently opened document\n  Set swDoc = swApp.ActiveDoc\n  \n  ' Check if Solidworks document is opened or not\n  If swDoc Is Nothing Then\n    MsgBox \"Solidworks document is not opened.\"\n    Exit Sub\n  End If\n  \n  ' Set Solidworks Assembly document\n  Set swAssembly = swDoc\n  \n  ' Get mate feature\n  Set swMateFeature = swDoc.Extension.GetLastFeatureAdded\n\n  ' Check if successfully Get mate\n  If swMateFeature Is Nothing Then\n    MsgBox \"Failed to Get Mate.\"\n    swDoc.ClearSelection2 True\n    Exit Sub\n  End If\n  \n  ' Select the mate\n  swMateFeature.Select True\n  \n  ' Start editing mate feature\n  swDoc.FeatEdit\n  \n  ' Variable for mate's new name\n  Dim newName As String\n  \n  ' Get mate's new name\n  newName = InputBox(\"New Name:\", \"Edit Mate\")\n  \n  ' This will handle empty value or cancel case\n  If Len(newName) = 0 Then\n    MsgBox \"Empty or no value. Please try again.\"\n    swDoc.ClearSelection2 True\n    Exit Sub\n  End If\n  \n  ' Update mate's new name\n  swMateFeature.Name = newName\n  \n  ' Clear all selection\n  swDoc.ClearSelection2 True\n  \n  ' Rebuild assembly\n  swDoc.ForceRebuild3 True\n  \nEnd Sub\n")),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("h2",{id:"prerequisite"},"Prerequisite"),(0,r.kt)("p",null,"There are some ",(0,r.kt)("em",{parentName:"p"},"prerequisites")," for this article."),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"Knowledge of ",(0,r.kt)("strong",{parentName:"p"},"VBA programming language")," is \u2757",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"required")),".")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"We use existing parts in Assembly document.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"One component is fully constraint and other component is Float as shown in below image."))),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:a(79456).Z},(0,r.kt)("img",{alt:"prerequisite",src:a(63726).Z,width:"1366",height:"728"}))),(0,r.kt)("admonition",{type:"note"},(0,r.kt)("p",{parentName:"admonition"},"We will apply checks in this article, so the code we write, should be ",(0,r.kt)("strong",{parentName:"p"},"mostly error free"),".")),(0,r.kt)("h2",{id:"steps-to-follow"},"Steps To Follow"),(0,r.kt)("p",null,"This ",(0,r.kt)("strong",{parentName:"p"},"VBA macro")," can be divided into following sections:"),(0,r.kt)("ol",null,(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Create global variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Initialize global variables")),(0,r.kt)("li",{parentName:"ol"},(0,r.kt)("em",{parentName:"li"},"Rename Mate"))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"Every section with each line is explained below."))),(0,r.kt)("admonition",{type:"tip"},(0,r.kt)("p",{parentName:"admonition"},"I also give some ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"links (see icon \ud83d\ude80)"))," so that you can go through them if there are anything I explained in previous articles.")),(0,r.kt)("h3",{id:"create-global-variables"},"Create global variables"),(0,r.kt)("p",null,"In this section, we create global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": Above line forces us to define every variable we are going to use. "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/solidworks-macros/open-new-document"},"SOLIDWORKS Macros - Open new Part document"))," article.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks application"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swApp")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.SldWorks")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.ISldWorks_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks document"),". "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swDoc")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.ModelDoc2")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Assembly\nDim swAssembly As SldWorks.AssemblyDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Assembly"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swAssembly")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.AssemblyDoc")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IAssemblyDoc_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for Solidworks Mate Feature\nDim swMateFeature As SldWorks.Feature\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable for ",(0,r.kt)("em",{parentName:"li"},"Solidworks Mate Feature"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"swMateFeature")," "),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("inlineCode",{parentName:"li"},"SldWorks.Feature"),"."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature_members.html"},"online SOLIDWORKS API Help")),".")),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"These all are our global variables."))),(0,r.kt)("p",null,"They are ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS API Objects"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Program to Rename Mate\nSub main()\n\nEnd Sub\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we create ",(0,r.kt)("em",{parentName:"li"},"main Program to Rename Mate in assembly"),"."),(0,r.kt)("li",{parentName:"ul"},"This is a ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Sub"))," procedure which has name of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"main")),". "),(0,r.kt)("li",{parentName:"ul"},"This procedure hold all the ",(0,r.kt)("em",{parentName:"li"},"statements (instructions)")," we give to computer."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Detailed information \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-sub-and-function-procedure/"},"VBA Sub and Function Procedures"))," article of this website.")),(0,r.kt)(l.Z,{mdxType:"AdComponent"}),(0,r.kt)("h3",{id:"initialize-global-variables"},"Initialize global variables"),(0,r.kt)("p",null,"In this section, we initialize global variables."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Application variable to current application\nSet swApp = Application.SldWorks\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swApp"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently opened Solidworks application.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks document variable to currently opened document\nSet swDoc = swApp.ActiveDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is currently ",(0,r.kt)("em",{parentName:"li"},"opened part document"),".")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Check if Solidworks document is opened or not\nIf swDoc Is Nothing Then\n  MsgBox ("Solidworks document is not opened.")\n  Exit Sub\nEnd If\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully set the value of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc Is Nothing"))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": ",(0,r.kt)("em",{parentName:"li"},"SOLIDWORKS document is not opened.")),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Set Solidworks Assembly document\nSet swAssembly = swDoc\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set ",(0,r.kt)("em",{parentName:"li"},"value")," of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swAssembly"))," variable."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("em",{parentName:"li"},"value")," is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," variable.")),(0,r.kt)("h3",{id:"add-rename-mate"},"Add Rename Mate"),(0,r.kt)("p",null,"In this section, we ",(0,r.kt)("em",{parentName:"p"},"Rename Mate"),"."),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Get mate feature\nSet swMateFeature = swDoc.Extension.GetLastFeatureAdded\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we set the value of variable ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swMateFeature"))," by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"GetLastFeatureAdded"))," method."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"GetLastFeatureAdded"))," method is part of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Extension"))," object."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Extension"))," object is then part of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swDoc"))," object."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"GetLastFeatureAdded"))," method gives us last added mate.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Check if successfully Get mate\nIf swMateFeature Is Nothing Then\n  MsgBox "Failed to Get Mate."\n  swDoc.ClearSelection2 True\n  Exit Sub\nEnd If\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check if we successfully get ",(0,r.kt)("strong",{parentName:"li"},"Last Mate")," or not."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swMateFeature Is Nothing"))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": ",(0,r.kt)("em",{parentName:"li"},"Failed to Get Mate"),"."),(0,r.kt)("li",{parentName:"ul"},"After that we clear the selection."),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Select the mate\nswMateFeature.Select True\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line we select the mate by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Select"))," method."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Select"))," method take either ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True"))," or ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"False")),".",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),": ",(0,r.kt)("em",{parentName:"li"},"Appends the feature to the current selection list.")),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"False")),": ",(0,r.kt)("em",{parentName:"li"},"Replaces the current selection list."))))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Start editing mate feature\nswDoc.FeatEdit\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line we start editing mate feature by ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"FeatEdit"))," method."),(0,r.kt)("li",{parentName:"ul"},"This ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"FeatEdit"))," method puts the current feature into edit mode. ")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Variable for mate's new name\nDim newName As String\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Purpose"),": In above line, we create a variable to store new name."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Variable Name"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"newName"))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Type"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"String")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Get mate\'s new name\nnewName = InputBox("New Name:", "Edit Mate")\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},"In above line of code we are doing ",(0,r.kt)("strong",{parentName:"p"},"2 steps")," in one line."),(0,r.kt)("p",{parentName:"li"},"Those 2 steps are explained below."),(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"Step 1")," - Getting ",(0,r.kt)("strong",{parentName:"p"},"New Name")," from user."),(0,r.kt)("p",{parentName:"li"},"Below image shows the message for ",(0,r.kt)("strong",{parentName:"p"},"New Name")," to the user."),(0,r.kt)("p",{parentName:"li"},(0,r.kt)("a",{target:"_blank",href:a(97533).Z},(0,r.kt)("img",{alt:"message-for-new-name",src:a(97342).Z,width:"359",height:"152"})))),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("p",{parentName:"li"},(0,r.kt)("strong",{parentName:"p"},"Step 2")," - Assigned input value to ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("inlineCode",{parentName:"strong"},"newName"))," variable."))))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' This will handle empty value or cancel case\nIf Len(newName) = 0 Then\n  MsgBox "Empty or no value. Please try again."\n  swDoc.ClearSelection2 True\n  Exit Sub\nEnd If\n')),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, we check the ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("em",{parentName:"strong"},"length of input value")),"."),(0,r.kt)("li",{parentName:"ul"},"This check will handle ",(0,r.kt)("strong",{parentName:"li"},"case for empty value")," or ",(0,r.kt)("strong",{parentName:"li"},"cancel operation case"),"."),(0,r.kt)("li",{parentName:"ul"},"We use \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-if-then-structure-select-case/"},"IF statement"))," for checking."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Condition"),": ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Len(newName) = 0")),(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Len()"))," is pre-build VBA function which check the length of a object."),(0,r.kt)("li",{parentName:"ul"},"In above cases, we will get ",(0,r.kt)("strong",{parentName:"li"},"0")," value."))),(0,r.kt)("li",{parentName:"ul"},"When this condition is ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"True")),", ",(0,r.kt)("ul",{parentName:"li"},(0,r.kt)("li",{parentName:"ul"},"We show and \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"/vba/vba-msgBox-function/"},"message window"))," to user."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Message"),": ",(0,r.kt)("em",{parentName:"li"},"Empty or no value. Please try again.")),(0,r.kt)("li",{parentName:"ul"},"Then we ",(0,r.kt)("strong",{parentName:"li"},"clear all selection")," and ",(0,r.kt)("strong",{parentName:"li"},"stop")," our macro here.")))),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Update mate's name\nswMateFeature.Name = newName\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above code block, ",(0,r.kt)("em",{parentName:"li"},"we update selected mate's name"),"."),(0,r.kt)("li",{parentName:"ul"},"For this we set the value ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"Name"))," property of ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("inlineCode",{parentName:"strong"},"swMateFeature"))," variable."),(0,r.kt)("li",{parentName:"ul"},(0,r.kt)("strong",{parentName:"li"},"Reference"),": Please visit \ud83d\ude80 ",(0,r.kt)("strong",{parentName:"li"},(0,r.kt)("a",{parentName:"strong",href:"https://help.solidworks.com/2021/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeature~Name.html"},"online SOLIDWORKS API Help"))," for more help.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Clear all selection\nswDoc.ClearSelection2 True\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we clear all selection."),(0,r.kt)("li",{parentName:"ul"},"For this we use ",(0,r.kt)("inlineCode",{parentName:"li"},"ClearSelection2")," method which is part of ",(0,r.kt)("em",{parentName:"li"},"SOLIDWORKS Document")," variable i.e ",(0,r.kt)("inlineCode",{parentName:"li"},"swDoc")," variable.")),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Rebuild assembly\nswDoc.ForceRebuild3 True\n")),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"In above line, we Rebuild assembly."),(0,r.kt)("li",{parentName:"ul"},"For this we use ",(0,r.kt)("inlineCode",{parentName:"li"},"ForceRebuild3")," method which is part of ",(0,r.kt)("em",{parentName:"li"},"SOLIDWORKS Document")," variable i.e ",(0,r.kt)("inlineCode",{parentName:"li"},"swDoc")," variable.")),(0,r.kt)("p",null,"Now we run the macro and after running macro we get ",(0,r.kt)("strong",{parentName:"p"},"Rename Mate")," as shown in below image."),(0,r.kt)("p",null,(0,r.kt)("a",{target:"_blank",href:a(4907).Z},(0,r.kt)("img",{alt:"assembly-rename-mate",src:a(38930).Z,width:"1366",height:"728"}))),(0,r.kt)("p",null,(0,r.kt)("strong",{parentName:"p"},"This is it !!!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")),(0,r.kt)("p",null,"If you found anything to ",(0,r.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,r.kt)("em",{parentName:"p"},"e-mail"),"."),(0,r.kt)("p",null,"Hope this post helps you to ",(0,r.kt)("strong",{parentName:"p"},"Rename Mate or any Feature")," with SOLIDWORKS VBA Macros."),(0,r.kt)("p",null,"For more such tutorials on ",(0,r.kt)("strong",{parentName:"p"},"SOLIDWORKS VBA Macro"),", do come to this website after sometime."),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,r.kt)("p",null,(0,r.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}k.isMDXComponent=!0},74753:(e,t,a)=>{a.d(t,{Z:()=>r});var n=a(67294);class r extends n.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return n.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}},4907:(e,t,a)=>{a.d(t,{Z:()=>n});const n=a.p+"assets/files/final-result-gif-117b28f124711f1cbe089ea8a36b8e49.gif"},97533:(e,t,a)=>{a.d(t,{Z:()=>n});const n=a.p+"assets/files/message-for-new-name-79d1ec860f28e3580da4ea770f5a7c5d.png"},79456:(e,t,a)=>{a.d(t,{Z:()=>n});const n=a.p+"assets/files/prerequisite-e9123db6e5f19e01e561d089781be4ea.gif"},38930:(e,t,a)=>{a.d(t,{Z:()=>n});const n=a.p+"assets/images/final-result-gif-117b28f124711f1cbe089ea8a36b8e49.gif"},97342:(e,t,a)=>{a.d(t,{Z:()=>n});const n="data:image/png;base64,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"},63726:(e,t,a)=>{a.d(t,{Z:()=>n});const n=a.p+"assets/images/prerequisite-e9123db6e5f19e01e561d089781be4ea.gif"}}]);