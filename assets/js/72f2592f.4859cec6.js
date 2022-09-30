"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[7697],{3905:(e,t,n)=>{n.d(t,{Zo:()=>m,kt:()=>c});var a=n(67294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,a,o=function(e,t){if(null==e)return{};var n,a,o={},r=Object.keys(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var l=a.createContext({}),p=function(e){var t=a.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},m=function(e){var t=p(e.components);return a.createElement(l.Provider,{value:t},e.children)},d={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},u=a.forwardRef((function(e,t){var n=e.components,o=e.mdxType,r=e.originalType,l=e.parentName,m=s(e,["components","mdxType","originalType","parentName"]),u=p(n),c=o,k=u["".concat(l,".").concat(c)]||u[c]||d[c]||r;return n?a.createElement(k,i(i({ref:t},m),{},{components:n})):a.createElement(k,i({ref:t},m))}));function c(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var r=n.length,i=new Array(r);i[0]=u;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:o,i[1]=s;for(var p=2;p<r;p++)i[p]=n[p];return a.createElement.apply(null,i)}return a.createElement.apply(null,n)}u.displayName="MDXCreateElement"},85472:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>p,contentTitle:()=>s,default:()=>u,frontMatter:()=>i,metadata:()=>l,toc:()=>m});var a=n(87462),o=(n(67294),n(3905)),r=n(74753);const i={categories:"Solidworks-macro",title:"Solidworks Macro - Open new Part document",permalink:"/solidworks-macros/open-new-document/",tags:["Solidworks Macro"],id:"open-new-document"},s=void 0,l={unversionedId:"open-new-document",id:"open-new-document",title:"Solidworks Macro - Open new Part document",description:"",source:"@site/docs/solidworks-macros/001.2-open-new-document.md",sourceDirName:".",slug:"/open-new-document",permalink:"/solidworks-macros/open-new-document",draft:!1,tags:[{label:"Solidworks Macro",permalink:"/solidworks-macros/tags/solidworks-macro"}],version:"current",frontMatter:{categories:"Solidworks-macro",title:"Solidworks Macro - Open new Part document",permalink:"/solidworks-macros/open-new-document/",tags:["Solidworks Macro"],id:"open-new-document"},sidebar:"tutorialSidebar",previous:{title:"VBA In Solidworks",permalink:"/solidworks-macros/vba-in-solidworks"},next:{title:"Solidworks Macro - Open Assembly and Drawing document",permalink:"/solidworks-macros/open-assembly-and-drawing"}},p={},m=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"Code Sample",id:"code-sample",level:2}],d={toc:m};function u(e){let{components:t,...n}=e;return(0,o.kt)("wrapper",(0,a.Z)({},d,n,{components:t,mdxType:"MDXLayout"}),(0,o.kt)(r.Z,{mdxType:"AdComponent"}),(0,o.kt)("p",null,"In this post, we open new document ",(0,o.kt)("strong",{parentName:"p"},"from")," ",(0,o.kt)("em",{parentName:"p"},"Solidworks VBA macros"),"."),(0,o.kt)("p",null,"Also we ",(0,o.kt)("em",{parentName:"p"},"understand")," each and every line of written code. So that you can understand why we written those lines and get some knowledge about how to write macro properly."),(0,o.kt)("hr",null),(0,o.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,o.kt)("p",null,"Please see below \ud83c\udfac video for visual details."),(0,o.kt)("iframe",{src:"https://www.youtube.com/embed/SXrdQ0vrTyI",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,o.kt)("p",null,"Please note that there are ",(0,o.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,o.kt)("p",null,(0,o.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,o.kt)("p",null,"To do this, we first need to create a new empty macro. If you don't know how to create an empty macro; then please go to this \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"/solidworks-macros/vba-in-solidworks"},"VBA in Solidworks"))," post."),(0,o.kt)("hr",null),(0,o.kt)("h2",{id:"code-sample"},"Code Sample"),(0,o.kt)("p",null,"After creating an empty macro, you need to copy paste below code into this empty macro."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n\n' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n\n' Main function of our VBA program\nSub main()\n\n    ' Setting Solidworks variable to Solidworks application\n    Set swApp = Application.SldWorks\n    \n    ' Creating string type variable for storing default part location\n    Dim defaultTemplate As String\n    ' Setting value of this string type variable to \"Default part template\"\n    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n\n    ' Setting Solidworks document to new part document\n    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n\nEnd Sub\n")),(0,o.kt)("p",null,"This code opens ",(0,o.kt)("strong",{parentName:"p"},"a new default part")," template in Solidworks."),(0,o.kt)("p",null,"Now let us walk through ",(0,o.kt)("em",{parentName:"p"},"each line")," in the above code, and ",(0,o.kt)("strong",{parentName:"p"},"understand")," the meaning of every line."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Option Explicit\n")),(0,o.kt)("p",null,"This line forces us to define every variable we are going to use. "),(0,o.kt)("p",null,"This is ",(0,o.kt)("strong",{parentName:"p"},"very important")," because if you don't declare above line, it is very difficult to caught ",(0,o.kt)("em",{parentName:"p"},"typo errors")," in variable names."),(0,o.kt)("p",null,"This type of error comes, when you mistakenly type wrong spelling of your defined variable."),(0,o.kt)("p",null,"In this case, VBE thinks that you have defined a new variable and use this variable. "),(0,o.kt)("p",null,"This causes issues because your program runs perfectly but you didn't get the desired result."),(0,o.kt)("p",null,"This most of the time discourage people and ultimately they left the programming."),(0,o.kt)("p",null,"So be on safe side and use this ",(0,o.kt)("strong",{parentName:"p"},"Option Explicit")," line in your every module."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,o.kt)("p",null,"As the comments in above code sample shows, in these 2 lines we are creating variables of different type."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks application\nDim swApp As SldWorks.SldWorks\n")),(0,o.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp")," and the type of this ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp")," variable is ",(0,o.kt)("inlineCode",{parentName:"p"},"SldWorks.SldWorks"),"."),(0,o.kt)("p",null,"If we ",(0,o.kt)("strong",{parentName:"p"},"omit")," 1 ",(0,o.kt)("inlineCode",{parentName:"p"},"SldWorks"),", then our ",(0,o.kt)("em",{parentName:"p"},"VBE")," show error if we try to run this macro."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating variable for Solidworks document\nDim swDoc As SldWorks.ModelDoc2\n")),(0,o.kt)("p",null,"In this line, we are creating a variable which we named as ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc")," and the type of this ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc")," variable is ",(0,o.kt)("inlineCode",{parentName:"p"},"SldWorks.ModelDoc2"),"."),(0,o.kt)("p",null,"Here, if we omit ",(0,o.kt)("inlineCode",{parentName:"p"},"SldWorks"),", the compiler won't give error. I used it to know who is the parent object."),(0,o.kt)("p",null,(0,o.kt)("inlineCode",{parentName:"p"},"ModelDoc2")," is the a ",(0,o.kt)("strong",{parentName:"p"},"object"),", which holds properties and methods related to this a document."),(0,o.kt)("p",null,"These methods and properties are common to ",(0,o.kt)("em",{parentName:"p"},"part"),", ",(0,o.kt)("em",{parentName:"p"},"assembly")," and ",(0,o.kt)("em",{parentName:"p"},"drawing")," documents."),(0,o.kt)("p",null,"You can see more about ",(0,o.kt)("inlineCode",{parentName:"p"},"ModelDoc2")," in this help \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html?verRedirect=1"},"link")),"."),(0,o.kt)("p",null,"Next is our ",(0,o.kt)("inlineCode",{parentName:"p"},"Sub")," procedure named ",(0,o.kt)("inlineCode",{parentName:"p"},"main"),". This procedure hold all the ",(0,o.kt)("em",{parentName:"p"},"statements (instructions)")," we give to computer."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks variable to Solidworks application\nSet swApp = Application.SldWorks\n")),(0,o.kt)("p",null,"In this line, we are setting the value of our Solidworks variable which we define earlier to Solidworks application."),(0,o.kt)("admonition",{type:"tip"},(0,o.kt)("p",{parentName:"admonition"},"Now it important to know that when you defined a variable of different type, which is not a common type, then you need to set the variable also.")),(0,o.kt)("p",null,"This is a standard way to set Solidworks application. This way is given in many Solidworks API VBA example."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Creating string type variable for storing default part location\nDim defaultTemplate As String\n' Setting value of this string type variable to \"Default part template\"\ndefaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)\n")),(0,o.kt)("p",null,"In 1st statement of above example, we are defining a variable of ",(0,o.kt)("inlineCode",{parentName:"p"},"string")," type and named it as ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),"."),(0,o.kt)("p",null,"This variable ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate"),", hold the location the location of ",(0,o.kt)("strong",{parentName:"p"},"Default Part Template"),"."),(0,o.kt)("p",null,"In 2nd line of above example. we assign value to our newly define ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate")," variable."),(0,o.kt)("p",null,"We assign the value by using a ",(0,o.kt)("em",{parentName:"p"},"Method")," named ",(0,o.kt)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". This method is a part of our main Solidworks variable ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp"),"."),(0,o.kt)("p",null,"To access this method, we need to write ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp"),' variable and then use a "." to access the ',(0,o.kt)("em",{parentName:"p"},"Public properties")," and ",(0,o.kt)("em",{parentName:"p"},"Methods")," inside this variable."),(0,o.kt)("admonition",{type:"info"},(0,o.kt)("p",{parentName:"admonition"},'This "." is called ',(0,o.kt)("strong",{parentName:"p"},"Dot operator"),". This operator provides us the access to the ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"Public properties"))," and ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("em",{parentName:"strong"},"Methods"))," inside an object.")),(0,o.kt)("p",null,'When you type after a "." you will notice that ',(0,o.kt)("em",{parentName:"p"},"Visual Basic Editor")," automatically provides a list of properties and methods inside this ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp")," object. This helps us to write correct name for these methods and properties."),(0,o.kt)("p",null,"Now we get the function ",(0,o.kt)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()"),". But this function needs some input to work with. This inputs are generally called ",(0,o.kt)("strong",{parentName:"p"},"Parameters"),"."),(0,o.kt)("admonition",{type:"info"},(0,o.kt)("p",{parentName:"admonition"},"In programming voculabury, we need to pass the parameter to this function so that this function can worked.")),(0,o.kt)("p",null,"This input parameter is a single value from a list of other values. This list is stored in ",(0,o.kt)("inlineCode",{parentName:"p"},"swUserPreferenceStringValue_e")," object."),(0,o.kt)("p",null,"In Solidworks API, if anything has ",(0,o.kt)("inlineCode",{parentName:"p"},"_e")," after it, it means that this object contains some type of lists. It is important to know because we frequently use these type of lists. They are called ",(0,o.kt)("strong",{parentName:"p"},"enum"),". The value they hold is called ",(0,o.kt)("strong",{parentName:"p"},"Constant"),"."),(0,o.kt)("p",null,"So our function ",(0,o.kt)("inlineCode",{parentName:"p"},"GetUserPreferenceStringValue()")," needs some constant value from ",(0,o.kt)("inlineCode",{parentName:"p"},"swUserPreferenceStringValue_e")," enum list to work."),(0,o.kt)("p",null,"Since we want ",(0,o.kt)("em",{parentName:"p"},"Default part template"),", we use ",(0,o.kt)("inlineCode",{parentName:"p"},"swDefaultTemplatePart")," constant value from the ",(0,o.kt)("inlineCode",{parentName:"p"},"swUserPreferenceStringValue_e")," enum list."),(0,o.kt)("admonition",{type:"info"},(0,o.kt)("p",{parentName:"admonition"},"Please note that there are lots of values inside this enum list. You can see these values from this \ud83d\ude80 ",(0,o.kt)("strong",{parentName:"p"},(0,o.kt)("a",{parentName:"strong",href:"http://help.solidworks.com/2019/English/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swUserPreferenceStringValue_e.html"},"link")),".")),(0,o.kt)("p",null,"Now we just need to set the value of our ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc")," variable to new document. We set the value as shown in below code snippet."),(0,o.kt)("pre",null,(0,o.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Setting Solidworks document to new part document\nSet swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)\n")),(0,o.kt)("p",null,"To set the value of ",(0,o.kt)("inlineCode",{parentName:"p"},"swDoc"),", we use ",(0,o.kt)("inlineCode",{parentName:"p"},"NewDocument()")," method. This method is inside ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp"),", hence we first need to invoke ",(0,o.kt)("inlineCode",{parentName:"p"},"swApp")," and then by using ",(0,o.kt)("em",{parentName:"p"},"Dot operator")," we access the ",(0,o.kt)("inlineCode",{parentName:"p"},"NewDocument()")," method."),(0,o.kt)("p",null,"Now this method needs 4 parameters (or input values) to work. If we don't provide any of these required value we get errors."),(0,o.kt)("p",null,"These 4 parameters are as follows:"),(0,o.kt)("ul",null,(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"TemplateName")," - ",(0,o.kt)("em",{parentName:"li"},"This can be a full path of the template, which we use to create New document"),"."),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"PaperSize")," - ",(0,o.kt)("em",{parentName:"li"},"Size of paper")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Width")," - ",(0,o.kt)("em",{parentName:"li"},"Width of paper")),(0,o.kt)("li",{parentName:"ul"},(0,o.kt)("strong",{parentName:"li"},"Height")," - ",(0,o.kt)("em",{parentName:"li"},"Height of paper"))),(0,o.kt)("p",null,"When successfully implemented, this method creates a new document. "),(0,o.kt)("p",null,"If this method failes to create a new document in that case this return ",(0,o.kt)("inlineCode",{parentName:"p"},"NULL")," value. We can use this ",(0,o.kt)("inlineCode",{parentName:"p"},"NULL")," value to check if the operation is successfull or not."),(0,o.kt)("p",null,"In our example, we use ",(0,o.kt)("inlineCode",{parentName:"p"},"defaultTemplate")," variable as ",(0,o.kt)("em",{parentName:"p"},"TemplateName")," parameter and use ",(0,o.kt)("strong",{parentName:"p"},"0")," in all other 3 parameter."),(0,o.kt)("admonition",{type:"info"},(0,o.kt)("p",{parentName:"admonition"},"Please note that ",(0,o.kt)("em",{parentName:"p"},"PaperSize"),", ",(0,o.kt)("em",{parentName:"p"},"Width")," and ",(0,o.kt)("em",{parentName:"p"},"Height")," is used only if we want to create a new ",(0,o.kt)("strong",{parentName:"p"},"Drawing document"),".")),(0,o.kt)("p",null,"This is all for now. This post is getting too long. "),(0,o.kt)("p",null,"In next post I will tell you how to create a new ",(0,o.kt)("em",{parentName:"p"},"Assembly")," document and new ",(0,o.kt)("em",{parentName:"p"},"Drawing")," document."))}u.isMDXComponent=!0},74753:(e,t,n)=>{n.d(t,{Z:()=>o});var a=n(67294);class o extends a.Component{componentDidMount(){(()=>{const e=document.createElement("script");e.src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js",e.async=!0,e.defer=!0,document.body.insertBefore(e,document.body.firstChild)})(),(window.adsbygoogle=window.adsbygoogle||[]).push({})}render(){return a.createElement("ins",{className:"adsbygoogle",style:{display:"block"},"data-ad-client":"ca-pub-8158659264340002","data-ad-slot":"6644001766","data-ad-format":"auto","data-full-width-responsive":"true"})}}}}]);