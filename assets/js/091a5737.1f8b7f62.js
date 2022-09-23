"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3502],{3905:function(e,t,o){o.d(t,{Zo:function(){return c},kt:function(){return m}});var n=o(67294);function r(e,t,o){return t in e?Object.defineProperty(e,t,{value:o,enumerable:!0,configurable:!0,writable:!0}):e[t]=o,e}function a(e,t){var o=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),o.push.apply(o,n)}return o}function l(e){for(var t=1;t<arguments.length;t++){var o=null!=arguments[t]?arguments[t]:{};t%2?a(Object(o),!0).forEach((function(t){r(e,t,o[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(o)):a(Object(o)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(o,t))}))}return e}function i(e,t){if(null==e)return{};var o,n,r=function(e,t){if(null==e)return{};var o,n,r={},a=Object.keys(e);for(n=0;n<a.length;n++)o=a[n],t.indexOf(o)>=0||(r[o]=e[o]);return r}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(n=0;n<a.length;n++)o=a[n],t.indexOf(o)>=0||Object.prototype.propertyIsEnumerable.call(e,o)&&(r[o]=e[o])}return r}var p=n.createContext({}),s=function(e){var t=n.useContext(p),o=t;return e&&(o="function"==typeof e?e(t):l(l({},t),e)),o},c=function(e){var t=s(e.components);return n.createElement(p.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},d=n.forwardRef((function(e,t){var o=e.components,r=e.mdxType,a=e.originalType,p=e.parentName,c=i(e,["components","mdxType","originalType","parentName"]),d=s(o),m=r,k=d["".concat(p,".").concat(m)]||d[m]||u[m]||a;return o?n.createElement(k,l(l({ref:t},c),{},{components:o})):n.createElement(k,l({ref:t},c))}));function m(e,t){var o=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var a=o.length,l=new Array(a);l[0]=d;var i={};for(var p in t)hasOwnProperty.call(t,p)&&(i[p]=t[p]);i.originalType=e,i.mdxType="string"==typeof e?e:r,l[1]=i;for(var s=2;s<a;s++)l[s]=o[s];return n.createElement.apply(null,l)}return n.createElement.apply(null,o)}d.displayName="MDXCreateElement"},64411:function(e,t,o){o.r(t),o.d(t,{assets:function(){return c},contentTitle:function(){return p},default:function(){return m},frontMatter:function(){return i},metadata:function(){return s},toc:function(){return u}});var n=o(87462),r=o(63366),a=(o(67294),o(3905)),l=["components"],i={categories:"Solidworks-C++-API",title:"Open Solidworks & Hello World",tags:["Solidworks C++ API"],permalink:"/solidworks-cpp/open-solidworks/",id:"open-solidworks"},p=void 0,s={unversionedId:"open-solidworks",id:"open-solidworks",title:"Open Solidworks & Hello World",description:"In this post, I tell you about how to Open Solidworks using Solidworks C++ API from Visual Studio.",source:"@site/docs/solidworks-cpp/001.3-open-solidworks.mdx",sourceDirName:".",slug:"/open-solidworks",permalink:"/solidworks-cpp/open-solidworks",draft:!1,tags:[{label:"Solidworks C++ API",permalink:"/solidworks-cpp/tags/solidworks-c-api"}],version:"current",frontMatter:{categories:"Solidworks-C++-API",title:"Open Solidworks & Hello World",tags:["Solidworks C++ API"],permalink:"/solidworks-cpp/open-solidworks/",id:"open-solidworks"},sidebar:"tutorialSidebar",previous:{title:"Solidworks C++ API - Prerequisite",permalink:"/solidworks-cpp/cpp-prerequisite"},next:{title:"Open Solidworks Part Document",permalink:"/solidworks-cpp/open-part-document"}},c={},u=[{value:"Video of Code on YouTube",id:"video-of-code-on-youtube",level:2},{value:"Create a New project",id:"create-a-new-project",level:2},{value:"Add Source file",id:"add-source-file",level:2},{value:"Add References to Solidworks Type Library files",id:"add-references-to-solidworks-type-library-files",level:2},{value:"Add Code to Source.cpp file",id:"add-code-to-sourcecpp-file",level:2},{value:"Final Result",id:"final-result",level:2}],d={toc:u};function m(e){var t=e.components,i=(0,r.Z)(e,l);return(0,a.kt)("wrapper",(0,n.Z)({},d,i,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"In this post, I tell you about ",(0,a.kt)("strong",{parentName:"p"},"how to Open Solidworks using Solidworks C++ API")," from Visual Studio."),(0,a.kt)("p",null,"I hope you have setup Visual Studio community version."),(0,a.kt)("p",null,"If not then please go to \ud83d\ude80 ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("a",{parentName:"strong",href:"/solidworks-cpp/cpp-prerequisite"},"Solidworks C++ API - Prerequisite"))," post and watch the suggested videos before proceeding further."),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"video-of-code-on-youtube"},"Video of Code on YouTube"),(0,a.kt)("p",null,"Please see below video on ",(0,a.kt)("strong",{parentName:"p"},"how to Open Solidworks using Solidworks C++ API")," from Visual Studio."),(0,a.kt)("iframe",{src:"https://www.youtube.com/embed/oL9kJoRoYcQ",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,a.kt)("p",null,"Please note that there are ",(0,a.kt)("strong",{parentName:"p"},"no explaination")," in the video. "),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"Explaination")," of each line and why we write code this way is given in this post."),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"create-a-new-project"},"Create a New project"),(0,a.kt)("p",null,"Fist, we will create a new project in Visual Studio."),(0,a.kt)("p",null,"There are 3 different ways for creating a new project."),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"From ",(0,a.kt)("strong",{parentName:"p"},"File")," \u27a1 ",(0,a.kt)("strong",{parentName:"p"},"New")," \u27a1 ",(0,a.kt)("strong",{parentName:"p"},"Project"))),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"From ",(0,a.kt)("strong",{parentName:"p"},"New Project")," icon.")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"Keyboard Short-cut \u27a1 ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("inlineCode",{parentName:"strong"},"Ctrl+Shift+N")),"."))),(0,a.kt)("p",null,"Below image show how to create a ",(0,a.kt)("em",{parentName:"p"},'"New Project"')," from ",(0,a.kt)("strong",{parentName:"p"},"File")," option:"),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"new-project-file-option",src:o(93669).Z,width:"816",height:"458"})),(0,a.kt)("p",null,"In above image, see Red color box."),(0,a.kt)("p",null,"Below image show how to create a ",(0,a.kt)("em",{parentName:"p"},"New Project")," from ",(0,a.kt)("strong",{parentName:"p"},"New Project Icon")," option:"),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"new-project-from-icon",src:o(87344).Z,width:"818",height:"432"})),(0,a.kt)("p",null,"In above image, see Red color box."),(0,a.kt)("p",null,"When we select one of the above option we get a new window which is shown in below."),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"create-project",src:o(43481).Z,width:"946",height:"589"})),(0,a.kt)("p",null,"In above image I have numbered the Red colored box."),(0,a.kt)("p",null,"These numbers are explained below:"),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},(0,a.kt)("em",{parentName:"p"},"The programming language")," you want to use for ",(0,a.kt)("strong",{parentName:"p"},"New Project"),'. For our purpose, we use "',(0,a.kt)("em",{parentName:"p"},"Visual C++"),'".')),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"It is, ",(0,a.kt)("strong",{parentName:"p"},"which type")," of project you want to create. There are ",(0,a.kt)("em",{parentName:"p"},"3 different type")," of projects we can create. In above image, we will create ",(0,a.kt)("em",{parentName:"p"},"an empty project"),".")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"It is ",(0,a.kt)("em",{parentName:"p"},"the name of project")," we want to create. We named our project as ",(0,a.kt)("strong",{parentName:"p"},"OpenSolidworkTest"),".")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"The location of project we want. We use default location provided in above image.")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"It is option ",(0,a.kt)("em",{parentName:"p"},"if we want to create a Solution file for this project or not"),". In our case, we want to create a ",(0,a.kt)("em",{parentName:"p"},"Solution file"),".")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"Hit ",(0,a.kt)("strong",{parentName:"p"},"Ok")," button after completing all fields."))),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"add-source-file"},"Add Source file"),(0,a.kt)("p",null,"After creating a new project, we get a screen as shown in below image."),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"after-new-project",src:o(54445).Z,width:"1366",height:"740"})),(0,a.kt)("p",null,"This project has no file to write."),(0,a.kt)("p",null,"Now we add a cpp file into ",(0,a.kt)("em",{parentName:"p"},"Source Files filter folder"),"."),(0,a.kt)("p",null,"For this please follow given steps:"),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"For this select ",(0,a.kt)("em",{parentName:"p"},"Source Files filter folder")," and ",(0,a.kt)("em",{parentName:"p"},"Click Right Mouse Button (RMB)"),".")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"By doing this ",(0,a.kt)("em",{parentName:"p"},"a context menu")," is appear as shown in below image.")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"From this ",(0,a.kt)("em",{parentName:"p"},"context menu"),", select ",(0,a.kt)("strong",{parentName:"p"},'"Add"')," \u27a1 ",(0,a.kt)("strong",{parentName:"p"},'"New Item"'),", as shown in below image."))),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"add-new-cpp-file",src:o(75532).Z,width:"1366",height:"736"})),(0,a.kt)("p",null,"This will open a new window as shown in below image."),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"add-new-cpp-file-window",src:o(16389).Z,width:"946",height:"586"})),(0,a.kt)("p",null,'Just select "',(0,a.kt)("strong",{parentName:"p"},"Add"),'" option as shown in above image.'),(0,a.kt)("p",null,'This will add "',(0,a.kt)("strong",{parentName:"p"},"Source.cpp"),'" file into our project.'),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"add-references-to-solidworks-type-library-files"},"Add References to Solidworks Type Library files"),(0,a.kt)("p",null,"Now we need to ",(0,a.kt)("em",{parentName:"p"},"add References to Solidworks Type Library files.")),(0,a.kt)("p",null,"For this please follow below steps."),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"Select the ",(0,a.kt)("strong",{parentName:"p"},"OpenSolidworkTest")," project and and ",(0,a.kt)("em",{parentName:"p"},"Click Right Mouse Button (RMB)"),".")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"By doing this ",(0,a.kt)("em",{parentName:"p"},"a context menu")," is appear as shown in below image.")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"From this ",(0,a.kt)("em",{parentName:"p"},"context menu"),", select ",(0,a.kt)("strong",{parentName:"p"},'"Properties"')," option, which is the last one, as shown in below image."))),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"open-property-window",src:o(87243).Z,width:"634",height:"738"})),(0,a.kt)("p",null,"This will open a new window as shown in below image."),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"project-property-window",src:o(95115).Z,width:"824",height:"560"})),(0,a.kt)("p",null,"Now following below steps:"),(0,a.kt)("ol",null,(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},"Select C/C++ option")),(0,a.kt)("li",{parentName:"ol"},(0,a.kt)("p",{parentName:"li"},'Add SOLIDWORKS folders path to 2nd Red colored box as shown in below image. Usually this path is "',(0,a.kt)("inlineCode",{parentName:"p"},"C:\\Program Files\\ Solidworks Corp\\SOLIDWORKS"),'" if installed in default location.'))),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"add-solidowrks-reference",src:o(22603).Z,width:"825",height:"562"})),(0,a.kt)("p",null,'After adding the folder path, select "',(0,a.kt)("strong",{parentName:"p"},"Apply"),'" button.'),(0,a.kt)("p",null,"This complete the process of ",(0,a.kt)("em",{parentName:"p"},"adding References to Solidworks Type Library files.")),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"add-code-to-sourcecpp-file"},"Add Code to Source.cpp file"),(0,a.kt)("p",null,"Now we need to add to ",(0,a.kt)("em",{parentName:"p"},"Source.cpp")," file."),(0,a.kt)("p",null,"Please copy the below code sample to your ",(0,a.kt)("em",{parentName:"p"},"Source.cpp")," file."),(0,a.kt)("pre",null,(0,a.kt)("code",{parentName:"pre",className:"language-cpp",metastring:'title="Copy below code" showLineNumbers',title:'"Copy',below:!0,'code"':!0,showLineNumbers:!0},'#include <atlbase.h>\n\n#import "sldworks.tlb" raw_interfaces_only, raw_native_types, no_namespace, named_guids  // SOLIDWORKS type library\n\n#import "swconst.tlb" raw_interfaces_only, raw_native_types, no_namespace, named_guids   // SOLIDWORKS constants type library\n\nint main()\n{\n    // Initialize COM\n    // Do this before using ATL smart pointers so COM is available.\n    CoInitialize(NULL);\n\n    // Use a block, so the smart pointers are destructed when the scope of this block is left\n    {\n        // COM Pointer of Soldiworks object\n        CComPtr<ISldWorks> swApp;\n\n        // Create an instance of Solidworks application\n        // If it fails then return 0 and close program\n        if (swApp.CoCreateInstance(__uuidof(SldWorks), NULL, CLSCTX_LOCAL_SERVER) != S_OK) \n        {\n            // Stop COM \n            CoUninitialize();\n            return(0);\n        }\n\n        // If created successfully, then visible the Solidworks\n        swApp->put_Visible(VARIANT_TRUE);\n\n        // COM Style String for message to user\n        CComBSTR _messageToUser(L"Hello World!!! I am from Solidworks C++ API.");\n\n        // long type variable to store the result value by user\n        long _lMessageResult;\n\n        // Send a message to user and store the return value in _lMessageResult by referencing it\n        swApp->SendMsgToUser2(_messageToUser, swMessageBoxIcon_e::swMbInformation, swMessageBoxBtn_e::swMbOk, &_lMessageResult);\n    }\n\n    // Stop COM \n    CoUninitialize();\n}\n')),(0,a.kt)("p",null,"Now Build the Solution as shown in below image."),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"build-solution",src:o(34062).Z,width:"648",height:"166"})),(0,a.kt)("p",null,"After Building Solution run the program by pressing ",(0,a.kt)("strong",{parentName:"p"},"F5"),"."),(0,a.kt)("hr",null),(0,a.kt)("h2",{id:"final-result"},"Final Result"),(0,a.kt)("p",null,"After running the program wait for few minute."),(0,a.kt)("p",null,"You will get result as shown in below image!!!"),(0,a.kt)("p",null,(0,a.kt)("img",{alt:"hello-world-message.png",src:o(91538).Z,width:"1085",height:"588"})),(0,a.kt)("hr",null),(0,a.kt)("p",null,(0,a.kt)("strong",{parentName:"p"},"This is it !!!")),(0,a.kt)("p",null,"We have completed our ",(0,a.kt)("em",{parentName:"p"},"Hello World")," program in ",(0,a.kt)("em",{parentName:"p"},"Solidworks")," using ",(0,a.kt)("strong",{parentName:"p"},"Solidworks C++ APIs"),"."),(0,a.kt)("p",null,"Hope this post helps you to start with ",(0,a.kt)("em",{parentName:"p"},"Solidworks C++ API"),"."),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Do let me know by you like this post or not! I will continue creating Solidworks C++ posts.")),(0,a.kt)("p",null,(0,a.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}m.isMDXComponent=!0},16389:function(e,t,o){t.Z=o.p+"assets/images/add-new-cpp-file-window-0c8908b909b3b72299caecfc65f37477.png"},75532:function(e,t,o){t.Z=o.p+"assets/images/add-new-cpp-file-e6daf7554eb4075270cf943bfb03780c.png"},22603:function(e,t,o){t.Z=o.p+"assets/images/add-solidowrks-reference-1a1ca5269a89c20dcf68687bed4f57f1.png"},54445:function(e,t,o){t.Z=o.p+"assets/images/after-new-project-50c7ebbac93b95a3e78a043b3aae04ee.png"},34062:function(e,t,o){t.Z=o.p+"assets/images/build-solution-588c7601214b803be92cd586a148bb2e.png"},43481:function(e,t,o){t.Z=o.p+"assets/images/create-project-94a23117f2a7ba3c241cb6ed4d1d2810.png"},91538:function(e,t,o){t.Z=o.p+"assets/images/hello-world-message-121ebb35b5936cf28ab810f9556a4992.png"},93669:function(e,t,o){t.Z=o.p+"assets/images/new-project-1-183e52449b810d1d939bd1f19581eed9.png"},87344:function(e,t,o){t.Z=o.p+"assets/images/new-project-2-c2128bf2aea6c97b2b58abd8c8607618.png"},87243:function(e,t,o){t.Z=o.p+"assets/images/open-property-window-fbed5c07225a50d940c99768ebb68f86.png"},95115:function(e,t,o){t.Z=o.p+"assets/images/project-property-window-71bcffa26b32a7297231b56e5ae5d052.png"}}]);