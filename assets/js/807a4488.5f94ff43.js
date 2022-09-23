"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[1503],{3905:(e,t,r)=>{r.d(t,{Zo:()=>c,kt:()=>k});var o=r(67294);function n(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function i(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,o)}return r}function a(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?i(Object(r),!0).forEach((function(t){n(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):i(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function p(e,t){if(null==e)return{};var r,o,n=function(e,t){if(null==e)return{};var r,o,n={},i=Object.keys(e);for(o=0;o<i.length;o++)r=i[o],t.indexOf(r)>=0||(n[r]=e[r]);return n}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)r=i[o],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(n[r]=e[r])}return n}var s=o.createContext({}),l=function(e){var t=o.useContext(s),r=t;return e&&(r="function"==typeof e?e(t):a(a({},t),e)),r},c=function(e){var t=l(e.components);return o.createElement(s.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},u=o.forwardRef((function(e,t){var r=e.components,n=e.mdxType,i=e.originalType,s=e.parentName,c=p(e,["components","mdxType","originalType","parentName"]),u=l(r),k=n,d=u["".concat(s,".").concat(k)]||u[k]||m[k]||i;return r?o.createElement(d,a(a({ref:t},c),{},{components:r})):o.createElement(d,a({ref:t},c))}));function k(e,t){var r=arguments,n=t&&t.mdxType;if("string"==typeof e||n){var i=r.length,a=new Array(i);a[0]=u;var p={};for(var s in t)hasOwnProperty.call(t,s)&&(p[s]=t[s]);p.originalType=e,p.mdxType="string"==typeof e?e:n,a[1]=p;for(var l=2;l<i;l++)a[l]=r[l];return o.createElement.apply(null,a)}return o.createElement.apply(null,r)}u.displayName="MDXCreateElement"},10524:(e,t,r)=>{r.r(t),r.d(t,{assets:()=>s,contentTitle:()=>a,default:()=>m,frontMatter:()=>i,metadata:()=>p,toc:()=>l});var o=r(87462),n=(r(67294),r(3905));const i={title:"Solidworks C++ API",permalink:"/Solidworks-cpp-api-tutorials/",tags:["Solidworks C++ API"],toc:!1,id:"solidworks-Cpp-Api"},a=void 0,p={unversionedId:"solidworks-Cpp-Api",id:"solidworks-Cpp-Api",title:"Solidworks C++ API",description:"This page is only for my learning purpose!!!",source:"@site/docs/solidworks-cpp/001.1-solidworks-Cpp-Api.md",sourceDirName:".",slug:"/solidworks-Cpp-Api",permalink:"/solidworks-cpp/solidworks-Cpp-Api",draft:!1,tags:[{label:"Solidworks C++ API",permalink:"/solidworks-cpp/tags/solidworks-c-api"}],version:"current",frontMatter:{title:"Solidworks C++ API",permalink:"/Solidworks-cpp-api-tutorials/",tags:["Solidworks C++ API"],toc:!1,id:"solidworks-Cpp-Api"},sidebar:"tutorialSidebar",next:{title:"Solidworks C++ API - Prerequisite",permalink:"/solidworks-cpp/cpp-prerequisite"}},s={},l=[],c={toc:l};function m(e){let{components:t,...r}=e;return(0,n.kt)("wrapper",(0,o.Z)({},c,r,{components:t,mdxType:"MDXLayout"}),(0,n.kt)("admonition",{title:"Disclaimer",type:"danger"},(0,n.kt)("p",{parentName:"admonition"},"This page is only for my learning purpose!!!")),(0,n.kt)("p",null,"Nothing special I have to say or write for this ",(0,n.kt)("em",{parentName:"p"},"Solidworks C++ API")," tutorials."),(0,n.kt)("p",null,"I will start posting on ",(0,n.kt)("em",{parentName:"p"},"Solidworks C++ API")," tutorials along with \ud83d\ude80 ",(0,n.kt)("strong",{parentName:"p"},(0,n.kt)("a",{parentName:"strong",href:"/solidworks-macros/vba-in-solidworks/"},"Solidworks VBA tutorials"))," on ",(0,n.kt)("strong",{parentName:"p"},"casual basis")," (Whenever I got time)."),(0,n.kt)("p",null,"This section might be interested for:"),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},(0,n.kt)("em",{parentName:"p"},"Those who want to learn Solidworks C++ API"))),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},(0,n.kt)("em",{parentName:"p"},"Student doing Masters in CAD/CAM"))),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},(0,n.kt)("em",{parentName:"p"},"Casual programmers like myself who want to apply what is learnt!")))),(0,n.kt)("p",null,"One thing is sure, I am not going to explain in very detail because I am more of a ",(0,n.kt)("strong",{parentName:"p"},(0,n.kt)("inlineCode",{parentName:"strong"},".NET developer"))," not a ",(0,n.kt)("strong",{parentName:"p"},(0,n.kt)("inlineCode",{parentName:"strong"},"C++ Developer")),"."),(0,n.kt)("p",null,"I just want to explore ",(0,n.kt)("strong",{parentName:"p"},"C++ in Visual studio")," and best way to do is writing some program using my existing knowledge of ",(0,n.kt)("strong",{parentName:"p"},(0,n.kt)("em",{parentName:"strong"},"Solidworks C# API")),"."),(0,n.kt)("p",null,"So ",(0,n.kt)("strong",{parentName:"p"},(0,n.kt)("em",{parentName:"strong"},"Solidworks C++ API")),' posts are more "',(0,n.kt)("strong",{parentName:"p"},"how"),'" to type not much descriptive like \ud83d\ude80 ',(0,n.kt)("strong",{parentName:"p"},(0,n.kt)("a",{parentName:"strong",href:"/solidworks-macros/vba-in-solidworks"},"Solidworks VBA posts")),", which tends to describe in detailed manner."),(0,n.kt)("p",null,"That's it!"),(0,n.kt)("p",null,"I hope you will like this and enjoy this section also."),(0,n.kt)("p",null,"Thank you!"))}m.isMDXComponent=!0}}]);