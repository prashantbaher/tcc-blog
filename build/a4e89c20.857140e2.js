(window.webpackJsonp=window.webpackJsonp||[]).push([[60],{117:function(e,t,n){"use strict";n.r(t),n.d(t,"frontMatter",(function(){return i})),n.d(t,"metadata",(function(){return p})),n.d(t,"rightToc",(function(){return c})),n.d(t,"default",(function(){return l}));var r=n(2),a=n(6),o=(n(0),n(152)),i={id:"sw-cpp",title:"Solidworks C++ API"},p={unversionedId:"sw-cpp",id:"sw-cpp",isDocsHomePage:!1,title:"Solidworks C++ API",description:"Disclaimer, this page is only for my learning purpose!",source:"@site/docs\\solidworksCpp.md",slug:"/sw-cpp",permalink:"/docs/sw-cpp",version:"current",sidebar:"swcpp",next:{title:"Prerequisite",permalink:"/docs/solidworks-Cpp-tutorials/sw-cpp-pre"}},c=[],s={rightToc:c};function l(e){var t=e.components,n=Object(a.a)(e,["components"]);return Object(o.b)("wrapper",Object(r.a)({},s,n,{components:t,mdxType:"MDXLayout"}),Object(o.b)("div",{className:"admonition admonition-caution alert alert--warning"},Object(o.b)("div",Object(r.a)({parentName:"div"},{className:"admonition-heading"}),Object(o.b)("h5",{parentName:"div"},Object(o.b)("span",Object(r.a)({parentName:"h5"},{className:"admonition-icon"}),Object(o.b)("svg",Object(r.a)({parentName:"span"},{xmlns:"http://www.w3.org/2000/svg",width:"16",height:"16",viewBox:"0 0 16 16"}),Object(o.b)("path",Object(r.a)({parentName:"svg"},{fillRule:"evenodd",d:"M8.893 1.5c-.183-.31-.52-.5-.887-.5s-.703.19-.886.5L.138 13.499a.98.98 0 0 0 0 1.001c.193.31.53.501.886.501h13.964c.367 0 .704-.19.877-.5a1.03 1.03 0 0 0 .01-1.002L8.893 1.5zm.133 11.497H6.987v-2.003h2.039v2.003zm0-3.004H6.987V5.987h2.039v4.006z"})))),"caution")),Object(o.b)("div",Object(r.a)({parentName:"div"},{className:"admonition-content"}),Object(o.b)("p",{parentName:"div"},"Disclaimer, this page is only for my learning purpose!"))),Object(o.b)("p",null,"Nothing special I have to say or write for this ",Object(o.b)("em",{parentName:"p"},"Solidworks C++ API")," tutorials."),Object(o.b)("p",null,"I will start posting on ",Object(o.b)("em",{parentName:"p"},"Solidworks C++ API")," tutorials along with ",Object(o.b)("a",Object(r.a)({parentName:"p"},{href:"vba-Intro"}),"Solidworks VBA tutorials")," on regular basis."),Object(o.b)("p",null,"For whom this section might be interested?"),Object(o.b)("ul",null,Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Those who want to learn Solidworks C++ API"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Student doing Masters in CAD/CAM"))),Object(o.b)("li",{parentName:"ul"},Object(o.b)("p",{parentName:"li"},Object(o.b)("em",{parentName:"p"},"Casual programmers like myself who want to apply what is learnt!")))),Object(o.b)("p",null,"One thing is sure, I am not going to explain in very detail because I am more of a ",Object(o.b)("inlineCode",{parentName:"p"},".NET developer")," not a ",Object(o.b)("inlineCode",{parentName:"p"},"C++ Developer"),"."),Object(o.b)("p",null,"I just want to explore ",Object(o.b)("strong",{parentName:"p"},"C++ in Visual studio")," and best way to do is writing some program using my existing knowledge of ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Solidworks C# API")),"."),Object(o.b)("p",null,"So ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Solidworks C++ API")),' posts are more "',Object(o.b)("strong",{parentName:"p"},"how"),'" to type not much descriptive like ',Object(o.b)("strong",{parentName:"p"},Object(o.b)("a",Object(r.a)({parentName:"strong"},{href:"vba-in-sw"}),"Solidworks VBA posts")),", which tends to describe all in detailed manner."),Object(o.b)("p",null,"That's it!!!"),Object(o.b)("p",null,"I hope you will like this and enjoy this section also."),Object(o.b)("p",null,"Thanks!!!"),Object(o.b)("p",null,"Please see below list of posts for ",Object(o.b)("strong",{parentName:"p"},Object(o.b)("em",{parentName:"strong"},"Solidworks C++ API")),". "))}l.isMDXComponent=!0},152:function(e,t,n){"use strict";n.d(t,"a",(function(){return b})),n.d(t,"b",(function(){return d}));var r=n(0),a=n.n(r);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function p(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function c(e,t){if(null==e)return{};var n,r,a=function(e,t){if(null==e)return{};var n,r,a={},o=Object.keys(e);for(r=0;r<o.length;r++)n=o[r],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);for(r=0;r<o.length;r++)n=o[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var s=a.a.createContext({}),l=function(e){var t=a.a.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):p(p({},t),e)),n},b=function(e){var t=l(e.components);return a.a.createElement(s.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return a.a.createElement(a.a.Fragment,{},t)}},u=a.a.forwardRef((function(e,t){var n=e.components,r=e.mdxType,o=e.originalType,i=e.parentName,s=c(e,["components","mdxType","originalType","parentName"]),b=l(n),u=r,d=b["".concat(i,".").concat(u)]||b[u]||m[u]||o;return n?a.a.createElement(d,p(p({ref:t},s),{},{components:n})):a.a.createElement(d,p({ref:t},s))}));function d(e,t){var n=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var o=n.length,i=new Array(o);i[0]=u;var p={};for(var c in t)hasOwnProperty.call(t,c)&&(p[c]=t[c]);p.originalType=e,p.mdxType="string"==typeof e?e:r,i[1]=p;for(var s=2;s<o;s++)i[s]=n[s];return a.a.createElement.apply(null,i)}return a.a.createElement.apply(null,n)}u.displayName="MDXCreateElement"}}]);