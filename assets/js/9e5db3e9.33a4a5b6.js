"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[4154],{3905:function(e,t,n){n.d(t,{Zo:function(){return u},kt:function(){return g}});var r=n(67294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function a(e,t){if(null==e)return{};var n,r,o=function(e,t){if(null==e)return{};var n,r,o={},i=Object.keys(e);for(r=0;r<i.length;r++)n=i[r],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)n=i[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var s=r.createContext({}),c=function(e){var t=r.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},u=function(e){var t=c(e.components);return r.createElement(s.Provider,{value:t},e.children)},p={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},m=r.forwardRef((function(e,t){var n=e.components,o=e.mdxType,i=e.originalType,s=e.parentName,u=a(e,["components","mdxType","originalType","parentName"]),m=c(n),g=o,d=m["".concat(s,".").concat(g)]||m[g]||p[g]||i;return n?r.createElement(d,l(l({ref:t},u),{},{components:n})):r.createElement(d,l({ref:t},u))}));function g(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var i=n.length,l=new Array(i);l[0]=m;var a={};for(var s in t)hasOwnProperty.call(t,s)&&(a[s]=t[s]);a.originalType=e,a.mdxType="string"==typeof e?e:o,l[1]=a;for(var c=2;c<i;c++)l[c]=n[c];return r.createElement.apply(null,l)}return r.createElement.apply(null,n)}m.displayName="MDXCreateElement"},96639:function(e,t,n){n.r(t),n.d(t,{assets:function(){return u},contentTitle:function(){return s},default:function(){return g},frontMatter:function(){return a},metadata:function(){return c},toc:function(){return p}});var r=n(87462),o=n(63366),i=(n(67294),n(3905)),l=["components"],a={title:"Controlling Program Flow and Making Decisions",tags:["VBA"],permalink:"/vba/controlling-flow-making-desicions/"},s=void 0,c={unversionedId:"vba-controlling-flow-making-desicions",id:"vba-controlling-flow-making-desicions",title:"Controlling Program Flow and Making Decisions",description:"Some VBA procedures start at the code\u2019s beginning and progress line by line to the end, never deviating from this top-to-bottom program flow.",source:"@site/docs/vba/17-vba-controlling-flow-making-desicions.md",sourceDirName:".",slug:"/vba-controlling-flow-making-desicions",permalink:"/vba/vba-controlling-flow-making-desicions",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:17,frontMatter:{title:"Controlling Program Flow and Making Decisions",tags:["VBA"],permalink:"/vba/controlling-flow-making-desicions/"},sidebar:"tutorialSidebar",previous:{title:"VBA Functions that do more",permalink:"/vba/vba-more-function"},next:{title:"If-Then-Else and Select Case structure",permalink:"/vba/vba-if-then-structure-select-case"}},u={},p=[],m={toc:p};function g(e){var t=e.components,n=(0,o.Z)(e,l);return(0,i.kt)("wrapper",(0,r.Z)({},m,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("p",null,"Some VBA ",(0,i.kt)("em",{parentName:"p"},"procedures")," start at the code\u2019s beginning and progress line by line to the end, never deviating from this top-to-bottom program flow. "),(0,i.kt)("p",null,"Macros that you record always work like this. "),(0,i.kt)("p",null,"In many cases, however, you need to ",(0,i.kt)("em",{parentName:"p"},"control")," the flow of your code by skipping over some statements, executing some statements multiple times, and testing conditions to determine what the procedure does next. "),(0,i.kt)("p",null,"Some programming newbies can\u2019t understand how a dumb computer can make intelligent decisions. "),(0,i.kt)("p",null,"The secret is in several programming constructs that most programming languages support. "),(0,i.kt)("p",null,"Following table provides a quick summary of these constructs. "),(0,i.kt)("table",{class:"w3-table-all w3-mobile w3-card-4"},(0,i.kt)("tr",null,(0,i.kt)("th",{class:"w3-center",colspan:"2"},"Programming Constructs for Making Decisions")),(0,i.kt)("tr",null,(0,i.kt)("th",null,"Construct"),(0,i.kt)("th",null,"How it works")),(0,i.kt)("tr",null,(0,i.kt)("td",null,"If-Then structure"),(0,i.kt)("td",null,"Does something if something else is true.")),(0,i.kt)("tr",null,(0,i.kt)("td",null,"Select Case"),(0,i.kt)("td",null,"Does any of several things, depending on something\u2019s value.")),(0,i.kt)("tr",null,(0,i.kt)("td",null,"For-Next loop"),(0,i.kt)("td",null,"Executes a series of statements a specified number of times.")),(0,i.kt)("tr",null,(0,i.kt)("td",null,"Do-While loop"),(0,i.kt)("td",null,"Does something as long as something else remains true.")),(0,i.kt)("tr",null,(0,i.kt)("td",null,"Do-Until loop"),(0,i.kt)("td",null,"Does something until something else becomes true."))),(0,i.kt)("p",null,"Next post will be about ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("em",{parentName:"strong"},"If-Then-Else and Select Case structure")),"."))}g.isMDXComponent=!0}}]);