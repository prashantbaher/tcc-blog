"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3634],{3905:function(e,t,n){n.d(t,{Zo:function(){return p},kt:function(){return d}});var r=n(67294);function o(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);t&&(r=r.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,r)}return n}function a(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){o(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function u(e,t){if(null==e)return{};var n,r,o=function(e,t){if(null==e)return{};var n,r,o={},i=Object.keys(e);for(r=0;r<i.length;r++)n=i[r],t.indexOf(n)>=0||(o[n]=e[n]);return o}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(r=0;r<i.length;r++)n=i[r],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(o[n]=e[n])}return o}var s=r.createContext({}),l=function(e){var t=r.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):a(a({},t),e)),n},p=function(e){var t=l(e.components);return r.createElement(s.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return r.createElement(r.Fragment,{},t)}},m=r.forwardRef((function(e,t){var n=e.components,o=e.mdxType,i=e.originalType,s=e.parentName,p=u(e,["components","mdxType","originalType","parentName"]),m=l(n),d=o,f=m["".concat(s,".").concat(d)]||m[d]||c[d]||i;return n?r.createElement(f,a(a({ref:t},p),{},{components:n})):r.createElement(f,a({ref:t},p))}));function d(e,t){var n=arguments,o=t&&t.mdxType;if("string"==typeof e||o){var i=n.length,a=new Array(i);a[0]=m;var u={};for(var s in t)hasOwnProperty.call(t,s)&&(u[s]=t[s]);u.originalType=e,u.mdxType="string"==typeof e?e:o,a[1]=u;for(var l=2;l<i;l++)a[l]=n[l];return r.createElement.apply(null,a)}return r.createElement.apply(null,n)}m.displayName="MDXCreateElement"},58431:function(e,t,n){n.r(t),n.d(t,{assets:function(){return p},contentTitle:function(){return s},default:function(){return d},frontMatter:function(){return u},metadata:function(){return l},toc:function(){return c}});var r=n(87462),o=n(63366),i=(n(67294),n(3905)),a=["components"],u={title:"VBA Bug Reduction Tips",tags:["VBA"],permalink:"/vba/bug-reduction-tips/"},s=void 0,l={unversionedId:"vba-bug-reduction-tips",id:"vba-bug-reduction-tips",title:"VBA Bug Reduction Tips",description:"I can\u2019t tell you how to completely eliminate bugs in your programs.",source:"@site/docs/vba/22-vba-bug-reduction-tips.md",sourceDirName:".",slug:"/vba-bug-reduction-tips",permalink:"/vba/vba-bug-reduction-tips",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:22,frontMatter:{title:"VBA Bug Reduction Tips",tags:["VBA"],permalink:"/vba/bug-reduction-tips/"},sidebar:"tutorialSidebar",previous:{title:"VBA Debugger",permalink:"/vba/vba-debugger"},next:{title:"VBA Dialog Boxes",permalink:"/vba/vba-dialog-boxes"}},p={},c=[],m={toc:c};function d(e){var t=e.components,n=(0,o.Z)(e,a);return(0,i.kt)("wrapper",(0,r.Z)({},m,n,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("p",null,"I can\u2019t tell you how to completely eliminate bugs in your programs. "),(0,i.kt)("p",null,"Finding bugs in software can be a profession by itself, but I can provide a few tips to help you keep those bugs to a minimum:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"Use an ",(0,i.kt)("inlineCode",{parentName:"p"},"Option Explicit")," statement at the beginning of your modules. This statement requires you to define the data type for every variable you use. This creates a bit more work for you, but you avoid the common error of misspelling a variable name. And it has a nice side benefit: ",(0,i.kt)("em",{parentName:"p"},"Your routines run a bit faster."))),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"Format your code with ",(0,i.kt)("strong",{parentName:"p"},"indentation"),". Using indentations helps delineate different code segments. If your program has several nested ",(0,i.kt)("inlineCode",{parentName:"p"},"For-Next")," loops, for example, consistent indentation helps you keep track of them all.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"Use lots of ",(0,i.kt)("strong",{parentName:"p"},"comments"),". Nothing is more frustrating than revisiting code you wrote six months ago and not having a clue as to how it works. By adding a few comments to describe your logic, you can save lots of time down the road.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"Keep your ",(0,i.kt)("inlineCode",{parentName:"p"},"Sub")," and ",(0,i.kt)("inlineCode",{parentName:"p"},"Function")," procedures simple. By writing your code in small modules, each of which has a single, well-defined purpose, you simplify the debugging process.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"Use the macro recorder to help identify properties and methods. When I can\u2019t remember the name or the syntax of a property or method, I often simply record a macro and look at the recorded code"))),(0,i.kt)("p",null,"Debugging code is not one of my favorite activities, but it\u2019s a necessary evil that goes along with programming. "),(0,i.kt)("p",null,"As you gain more experience with VBA, you spend less time debugging and, when you have to debug, are more efficient at doing so."),(0,i.kt)("p",null,"Next post will be about ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("em",{parentName:"strong"},"VBA Dialog Boxes")),"."))}d.isMDXComponent=!0}}]);