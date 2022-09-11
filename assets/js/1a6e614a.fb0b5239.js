"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[6279],{3905:function(e,t,n){n.d(t,{Zo:function(){return p},kt:function(){return d}});var o=n(67294);function r(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function a(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?a(Object(n),!0).forEach((function(t){r(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):a(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function l(e,t){if(null==e)return{};var n,o,r=function(e,t){if(null==e)return{};var n,o,r={},a=Object.keys(e);for(o=0;o<a.length;o++)n=a[o],t.indexOf(n)>=0||(r[n]=e[n]);return r}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(o=0;o<a.length;o++)n=a[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(r[n]=e[n])}return r}var s=o.createContext({}),u=function(e){var t=o.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},p=function(e){var t=u(e.components);return o.createElement(s.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},m=o.forwardRef((function(e,t){var n=e.components,r=e.mdxType,a=e.originalType,s=e.parentName,p=l(e,["components","mdxType","originalType","parentName"]),m=u(n),d=r,f=m["".concat(s,".").concat(d)]||m[d]||c[d]||a;return n?o.createElement(f,i(i({ref:t},p),{},{components:n})):o.createElement(f,i({ref:t},p))}));function d(e,t){var n=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var a=n.length,i=new Array(a);i[0]=m;var l={};for(var s in t)hasOwnProperty.call(t,s)&&(l[s]=t[s]);l.originalType=e,l.mdxType="string"==typeof e?e:r,i[1]=l;for(var u=2;u<a;u++)i[u]=n[u];return o.createElement.apply(null,i)}return o.createElement.apply(null,n)}m.displayName="MDXCreateElement"},6397:function(e,t,n){n.r(t),n.d(t,{assets:function(){return p},contentTitle:function(){return s},default:function(){return d},frontMatter:function(){return l},metadata:function(){return u},toc:function(){return c}});var o=n(87462),r=n(63366),a=(n(67294),n(3905)),i=["components"],l={title:"VBA Dialog Boxes",tags:["VBA"],permalink:"/vba/dialog-boxes/"},s=void 0,u={unversionedId:"vba-dialog-boxes",id:"vba-dialog-boxes",title:"VBA Dialog Boxes",description:"You can\u2019t use VBA very long without being exposed to dialog boxes.",source:"@site/docs/vba/23-vba-dialog-boxes.md",sourceDirName:".",slug:"/vba-dialog-boxes",permalink:"/vba/vba-dialog-boxes",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:23,frontMatter:{title:"VBA Dialog Boxes",tags:["VBA"],permalink:"/vba/dialog-boxes/"},sidebar:"tutorialSidebar",previous:{title:"VBA Bug Reduction Tips",permalink:"/vba/vba-bug-reduction-tips"},next:{title:"VBA MsgBox Function",permalink:"/vba/vba-msgBox-function"}},p={},c=[{value:"UserForm Alternatives",id:"userform-alternatives",level:2}],m={toc:c};function d(e){var t=e.components,n=(0,r.Z)(e,i);return(0,a.kt)("wrapper",(0,o.Z)({},m,n,{components:t,mdxType:"MDXLayout"}),(0,a.kt)("p",null,"You can\u2019t use VBA very long without being exposed to dialog boxes. "),(0,a.kt)("p",null,"They seem to pop up almost every time you select a command. "),(0,a.kt)("p",null,"VBA uses dialog boxes to obtain information, clarify commands, and display messages. "),(0,a.kt)("p",null,"If you develop VBA macros, you can create your own dialog boxes that work just like those built in. "),(0,a.kt)("p",null,"Those custom dialog boxes are called ",(0,a.kt)("inlineCode",{parentName:"p"},"UserForms")," in VBA. About which we look into next section."),(0,a.kt)("h2",{id:"userform-alternatives"},"UserForm Alternatives"),(0,a.kt)("p",null,"Some of the VBA macros you create behave the same every time you execute them. "),(0,a.kt)("p",null,"For example, you may develop a macro for intermediate steps you do every day. "),(0,a.kt)("p",null,"This macro always produces the same result and requires no additional user input."),(0,a.kt)("p",null,"You might develop other macros that behave differently under various circumstances or that offer the user options. "),(0,a.kt)("p",null,"In such cases, the macro may benefit from a custom dialog box. "),(0,a.kt)("p",null,"A custom dialog box provides a simple means for getting information from the user. "),(0,a.kt)("p",null,"Your macro then uses that information to determine what it should do."),(0,a.kt)("p",null,(0,a.kt)("inlineCode",{parentName:"p"},"UserForms")," can be quite useful, but creating them takes time. "),(0,a.kt)("p",null,"Before I cover the topic of creating UserForms in the next section, you need to know about some potentially timesaving alternatives."),(0,a.kt)("p",null,"VBA lets you display several different types of dialog boxes that you can sometimes use in place of a ",(0,a.kt)("inlineCode",{parentName:"p"},"UserForm"),". "),(0,a.kt)("p",null,"You can customize these built-in dialog boxes in some ways, but they certainly don\u2019t offer the options available in a UserForm. "),(0,a.kt)("p",null,"In some cases, however, they\u2019re just what you need."),(0,a.kt)("p",null,"In the following sections you read about"),(0,a.kt)("ul",null,(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"VBA ",(0,a.kt)("inlineCode",{parentName:"p"},"MsgBox")," function")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"VBA ",(0,a.kt)("inlineCode",{parentName:"p"},"InputBox")," function")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"VBA ",(0,a.kt)("inlineCode",{parentName:"p"},"GetOpenFilename")," method")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"VBA ",(0,a.kt)("inlineCode",{parentName:"p"},"GetSaveAsFilename")," method")),(0,a.kt)("li",{parentName:"ul"},(0,a.kt)("p",{parentName:"li"},"VBA ",(0,a.kt)("inlineCode",{parentName:"p"},"FileDialog")," method"))),(0,a.kt)("p",null,"Next post will be about ",(0,a.kt)("strong",{parentName:"p"},(0,a.kt)("em",{parentName:"strong"},"VBA MsgBox Function")),"."))}d.isMDXComponent=!0}}]);