"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[7280],{3905:function(e,t,a){a.d(t,{Zo:function(){return m},kt:function(){return d}});var n=a(67294);function r(e,t,a){return t in e?Object.defineProperty(e,t,{value:a,enumerable:!0,configurable:!0,writable:!0}):e[t]=a,e}function i(e,t){var a=Object.keys(e);if(Object.getOwnPropertySymbols){var n=Object.getOwnPropertySymbols(e);t&&(n=n.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),a.push.apply(a,n)}return a}function o(e){for(var t=1;t<arguments.length;t++){var a=null!=arguments[t]?arguments[t]:{};t%2?i(Object(a),!0).forEach((function(t){r(e,t,a[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(a)):i(Object(a)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(a,t))}))}return e}function l(e,t){if(null==e)return{};var a,n,r=function(e,t){if(null==e)return{};var a,n,r={},i=Object.keys(e);for(n=0;n<i.length;n++)a=i[n],t.indexOf(a)>=0||(r[a]=e[a]);return r}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(n=0;n<i.length;n++)a=i[n],t.indexOf(a)>=0||Object.prototype.propertyIsEnumerable.call(e,a)&&(r[a]=e[a])}return r}var s=n.createContext({}),p=function(e){var t=n.useContext(s),a=t;return e&&(a="function"==typeof e?e(t):o(o({},t),e)),a},m=function(e){var t=p(e.components);return n.createElement(s.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return n.createElement(n.Fragment,{},t)}},c=n.forwardRef((function(e,t){var a=e.components,r=e.mdxType,i=e.originalType,s=e.parentName,m=l(e,["components","mdxType","originalType","parentName"]),c=p(a),d=r,k=c["".concat(s,".").concat(d)]||c[d]||u[d]||i;return a?n.createElement(k,o(o({ref:t},m),{},{components:a})):n.createElement(k,o({ref:t},m))}));function d(e,t){var a=arguments,r=t&&t.mdxType;if("string"==typeof e||r){var i=a.length,o=new Array(i);o[0]=c;var l={};for(var s in t)hasOwnProperty.call(t,s)&&(l[s]=t[s]);l.originalType=e,l.mdxType="string"==typeof e?e:r,o[1]=l;for(var p=2;p<i;p++)o[p]=a[p];return n.createElement.apply(null,o)}return n.createElement.apply(null,a)}c.displayName="MDXCreateElement"},69209:function(e,t,a){a.r(t),a.d(t,{assets:function(){return m},contentTitle:function(){return s},default:function(){return d},frontMatter:function(){return l},metadata:function(){return p},toc:function(){return u}});var n=a(87462),r=a(63366),i=(a(67294),a(3905)),o=["components"],l={title:"VBA Variables",tags:["VBA"],permalink:"/vba/variables/"},s=void 0,p={unversionedId:"vba-variables",id:"vba-variables",title:"VBA Variables",description:"VBA\u2019s main purpose is to manipulate data. VBA stores the data in your computer\u2019s memory; it may or may not end up on disk.",source:"@site/docs/vba/07-vba-variables.md",sourceDirName:".",slug:"/vba-variables",permalink:"/vba/vba-variables",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:7,frontMatter:{title:"VBA Variables",tags:["VBA"],permalink:"/vba/variables/"},sidebar:"tutorialSidebar",previous:{title:"Programming Concepts, Comments and Data-types",permalink:"/vba/vba-programming-concepts"},next:{title:"Declaring and Scoping of Variables",permalink:"/vba/vba-declaring-and-scoping-of-variables"}},m={},u=[],c={toc:u};function d(e){var t=e.components,a=(0,r.Z)(e,o);return(0,i.kt)("wrapper",(0,n.Z)({},c,a,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"VBA\u2019s")," main purpose is to manipulate data. ",(0,i.kt)("inlineCode",{parentName:"p"},"VBA")," stores the ",(0,i.kt)("em",{parentName:"p"},"data")," in your computer\u2019s ",(0,i.kt)("em",{parentName:"p"},"memory"),"; it may or may not end up on disk. "),(0,i.kt)("p",null,"Some ",(0,i.kt)("em",{parentName:"p"},"data"),", such as ",(0,i.kt)("em",{parentName:"p"},"sketch"),", resides in ",(0,i.kt)("inlineCode",{parentName:"p"},"objects"),". "),(0,i.kt)("p",null,"Other ",(0,i.kt)("em",{parentName:"p"},"data")," is stored in ",(0,i.kt)("inlineCode",{parentName:"p"},"variables")," that you create."),(0,i.kt)("p",null,"A ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," is simply a ",(0,i.kt)("em",{parentName:"p"},"named storage location")," in your computer\u2019s memory. "),(0,i.kt)("p",null,"You have lots of flexibility in naming your ",(0,i.kt)("inlineCode",{parentName:"p"},"variables"),", so make the ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," names as descriptive as possible."),(0,i.kt)("p",null,"You assign a value to a ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," by using the equal ",(0,i.kt)("strong",{parentName:"p"},"sign operator"),"."),(0,i.kt)("p",null,"The ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," names in these examples appear on both the left and right sides of the equal signs. "),(0,i.kt)("p",null,"Note that the last example uses two ",(0,i.kt)("inlineCode",{parentName:"p"},"variables"),"."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'x = 1\nInterestRate = 0.075\nLoanPayoffAmount = 243089\nDataEntered = False\nx = x + 1\nUserName = "Bill Gates"\nDateStarted = #3/14/2010#\nMyNum = YourNum * 1.25\n')),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"VBA")," enforces a few rules regarding ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," names:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},"You can use ",(0,i.kt)("em",{parentName:"li"},"letters, numbers, and some punctuation characters"),", but the ",(0,i.kt)("strong",{parentName:"li"},"first character")," must be a letter."),(0,i.kt)("li",{parentName:"ul"},"You ",(0,i.kt)("strong",{parentName:"li"},"cannot")," use any ",(0,i.kt)("em",{parentName:"li"},"spaces or periods")," in a ",(0,i.kt)("inlineCode",{parentName:"li"},"variable")," name."),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"VBA")," does not distinguish between ",(0,i.kt)("em",{parentName:"li"},"uppercase")," and ",(0,i.kt)("em",{parentName:"li"},"lowercase")," letters."),(0,i.kt)("li",{parentName:"ul"},"You ",(0,i.kt)("strong",{parentName:"li"},"cannot")," use the following characters in a variable name: ",(0,i.kt)("strong",{parentName:"li"},"#, $, %, &, or !.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("inlineCode",{parentName:"li"},"Variable")," names can be no longer than ",(0,i.kt)("em",{parentName:"li"},"255")," characters. Of course, you\u2019re only asking for trouble if you use variable names ",(0,i.kt)("em",{parentName:"li"},"255")," characters long.")),(0,i.kt)("p",null,"To make ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," names more ",(0,i.kt)("em",{parentName:"p"},"readable"),", programmers often use mixed case (for example, ",(0,i.kt)("em",{parentName:"p"},"PartDimension"),") or the underscore character (part_dimension)."),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"VBA")," has many ",(0,i.kt)("em",{parentName:"p"},"reserved")," words that you ",(0,i.kt)("strong",{parentName:"p"},"can\u2019t")," use for ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," names or ",(0,i.kt)("inlineCode",{parentName:"p"},"procedure")," names. "),(0,i.kt)("p",null,"These include words such as ",(0,i.kt)("inlineCode",{parentName:"p"},"Sub, Dim, With, End, Next, and For"),". "),(0,i.kt)("p",null,"If you attempt to use one of these words as a ",(0,i.kt)("inlineCode",{parentName:"p"},"variable"),", you may get a compile error (which means your code won\u2019t run. "),(0,i.kt)("p",null,"So, if an assignment statement produces an ",(0,i.kt)("em",{parentName:"p"},"error message"),", double-check and make sure that the ",(0,i.kt)("inlineCode",{parentName:"p"},"variable")," name isn\u2019t a ",(0,i.kt)("em",{parentName:"p"},"reserved")," word."),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"VBA")," does allow you to create ",(0,i.kt)("inlineCode",{parentName:"p"},"variables")," with names that match names in your ",(0,i.kt)("inlineCode",{parentName:"p"},"CAD's object model"),", such as sketch and part. "),(0,i.kt)("p",null,"But, obviously, using ",(0,i.kt)("inlineCode",{parentName:"p"},"names")," like that just increases the possibility of getting confused. "),(0,i.kt)("p",null,"So resist the urge to use a variable named ",(0,i.kt)("em",{parentName:"p"},"sketch"),", and use something like ",(0,i.kt)("em",{parentName:"p"},"swSketch"),", ",(0,i.kt)("em",{parentName:"p"},"mySketch")," or any meaning full name instead."))}d.isMDXComponent=!0}}]);