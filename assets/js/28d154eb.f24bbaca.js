"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[1166],{3905:function(e,t,n){n.d(t,{Zo:function(){return p},kt:function(){return h}});var i=n(67294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);t&&(i=i.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,i)}return n}function o(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function s(e,t){if(null==e)return{};var n,i,a=function(e,t){if(null==e)return{};var n,i,a={},r=Object.keys(e);for(i=0;i<r.length;i++)n=r[i],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(i=0;i<r.length;i++)n=r[i],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var l=i.createContext({}),u=function(e){var t=i.useContext(l),n=t;return e&&(n="function"==typeof e?e(t):o(o({},t),e)),n},p=function(e){var t=u(e.components);return i.createElement(l.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return i.createElement(i.Fragment,{},t)}},c=i.forwardRef((function(e,t){var n=e.components,a=e.mdxType,r=e.originalType,l=e.parentName,p=s(e,["components","mdxType","originalType","parentName"]),c=u(n),h=a,d=c["".concat(l,".").concat(h)]||c[h]||m[h]||r;return n?i.createElement(d,o(o({ref:t},p),{},{components:n})):i.createElement(d,o({ref:t},p))}));function h(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var r=n.length,o=new Array(r);o[0]=c;var s={};for(var l in t)hasOwnProperty.call(t,l)&&(s[l]=t[l]);s.originalType=e,s.mdxType="string"==typeof e?e:a,o[1]=s;for(var u=2;u<r;u++)o[u]=n[u];return i.createElement.apply(null,o)}return i.createElement.apply(null,n)}c.displayName="MDXCreateElement"},19411:function(e,t,n){n.r(t),n.d(t,{assets:function(){return p},contentTitle:function(){return l},default:function(){return h},frontMatter:function(){return s},metadata:function(){return u},toc:function(){return m}});var i=n(87462),a=n(63366),r=(n(67294),n(3905)),o=["components"],s={title:"VBA Functions",tags:["VBA"],permalink:"/vba/functions/"},l=void 0,u={unversionedId:"vba-functions",id:"vba-functions",title:"VBA Functions",description:"A function essentially performs a calculation and returns a single value.",source:"@site/docs/vba/15-vba-functions.md",sourceDirName:".",slug:"/vba-functions",permalink:"/vba/vba-functions",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:15,frontMatter:{title:"VBA Functions",tags:["VBA"],permalink:"/vba/functions/"},sidebar:"tutorialSidebar",previous:{title:"VBA Arrays",permalink:"/vba/vba-arrays"},next:{title:"VBA Functions that do more",permalink:"/vba/vba-more-function"}},p={},m=[{value:"Built-In VBA Functions",id:"built-in-vba-functions",level:2},{value:"Displaying the system date or time",id:"displaying-the-system-date-or-time",level:2},{value:"Finding a string length",id:"finding-a-string-length",level:2},{value:"Displaying the integer part of a number",id:"displaying-the-integer-part-of-a-number",level:2},{value:"Determining a file size",id:"determining-a-file-size",level:2},{value:"Identifying the type of a selected object",id:"identifying-the-type-of-a-selected-object",level:2}],c={toc:m};function h(e){var t=e.components,n=(0,a.Z)(e,o);return(0,r.kt)("wrapper",(0,i.Z)({},c,n,{components:t,mdxType:"MDXLayout"}),(0,r.kt)("p",null,"A ",(0,r.kt)("inlineCode",{parentName:"p"},"function")," essentially performs a calculation and returns a single value. "),(0,r.kt)("p",null,"The ",(0,r.kt)("inlineCode",{parentName:"p"},"SUM")," function in ",(0,r.kt)("strong",{parentName:"p"},"MS Excel")," returns the sum of a range of values. "),(0,r.kt)("p",null,"The same holds true for functions used in your ",(0,r.kt)("strong",{parentName:"p"},"VBA expressions"),": Each function does its thing and returns a single value."),(0,r.kt)("p",null,"The functions you use in VBA can come from two sources:"),(0,r.kt)("ul",null,(0,r.kt)("li",{parentName:"ul"},"Built-in functions provided by VBA"),(0,r.kt)("li",{parentName:"ul"},"Custom functions that you (or someone else) write, using VBA.")),(0,r.kt)("h2",{id:"built-in-vba-functions"},"Built-In VBA Functions"),(0,r.kt)("p",null,"VBA provides numerous ",(0,r.kt)("em",{parentName:"p"},"built-in")," functions. Some of these functions take arguments and some do not."),(0,r.kt)("p",null,"I present a few examples of VBA functions in code. "),(0,r.kt)("p",null,"In many of these examples, I use the ",(0,r.kt)("inlineCode",{parentName:"p"},"MsgBox")," function to display a value in a message box. "),(0,r.kt)("p",null,"Yes, ",(0,r.kt)("inlineCode",{parentName:"p"},"MsgBox")," is a VBA function \u2014 a rather unusual one, but a function nonetheless. "),(0,r.kt)("p",null,"This useful function displays a message in a pop-up dialog box. "),(0,r.kt)("h2",{id:"displaying-the-system-date-or-time"},"Displaying the system date or time"),(0,r.kt)("p",null,"The first example uses VBA\u2019s ",(0,r.kt)("inlineCode",{parentName:"p"},"Date")," function to display the current system date in a message box:"),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowDate()\n  MsgBox Date\nEnd Sub\n")),(0,r.kt)("p",null,"Notice that the ",(0,r.kt)("inlineCode",{parentName:"p"},"Date")," function doesn\u2019t use an argument. "),(0,r.kt)("p",null,"A VBA function with no argument doesn\u2019t require an empty set of parentheses. "),(0,r.kt)("p",null,"In fact, if you type an empty set of parentheses, the VBE will promptly remove them."),(0,r.kt)("p",null,"To get the system time, use the ",(0,r.kt)("inlineCode",{parentName:"p"},"Time")," function. And if you want it all, use the ",(0,r.kt)("inlineCode",{parentName:"p"},"Now")," function to return both the date and the time. "),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowDate()\n  MsgBox Now\nEnd Sub\n")),(0,r.kt)("h2",{id:"finding-a-string-length"},"Finding a string length"),(0,r.kt)("p",null,"The following procedure uses the VBA's ",(0,r.kt)("inlineCode",{parentName:"p"},"Len")," function, which returns the length of a text string. "),(0,r.kt)("p",null,"The ",(0,r.kt)("inlineCode",{parentName:"p"},"Len")," function takes one argument: the ",(0,r.kt)("inlineCode",{parentName:"p"},"string"),". "),(0,r.kt)("p",null,"When you execute this procedure, the ",(0,r.kt)("em",{parentName:"p"},"message box")," displays ",(0,r.kt)("strong",{parentName:"p"},"11")," because the argument has ",(0,r.kt)("strong",{parentName:"p"},"11")," characters. "),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub StringLength()\n  Dim MyString As String\n  Dim StringLength As Integer\n  MyString = \u201cHello World\u201d\n  StringLength = Len(MyString)\n  MsgBox StringLength\nEnd Sub\n")),(0,r.kt)("h2",{id:"displaying-the-integer-part-of-a-number"},"Displaying the integer part of a number"),(0,r.kt)("p",null,"The following procedure uses the ",(0,r.kt)("inlineCode",{parentName:"p"},"Fix")," function, which returns the integer portion of a value \u2014 ",(0,r.kt)("em",{parentName:"p"},"the value without any decimal digits"),": "),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub GetIntegerPart()\n  Dim MyValue As Double\n  Dim IntValue As Integer\n  MyValue = 123.456\n  IntValue = Fix(MyValue)\n  MsgBox IntValue\nEnd Sub\n")),(0,r.kt)("p",null,"In this case, the message box displays ",(0,r.kt)("strong",{parentName:"p"},"123"),"."),(0,r.kt)("p",null,"VBA has a similar function called ",(0,r.kt)("inlineCode",{parentName:"p"},"Int")," Function. "),(0,r.kt)("p",null,"The difference between ",(0,r.kt)("inlineCode",{parentName:"p"},"Int")," and ",(0,r.kt)("inlineCode",{parentName:"p"},"Fix")," is how each deals with negative numbers. "),(0,r.kt)("p",null,"It\u2019s a subtle difference, but sometimes it\u2019s important. "),(0,r.kt)("p",null,(0,r.kt)("inlineCode",{parentName:"p"},"Int")," Function returns the first negative integer that\u2019s less than or equal to the argument. ",(0,r.kt)("inlineCode",{parentName:"p"},"Int(-123.456)")," returns ",(0,r.kt)("strong",{parentName:"p"},"-124"),". "),(0,r.kt)("p",null,(0,r.kt)("inlineCode",{parentName:"p"},"Fix")," Function returns the first negative integer that\u2019s greater than or equal to the argument. ",(0,r.kt)("inlineCode",{parentName:"p"},"Fix(-123.456)")," returns ",(0,r.kt)("strong",{parentName:"p"},"-123"),". "),(0,r.kt)("h2",{id:"determining-a-file-size"},"Determining a file size"),(0,r.kt)("p",null,"The following ",(0,r.kt)("inlineCode",{parentName:"p"},"Sub")," procedure displays the size, in bytes, of the executable file. "),(0,r.kt)("p",null,"It finds this value by using the ",(0,r.kt)("inlineCode",{parentName:"p"},"FileLen")," function. "),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub GetFileSize()\n  Dim TheFile As String\n  TheFile \u201cC:\\ProgramFiles\\Program File\\SolidworksCorp\\SLDWORKS\\SLDWORKS.exe\u201d\n  MsgBox FileLen(TheFile)\nEnd Sub\n")),(0,r.kt)("p",null,"Notice that this routine hard codes the filename (that is, it explicitly states the path). "),(0,r.kt)("p",null,"Generally, this ",(0,r.kt)("strong",{parentName:"p"},"isn\u2019t")," a good idea. The file might not be on the ",(0,r.kt)("em",{parentName:"p"},"C drive"),", or the Program File folder may have a different location. "),(0,r.kt)("p",null,"The following statement shows a better approach: "),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers",showlinenumbers:!0},"TheFile = Application.Path & \u201c\\SLDWORKS.EXE\u201d \n")),(0,r.kt)("p",null,"Path is a property of the Application object. "),(0,r.kt)("p",null,"It simply returns the name of the folder in which the application (that is, ",(0,r.kt)("em",{parentName:"p"},"Solidworks"),") is installed (without a trailing backslash). "),(0,r.kt)("h2",{id:"identifying-the-type-of-a-selected-object"},"Identifying the type of a selected object"),(0,r.kt)("p",null,"The following procedure uses the ",(0,r.kt)("inlineCode",{parentName:"p"},"TypeName")," function, which returns the type of the selection (as a ",(0,r.kt)("inlineCode",{parentName:"p"},"string"),"): "),(0,r.kt)("pre",null,(0,r.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"Sub ShowSelectionType()\n  Dim SelType As String\n  SelType = TypeName(Selection)\n  MsgBox SelType\nEnd Sub\n")),(0,r.kt)("p",null,"This could be ",(0,r.kt)("em",{parentName:"p"},"a Sketch, a Part, a Assembly")," or any ",(0,r.kt)("em",{parentName:"p"},"other type")," of object that can be selected."),(0,r.kt)("p",null,"The ",(0,r.kt)("inlineCode",{parentName:"p"},"TypeName")," function is very versatile. You can also use this function to determine the data type of a variable. "),(0,r.kt)("p",null,"Next post will be about ",(0,r.kt)("strong",{parentName:"p"},(0,r.kt)("em",{parentName:"strong"},"VBA Functions that do more")),"."))}h.isMDXComponent=!0}}]);